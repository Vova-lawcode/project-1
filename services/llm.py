from __future__ import annotations
import json
import os
import re
from datetime import datetime
from typing import Any
from pydantic import ValidationError
from .prompts import SYSTEM_PROMPT, ALL_REQUIRED, FIELD_LABELS
from .schema import CaseData, AssistantTurn

class LLMError(RuntimeError):
    pass

DATE_RE = re.compile(r'\b\d{1,2}[./-]\d{1,2}[./-]\d{4}\b')
PASSPORT_RE = re.compile(r'\b\d{2}\s?\d{2}\s?\d{6}\b')
INN_RE = re.compile(r'\b\d{10}(?:\d{2})?\b')
SNILS_RE = re.compile(r'\b\d{3}[- ]?\d{3}[- ]?\d{3}[ -]?\d{2}\b')
PHONE_RE = re.compile(r'(?:\+7|8)[\s()-]?\d{3}[\s()-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}')
EMAIL_RE = re.compile(r'\b[^\s@]+@[^\s@]+\.[^\s@]+\b')

class LLMService:
    def __init__(self):
        self.provider = os.getenv('LLM_PROVIDER', 'mock').lower()
        if self.provider == 'deepseek':
            from openai import OpenAI
            self.client = OpenAI(api_key=os.getenv('DEEPSEEK_API_KEY'), base_url='https://api.deepseek.com')
            self.model = os.getenv('DEEPSEEK_MODEL', 'deepseek-flash')
        elif self.provider == 'openai':
            from openai import OpenAI
            self.client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))
            self.model = os.getenv('OPENAI_MODEL', 'gpt-5.6-luna')
        else:
            self.client = None
            self.model = 'mock'

    def next_turn(self, history: list[dict[str, str]], case: CaseData, user_text: str) -> AssistantTurn:
        if self.provider == 'mock':
            return self._mock(history, case, user_text)
        if not self.client:
            raise LLMError('Не задан API-ключ для выбранного провайдера.')
        payload = {
            'case': case.model_dump(),
            'required_fields': {k: FIELD_LABELS[k] for k in ALL_REQUIRED},
            'history': history[-12:],
            'user_message': user_text,
        }
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {'role': 'system', 'content': SYSTEM_PROMPT},
                    {'role': 'user', 'content': json.dumps(payload, ensure_ascii=False)},
                ],
                response_format={'type': 'json_object'},
                temperature=0.1,
            )
            raw = response.choices[0].message.content or '{}'
            obj = json.loads(raw)
            turn = AssistantTurn.model_validate(obj)
            # Treat model output as untrusted input and validate each field independently.
            safe_fields = {}
            for key, value in turn.fields.items():
                if key not in CaseData.model_fields or getattr(case, key) or value is None:
                    continue
                try:
                    validated = CaseData.model_validate({key: value})
                except ValidationError:
                    continue
                normalized_value = getattr(validated, key)
                if not isinstance(normalized_value, str) or normalized_value.strip():
                    safe_fields[key] = normalized_value
            return turn.model_copy(update={'fields': safe_fields})
        except Exception as exc:
            raise LLMError(f'Ошибка AI-провайдера: {exc}') from exc

    def _mock(self, history: list[dict[str, str]], case: CaseData, user_text: str) -> AssistantTurn:
        text = user_text.strip()
        missing = [k for k in ALL_REQUIRED if not getattr(case, k)]
        if not text:
            return self._ask(case, missing[0] if missing else '')

        # На первом пользовательском сообщении человек волен описывать ситуацию как угодно.
        # Веб-интерфейс добавляет текущее сообщение пользователя в history ДО вызова API,
        # поэтому нельзя определять первый ход только по len(history) == 0.
        prior_assistant_messages = [h for h in history if h.get('role') == 'assistant']
        is_first_message = (
            not any(case.model_dump().values())
            and len(prior_assistant_messages) == 0
        )
        updates = self._extract_many(text, case)

        # Если первое сообщение похоже на прямой ответ на первый вопрос (например,
        # пользователь сразу написал «12»), принимаем его как номер участка.
        # Иначе считаем сообщение свободным рассказом о ситуации.
        if is_first_message and not updates:
            direct_first = self._accept_for_field('court_number', text)
            if direct_first is not None:
                updates['court_number'] = direct_first
                is_first_message = False

        # Сохраняем исходные факты, даже если mock-режим пока не умеет разложить
        # их по отдельным полям. В режиме с настоящей LLM именно эти факты будут
        # структурироваться моделью.
        if is_first_message:
            facts = list(case.source_facts)
            if text not in facts:
                facts.append(text)
            updates['source_facts'] = facts

            merged = case.model_copy(update=updates)
            missing_after = [k for k in ALL_REQUIRED if not getattr(merged, k)]
            if missing_after:
                field = missing_after[0]
                accepted = 'Я записал ваше описание ситуации.'
                if updates and any(k != 'source_facts' for k in updates):
                    recognized = [FIELD_LABELS[k] for k in updates if k in FIELD_LABELS and k != 'source_facts']
                    if recognized:
                        accepted += ' Также распознал: ' + ', '.join(recognized) + '.'
                msg = accepted + '\nТеперь укажите ' + FIELD_LABELS[field] + '.'
                return AssistantTurn(
                    assistant_message=msg,
                    fields=updates,
                    missing_fields=missing_after,
                    current_field=field,
                    ready_to_generate=False,
                )

        # В последующих сообщениях пытаемся извлечь несколько полей. Если ничего
        # не извлеклось, принимаем ответ только в соответствии с текущим вопросом.
        if not updates:
            if missing:
                field = missing[0]
                value = self._accept_for_field(field, text)
                if value is not None:
                    updates[field] = value
                else:
                    return self._ask(case, field, f'Не удалось распознать {FIELD_LABELS[field]}.')

        merged = case.model_copy(update=updates)
        missing_after = [k for k in ALL_REQUIRED if not getattr(merged, k)]
        if missing_after:
            field = missing_after[0]
            msg = 'Принял данные. Теперь укажите ' + FIELD_LABELS[field] + '.'
            return AssistantTurn(
                assistant_message=msg,
                fields=updates,
                missing_fields=missing_after,
                current_field=field,
                ready_to_generate=False,
            )
        return AssistantTurn(
            assistant_message='Основные данные собраны. Проверьте сводку ниже. Если всё верно, нажмите «Сформировать заявление».',
            fields=updates,
            missing_fields=[],
            current_field='',
            ready_to_generate=True,
        )

    def _ask(self, case: CaseData, field: str, error: str = '') -> AssistantTurn:
        missing = [k for k in ALL_REQUIRED if not getattr(case, k)]
        msg = (error + ' ' if error else '') + f'Укажите {FIELD_LABELS[field]}.'
        return AssistantTurn(assistant_message=msg, missing_fields=missing, current_field=field, validation_error=error)

    def _accept_for_field(self, field: str, text: str) -> str | None:
        if field == 'court_number':
            cleaned = text.strip()
            if re.fullmatch(r'\d{1,4}', cleaned):
                return cleaned
            m = re.fullmatch(r'(?:судебного\s+)?участка\s*№?\s*(\d{1,4})', cleaned, re.I)
            return m.group(1) if m else None
        if field in {'court_district', 'court_address', 'applicant_address', 'creditor_address', 'fssp_office', 'passport_issued'}:
            return text if len(text) >= 5 else None
        if field == 'applicant_name':
            return text if self._looks_like_name(text) else None
        if field == 'birth_date' or field == 'order_date' or field == 'awareness_date' or field == 'effective_start_date':
            return self._first_date(text)
        if field == 'passport':
            m = PASSPORT_RE.search(text)
            return self._normalize_spaces(m.group(0)) if m else None
        if field == 'creditor_name':
            return text if len(text) >= 2 else None
        if field in {'case_number', 'order_subject', 'awareness_source', 'missed_reason', 'disagreement_basis'}:
            return text if len(text) >= 3 else None
        return text if text else None

    def _extract_many(self, text: str, case: CaseData) -> dict[str, Any]:
        updates: dict[str, Any] = {}
        # Dates
        dates = DATE_RE.findall(text)
        if dates:
            norm = [self._normalize_date(x) for x in dates]
            if not case.birth_date and re.search(r'рожд|родил', text, re.I): updates['birth_date'] = norm[0]
            if not case.order_date and re.search(r'приказ(?:а|е)?[^.]{0,40}(?:от|дата)|вынесен[^.]{0,40}\d', text, re.I): updates['order_date'] = norm[0]
            if not case.awareness_date and re.search(r'узнал|известно|уведом', text, re.I): updates['awareness_date'] = norm[-1]
        # Passport / date issued
        pm = PASSPORT_RE.search(text)
        if pm and not case.passport:
            updates['passport'] = self._normalize_spaces(pm.group(0))
        if 'паспорт' in text.lower() and not case.passport_issued:
            m = re.search(r'(?:выдан|выда?н)[\s:—-]*(.+?)(?=\.|$)', text, re.I)
            if m: updates['passport_issued'] = m.group(1).strip()
        # INN/SNILS/phone/email
        if not case.inn:
            m = INN_RE.search(text)
            if m and len(m.group(0)) in (10,12): updates['inn'] = m.group(0)
        if not case.snils:
            m = SNILS_RE.search(text)
            if m: updates['snils'] = m.group(0)
        if not case.phone:
            m = PHONE_RE.search(text)
            if m: updates['phone'] = m.group(0)
        if not case.email:
            m = EMAIL_RE.search(text)
            if m: updates['email'] = m.group(0)
        # Applicant name: explicit marker or a clean 3-word FIO line
        if not case.applicant_name:
            m = re.search(r'(?:заявитель|мое\s+фио|фио)[\s:—-]+([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)', text)
            if m: updates['applicant_name'] = m.group(1).strip()
            elif self._looks_like_name(text): updates['applicant_name'] = text
        # Court number / district / address
        if not case.court_number:
            m = re.search(r'(?:судебн(?:ого|ый)\s+)?участк(?:а|у|е)\s*№?\s*(\d{1,4})', text, re.I)
            if m: updates['court_number'] = m.group(1)
        if not case.court_district and re.search(r'судебн(?:ого|ый)\s+район|района', text, re.I):
            m = re.search(
                r'((?:[А-ЯЁ][а-яё-]*\s+){1,3}(?:судебн(?:ого|ый)\s+)?район\s+(?:города|г\.)\s+[А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё-]+)?)',
                text,
                re.I,
            )
            if m: updates['court_district'] = m.group(1).strip()
        if not case.court_address and re.search(r'адрес\s+суда|г\.\s*[А-ЯЁ][а-яё]+.*(?:ул\.|просп|д\.|дом|шоссе)', text, re.I):
            m = re.search(r'(?:адрес\s+суда[\s:—-]*)?((?:г\.\s*)?[А-ЯЁ][^,;\n]+(?:,\s*(?:ул\.|улица|проспект|просп\.|пер\.|д\.|дом|шоссе)[^\n]+))', text, re.I)
            if m: updates['court_address'] = m.group(1).strip()
        # Applicant address
        if not case.applicant_address:
            m = re.search(r'(?:адрес|проживаю|проживание)[\s:—-]+(.+)', text, re.I)
            if m and ('ул.' in m.group(1).lower() or 'д.' in m.group(1).lower() or 'г.' in m.group(1).lower()):
                updates['applicant_address'] = m.group(1).strip().rstrip('.')
        # Creditor
        if not case.creditor_name:
            m = re.search(r'(?:взыскатель|кредитор)[\s:—-]+(.+?)(?=,?\s*(?:ИНН|адрес|номер\s+дела|$))', text, re.I)
            if m: updates['creditor_name'] = m.group(1).strip()
        # Case number
        if not case.case_number:
            m = re.search(r'дело\s*№\s*([A-Za-zА-Яа-я0-9./-]+)', text, re.I)
            if m: updates['case_number'] = m.group(1)
        # Order subject
        if not case.order_subject:
            m = re.search(r'(?:приказ\s+(?:о|об)\s+|взыскани[ея]\s+)(.+?)(?=\.|$)', text, re.I)
            if m and len(m.group(1)) > 5: updates['order_subject'] = m.group(1).strip().rstrip('.')
        # Awareness source
        if not case.awareness_source and re.search(r'госуслуг|госпочт|пристав|постановлен', text, re.I):
            updates['awareness_source'] = text
        # Disagreement
        if not case.disagreement_basis and re.search(r'не согласен|не согласна|оспарива', text, re.I):
            updates['disagreement_basis'] = text
        return updates

    @staticmethod
    def _first_date(text: str) -> str | None:
        m = DATE_RE.search(text)
        return LLMService._normalize_date(m.group(0)) if m else None

    @staticmethod
    def _normalize_date(value: str) -> str:
        try:
            return datetime.strptime(value.replace('/', '.').replace('-', '.'), '%d.%m.%Y').strftime('%d.%m.%Y')
        except ValueError:
            return value

    @staticmethod
    def _normalize_spaces(value: str) -> str:
        return re.sub(r'\s+', ' ', value).strip()

    @staticmethod
    def _looks_like_name(text: str) -> bool:
        return bool(re.fullmatch(r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+', text.strip()))
