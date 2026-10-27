from pathlib import Path
import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm

from services.mock_case import demo_case
from services.docx_generator import generate_order_cancellation, _normalize_court_district, _normalize_order_subject
from services.llm import LLMService
from services.schema import CaseData


@pytest.fixture(autouse=True)
def use_mock_provider(monkeypatch):
    monkeypatch.setenv('LLM_PROVIDER', 'mock')


def _all_text(doc):
    texts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texts.append(cell.text)
    return '\n'.join(texts)


def test_generate_demo_contains_passport_and_polished_court_name(tmp_path: Path):
    out = tmp_path / 'demo.docx'
    generate_order_cancellation(demo_case(), out)
    doc = Document(out)
    text = _all_text(doc)
    assert 'МИРОВОМУ СУДЬЕ' in text
    assert 'Паспорт: 69 12 345678' in text
    assert 'Выдан: УМВД России по Томской области 20.04.2015' in text
    assert 'Дата рождения: 15.03.1989' in text
    assert 'Дело № 2-1234/2026' in text
    assert 'Октябрьского судебного района города Томска' in text
    assert 'Приложения (копии):' in text
    assert 'Дата: 30.09.2026' in text
    assert '________________ / Иванов Иван Иванович' in text
    assert '[НАИМЕНОВАНИЕ ОСП' not in text


def test_formatting_is_applied_to_body_header_right_aligned_and_request_heading_centered(tmp_path: Path):
    out = tmp_path / 'format.docx'
    case = demo_case().model_copy(update={'court_district': 'Кировский район города Томска'})
    generate_order_cancellation(case, out)
    doc = Document(out)
    assert doc.sections[0].page_width.cm == pytest.approx(21.0, abs=0.1)
    assert doc.styles['Normal'].font.name == 'Times New Roman'
    court_line = next(p for p in doc.paragraphs if p.text == 'Кировского судебного района города Томска')
    assert court_line.alignment == WD_ALIGN_PARAGRAPH.RIGHT
    request_heading = next(p for p in doc.paragraphs if p.text == 'ПРОШУ:')
    assert request_heading.alignment == WD_ALIGN_PARAGRAPH.CENTER
    body = next(p for p in doc.paragraphs if p.text.startswith('12.08.2026 в отношении меня'))
    assert body.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
    assert body.paragraph_format.line_spacing == pytest.approx(1.5)
    assert body.paragraph_format.first_line_indent.cm == pytest.approx(1.25, abs=0.01)
    assert all((run.font.name or 'Times New Roman') == 'Times New Roman' for p in doc.paragraphs for run in p.runs)


def test_court_and_subject_normalization():
    assert _normalize_court_district('Кировский район города Томска') == 'Кировского судебного района города Томска'
    assert _normalize_court_district('Кировского судебного района города Томска') == 'Кировского судебного района города Томска'
    assert _normalize_court_district('Советский район города Казани') == 'Советского судебного района города Казани'
    assert _normalize_court_district('Центральный район г. Новосибирска') == 'Центрального судебного района г. Новосибирска'
    assert _normalize_order_subject('взыскание задолженности по коммунальным услугам') == 'взыскании задолженности по коммунальным услугам'


def test_confirmed_llm_fields_are_not_overwritten(monkeypatch):
    from services import llm as llm_module

    class FakeCompletions:
        def create(self, **kwargs):
            class Message:
                content = '{"assistant_message":"Продолжим","fields":{"applicant_name":"Подменённое Имя","birth_date":"01.01.2000"}}'

            class Choice:
                message = Message()

            class Response:
                choices = [Choice()]

            return Response()

    class FakeClient:
        class chat:
            class completions:
                create = FakeCompletions().create

    monkeypatch.setattr(llm_module, 'OpenAI', FakeClient, raising=False)
    service = LLMService.__new__(LLMService)
    service.provider = 'openai'
    service.client = FakeClient()
    service.model = 'test'
    result = service.next_turn([], CaseData(applicant_name='Иванов Иван Иванович'), '')
    assert 'applicant_name' not in result.fields
    assert result.fields['birth_date'] == '01.01.2000'


def test_case_schema_rejects_oversized_input():
    from pydantic import ValidationError

    with pytest.raises(ValidationError):
        CaseData(applicant_name='x' * 501)


def test_each_generated_document_gets_a_separate_download_id(tmp_path: Path, monkeypatch):
    from fastapi.testclient import TestClient
    import app as app_module

    monkeypatch.setattr(app_module, 'GEN', tmp_path)
    client = TestClient(app_module.app)
    case = demo_case().model_dump()

    first = client.post('/api/generate', json=case)
    second = client.post('/api/generate', json=case)

    assert first.status_code == 200
    assert second.status_code == 200
    first_url = first.json()['download_url']
    second_url = second.json()['download_url']
    assert first_url != second_url
    assert client.get(first_url).status_code == 200
    assert client.get(second_url).status_code == 200
    assert client.get('/api/download/../../app.py').status_code == 404


def test_initial_free_text_is_not_rejected_or_put_into_court_number():
    svc = LLMService()
    case = CaseData()
    result = svc.next_turn([], case, 'Коммунальщики взыскали 15000 рублей долга за коммунальные услуги.')
    assert result.fields.get('court_number') in (None, '')
    assert result.current_field == 'court_number'
    assert result.fields.get('source_facts') == ['Коммунальщики взыскали 15000 рублей долга за коммунальные услуги.']
    assert 'записал ваше описание' in result.assistant_message.lower()


def test_initial_free_text_with_browser_history_is_not_rejected():
    svc = LLMService()
    case = CaseData()
    history = [{'role': 'user', 'content': 'долг по коммуналке 15000 рублей'}]
    result = svc.next_turn(history, case, 'долг по коммуналке 15000 рублей')
    assert result.fields.get('court_number') in (None, '')
    assert result.current_field == 'court_number'
    assert result.fields.get('source_facts') == ['долг по коммуналке 15000 рублей']
    assert 'записал ваше описание' in result.assistant_message.lower()


def test_guided_answers_reach_passport_and_beyond():
    svc = LLMService(); case = CaseData()
    answers = [
        ('12', 'court_number'),
        ('Октябрьского судебного района города Томска', 'court_district'),
        ('634050, г. Томск, ул. Ленина, д. 12', 'court_address'),
        ('Иванов Иван Иванович', 'applicant_name'),
        ('15.03.1989', 'birth_date'),
        ('69 12 345678', 'passport'),
        ('УМВД России по Томской области 20.04.2015', 'passport_issued'),
        ('634041, г. Томск, ул. Пушкина, д. 27, кв. 14', 'applicant_address'),
        ('ООО «Вектор»', 'creditor_name'),
        ('2-1234/2026', 'case_number'),
        ('12.08.2026', 'order_date'),
        ('взыскании задолженности по договору займа в размере 500 000 рублей', 'order_subject'),
        ('18.09.2026', 'awareness_date'),
        ('из уведомления в Госуслугах', 'awareness_source'),
        ('Копию судебного приказа я не получал, о нем узнал только из Госуслуг.', 'missed_reason'),
        ('с наличием задолженности', 'disagreement_basis'),
        ('18.09.2026', 'effective_start_date'),
        ('ОСП по Октябрьскому району г. Томска УФССП России по Томской области', 'fssp_office'),
    ]
    history = []
    for answer, field in answers:
        turn = svc.next_turn(history, case, answer)
        assert turn.fields.get(field) == answer or field in turn.fields
        case = case.model_copy(update=turn.fields)
        history.extend([{'role': 'user', 'content': answer}, {'role': 'assistant', 'content': turn.assistant_message}])
    assert all(getattr(case, k) for k in ['birth_date', 'passport', 'passport_issued'])

