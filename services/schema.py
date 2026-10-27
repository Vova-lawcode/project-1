from __future__ import annotations
from typing import Annotated, Dict, Any, Literal
from pydantic import BaseModel, ConfigDict, Field

ShortText = Annotated[str, Field(max_length=500)]
LongText = Annotated[str, Field(max_length=4_000)]
HistoryText = Annotated[str, Field(max_length=2_000)]
DocumentTitle = Annotated[str, Field(max_length=300)]

class CaseData(BaseModel):
    model_config = ConfigDict(extra='forbid')

    court_number: ShortText = ''
    court_district: ShortText = ''
    court_address: ShortText = ''
    applicant_name: ShortText = ''
    birth_date: ShortText = ''
    passport: ShortText = ''
    passport_issued: ShortText = ''
    inn: ShortText = ''
    snils: ShortText = ''
    applicant_address: ShortText = ''
    phone: ShortText = ''
    email: ShortText = ''
    creditor_name: ShortText = ''
    creditor_inn: ShortText = ''
    creditor_address: ShortText = ''
    case_number: ShortText = ''
    order_date: ShortText = ''
    order_subject: LongText = ''
    ip_date: ShortText = ''
    ip_number: ShortText = ''
    awareness_date: ShortText = ''
    awareness_source: LongText = ''
    missed_reason: LongText = ''
    supporting_documents: list[DocumentTitle] = Field(default_factory=list, max_length=50)
    disagreement_basis: LongText = ''
    effective_start_date: ShortText = ''
    fssp_office: ShortText = ''
    correspondence_address: ShortText = ''
    composition_date: ShortText = ''
    source_facts: list[HistoryText] = Field(default_factory=list, max_length=50)

class AssistantTurn(BaseModel):
    assistant_message: Annotated[str, Field(max_length=4_000)]
    fields: Dict[str, Any] = Field(default_factory=dict, max_length=40)
    ready_to_generate: bool = False
    missing_fields: list[str] = Field(default_factory=list)
    current_field: Annotated[str, Field(max_length=100)] = ''
    validation_error: Annotated[str, Field(max_length=1_000)] = ''


class HistoryEntry(BaseModel):
    model_config = ConfigDict(extra='forbid')

    role: Literal['user', 'assistant']
    content: HistoryText
