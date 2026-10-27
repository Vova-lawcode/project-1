from __future__ import annotations
import logging
from uuid import UUID, uuid4
from pathlib import Path
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, ConfigDict, Field
from dotenv import load_dotenv

from services.schema import CaseData, HistoryEntry
from services.llm import LLMService, LLMError
from services.docx_generator import generate_order_cancellation
from services.mock_case import demo_case
from services.prompts import ALL_REQUIRED

BASE = Path(__file__).parent
GEN = BASE / 'generated'
GEN.mkdir(exist_ok=True)
load_dotenv(BASE / '.env')

app = FastAPI(title='Иск-Бот AI MVP', version='0.3.0')
app.mount('/static', StaticFiles(directory=BASE/'static'), name='static')
llm = LLMService()
logger = logging.getLogger(__name__)

class TurnRequest(BaseModel):
    model_config = ConfigDict(extra='forbid')

    history: list[HistoryEntry] = Field(default_factory=list, max_length=12)
    case: CaseData = Field(default_factory=CaseData)
    user_text: str = Field(default='', max_length=2_000)


def _new_document_path() -> tuple[str, Path]:
    document_id = str(uuid4())
    return document_id, GEN / f'{document_id}.docx'

@app.get('/', response_class=HTMLResponse)
def index():
    return (BASE/'templates'/'index.html').read_text(encoding='utf-8')

@app.get('/health')
def health():
    return {'ok': True, 'provider': llm.provider, 'model': llm.model}

@app.post('/api/turn')
def turn(req: TurnRequest):
    try:
        result = llm.next_turn([entry.model_dump() for entry in req.history], req.case, req.user_text)
        merged = req.case.model_copy(update=result.fields)
        return {'ok': True, 'turn': result.model_dump(), 'case': merged.model_dump()}
    except LLMError as exc:
        logger.warning('AI provider failed: %s', exc, exc_info=True)
        raise HTTPException(502, 'Не удалось обработать ответ AI. Попробуйте ещё раз позже.') from exc

@app.post('/api/demo')
def demo():
    case = demo_case()
    document_id, out = _new_document_path()
    generate_order_cancellation(case, out)
    return {'ok': True, 'download_url': f'/api/download/{document_id}', 'case': case.model_dump()}

@app.post('/api/generate')
def generate(case: CaseData):
    missing = [field for field in ALL_REQUIRED if not getattr(case, field)]
    if missing:
        raise HTTPException(400, f'Не заполнены обязательные сведения: {", ".join(missing)}')
    document_id, out = _new_document_path()
    try:
        generate_order_cancellation(case, out)
    except Exception as exc:
        logger.exception('Document generation failed')
        raise HTTPException(500, 'Не удалось сформировать документ. Проверьте данные и попробуйте ещё раз.') from exc
    return {'ok': True, 'download_url': f'/api/download/{document_id}'}

@app.get('/api/download/{document_id}')
def download_document(document_id: str):
    try:
        safe_id = str(UUID(document_id))
    except ValueError as exc:
        raise HTTPException(404, 'Документ не найден') from exc
    path = GEN / f'{safe_id}.docx'
    if not path.is_file() or path.parent != GEN:
        raise HTTPException(404, 'Документ не найден')
    return FileResponse(path, filename='zayavlenie_ob_otmene_sudebnogo_prikaza.docx', media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
