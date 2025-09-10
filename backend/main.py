from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Any, Dict, List, Optional

from .extract import parse_docx
from .state import get_state, set_state, save_docs, list_docs, delete_doc, list_filenames

app = FastAPI(title="DOCX Parser API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class ParseResponse(BaseModel):
    highlights: Any
    comments: Any
    paragraphs: Any


@app.post("/api/parse", response_model=ParseResponse)
async def parse(file: UploadFile = File(...)) -> Dict[str, Any]:
    data = await file.read()
    parsed = parse_docx(data, file.filename)
    return parsed


@app.post("/api/parse-multi", response_model=ParseResponse)
async def parse_multi(files: List[UploadFile] = File(...)) -> Dict[str, Any]:
    all_highlights: List[Dict[str, Any]] = []
    all_comments: List[Dict[str, Any]] = []
    all_paragraphs: List[Dict[str, Any]] = []
    for f in files:
        content = await f.read()
        parsed = parse_docx(content, f.filename)
        all_highlights.extend(parsed.get("highlights", []))
        all_comments.extend(parsed.get("comments", []))
        all_paragraphs.extend(parsed.get("paragraphs", []))
    return {"highlights": all_highlights, "comments": all_comments, "paragraphs": all_paragraphs}


@app.post("/api/upload-multi", response_model=ParseResponse)
async def upload_multi(files: List[UploadFile] = File(...)) -> Dict[str, Any]:
    """Parse and persist multiple files, returning the aggregated dataset including previous docs."""
    for f in files:
        content = await f.read()
        parsed = parse_docx(content, f.filename)
        save_docs({f.filename: parsed})
    return list_docs()


@app.get("/api/docs", response_model=ParseResponse)
def get_docs() -> Dict[str, Any]:
    return list_docs()


@app.delete("/api/docs/{filename}")
def remove_doc(filename: str) -> Dict[str, Any]:
    delete_doc(filename)
    return {"filenames": list_filenames(), **list_docs()}


@app.get("/api/state")
def read_state() -> Dict[str, Any]:
    return get_state()


class StatePatch(BaseModel):
    colorMap: Optional[Any] = None
    categoryColors: Optional[Any] = None
    catOverride: Optional[Any] = None
    meta: Optional[Any] = None
    extraCategories: Optional[Any] = None


@app.post("/api/state")
def write_state(patch: StatePatch) -> Dict[str, Any]:
    payload: Dict[str, Any] = {k: v for k, v in patch.dict().items() if v is not None}
    return set_state(payload)
