from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from uuid import uuid4

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from scripts.build_qingcha_table import (
    BuildDiagnostics,
    MissingSourceFilesError,
    OUTPUT_NAME,
    TEMPLATE_NAME,
    build_workbook,
    required_source_filenames,
)

BASE_DIR = Path(__file__).resolve().parents[2]
BUNDLED_TEMPLATE_PATH = BASE_DIR / "backend" / "templates" / TEMPLATE_NAME
TEMPLATE_PATH = Path("/app/templates") / TEMPLATE_NAME
if not TEMPLATE_PATH.exists():
    TEMPLATE_PATH = BUNDLED_TEMPLATE_PATH

FRONTEND_DIR = BASE_DIR / "frontend"

JOB_ROOT = Path(tempfile.gettempdir()) / "qingcha-web-jobs"
JOB_ROOT.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="清查表生成工具")


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/api/required-files")
def required_files() -> dict[str, object]:
    return {
        "templateProvidedByServer": True,
        "templateName": TEMPLATE_NAME,
        "requiredFiles": required_source_filenames(),
    }


def match_required_filename(filename: str) -> str | None:
    """按“必需文件名去掉扩展名作为前缀”识别上传文件。"""
    if not filename.endswith(".xlsx"):
        return None
    for required in required_source_filenames():
        required_stem = required.removesuffix(".xlsx")
        if filename.startswith(required_stem):
            return required
    return None


def normalize_output_filename(filename: str | None) -> str:
    """规范前端传入的下载文件名，避免路径穿越并补齐 xlsx 后缀。"""
    cleaned = Path((filename or "").strip()).name
    if not cleaned:
        return OUTPUT_NAME
    if not cleaned.endswith(".xlsx"):
        cleaned = f"{cleaned}.xlsx"
    return cleaned


@app.post("/api/build")
async def build(
    files: list[UploadFile] = File(...),
    output_filename: str | None = Form(default=None),
) -> JSONResponse:
    job_id = uuid4().hex
    job_dir = JOB_ROOT / job_id
    upload_dir = job_dir / "uploads"
    output_dir = job_dir / "outputs"
    upload_dir.mkdir(parents=True)
    output_dir.mkdir()

    diagnostics = BuildDiagnostics.create()
    final_filename = normalize_output_filename(output_filename)
    seen_required: set[str] = set()

    for uploaded in files:
        filename = Path(uploaded.filename or "").name
        if not filename:
            diagnostics.warning("忽略了没有文件名的上传项。")
            continue
        if filename == TEMPLATE_NAME:
            diagnostics.warning(f"忽略上传的模板文件：{filename}。系统使用服务端固定模板。")
            continue
        matched_filename = match_required_filename(filename)
        if matched_filename is None:
            diagnostics.warning(f"忽略非必需上传文件：{filename}")
            continue

        destination = upload_dir / matched_filename
        with destination.open("wb") as out_file:
            shutil.copyfileobj(uploaded.file, out_file)
        seen_required.add(matched_filename)
        if filename == matched_filename:
            diagnostics.info(f"已接收上传文件：{filename}")
        else:
            diagnostics.warning(f"已将上传文件 {filename} 识别为 {matched_filename}")

    missing_before_build = [filename for filename in required_source_filenames() if filename not in seen_required]
    if missing_before_build:
        diagnostics.missing_files.extend(missing_before_build)
        for filename in missing_before_build:
            diagnostics.warning(f"缺少上传文件：{filename}")
        return JSONResponse(
            status_code=400,
            content={
                "error": "missing_required_files",
                "missingFiles": missing_before_build,
                "logs": diagnostics.logs,
            },
        )

    output_path = output_dir / final_filename
    try:
        build_workbook(upload_dir, output_path, template_path=TEMPLATE_PATH, diagnostics=diagnostics)
    except MissingSourceFilesError as exc:
        return JSONResponse(
            status_code=400,
            content={
                "error": "missing_required_files",
                "missingFiles": exc.missing_files,
                "logs": diagnostics.logs,
            },
        )
    except Exception as exc:  # noqa: BLE001 - API 需要把构建异常转成前端可展示日志。
        diagnostics.warning(f"构建失败：{exc}")
        return JSONResponse(
            status_code=422,
            content={"error": "build_failed", "logs": diagnostics.logs},
        )

    return JSONResponse(
        content={
            "jobId": job_id,
            "filename": final_filename,
            "downloadUrl": f"/api/download/{job_id}",
            "logs": diagnostics.logs,
        }
    )


@app.get("/api/download/{job_id}")
def download(job_id: str) -> FileResponse:
    if not job_id.isalnum():
        raise HTTPException(status_code=404, detail="job_not_found")
    output_dir = JOB_ROOT / job_id / "outputs"
    if not output_dir.is_dir():
        raise HTTPException(status_code=404, detail="job_not_found")
    output_files = sorted(output_dir.glob("*.xlsx"))
    if not output_files:
        raise HTTPException(status_code=404, detail="job_not_found")
    output_path = output_files[0]
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_path.name,
    )


if FRONTEND_DIR.exists():
    app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")
