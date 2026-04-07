import json

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import StreamingResponse

from src.performance_planner import update_project_schedule_stream

app = FastAPI(title="Project Scheduler API")


@app.get("/")
def home():
    return {"message": "API is running"}


@app.post("/schedule")
async def generate_schedule(
    file: UploadFile = File(...),
    project_start_date: str = Form(...),
    holidays: str = Form("[]"),
    role: str = Form(""),
):
    # Accept either JSON list or comma-separated strings
    try:
        parsed = json.loads(holidays)
        if isinstance(parsed, list):
            holiday_list = parsed
        elif isinstance(parsed, str):
            holiday_list = [parsed]
        else:
            holiday_list = []
    except Exception:
        holiday_list = [h.strip() for h in holidays.split(",") if h.strip()]

    file_bytes = await file.read()

    output_file = update_project_schedule_stream(
        file_bytes=file_bytes,
        project_start_date=project_start_date,
        holidays=holiday_list,
        role=role,
    )

    return StreamingResponse(
        output_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename=updated_{file.filename}",
        },
    )