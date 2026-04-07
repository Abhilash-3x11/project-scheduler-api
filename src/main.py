from fastapi import FastAPI, Form
from fastapi.responses import StreamingResponse
import json

from src.performance_planner import update_project_schedule_stream

app = FastAPI(title="Project Scheduler API")


@app.get("/")
def home():
    return {"message": "API running"}


@app.post("/schedule")
async def generate_schedule(
    project_start_date: str = Form(...),
    holidays: str = Form("[]"),
    role: str = Form(""),
):
    try:
        parsed = json.loads(holidays)
        if isinstance(parsed, list):
            holiday_list = parsed
        else:
            holiday_list = []
    except Exception:
        holiday_list = [h.strip() for h in holidays.split(",") if h.strip()]

    output_file = update_project_schedule_stream(
        project_start_date=project_start_date,
        holidays=holiday_list,
        role=role,
    )

    return StreamingResponse(
        output_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=output.xlsx"},
    )