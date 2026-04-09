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
    project_manager: str = Form(""),
    business_consultant: str = Form(""),
    technical_consultant: str = Form(""),
    customer: str = Form(""),
):
    try:
        parsed = json.loads(holidays) if holidays else []
        holiday_list = parsed if isinstance(parsed, list) else []
    except Exception:
        holiday_list = [h.strip() for h in holidays.split(",") if h.strip()]

    output_file = update_project_schedule_stream(
        project_start_date=project_start_date,
        holidays=holiday_list,
        project_manager=project_manager,
        business_consultant=business_consultant,
        technical_consultant=technical_consultant,
        customer=customer,
    )

    return StreamingResponse(
        output_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=output.xlsx"},
    )