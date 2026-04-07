from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse
import json

from src.performance_planner import update_project_schedule_stream

app = FastAPI()


@app.get("/")
def home():
    return {"message": "API running on Vercel 🚀"}


@app.post("/schedule")
async def generate_schedule(
    file: UploadFile = File(...),
    project_start_date: str = Form(...),
    holidays: str = Form("[]"),
):
    try:
        holiday_list = json.loads(holidays)
    except:
        holiday_list = [h.strip() for h in holidays.split(",") if h.strip()]

    file_bytes = await file.read()

    output_file = update_project_schedule_stream(
        file_bytes=file_bytes,
        project_start_date=project_start_date,
        holidays=holiday_list,
    )

    return StreamingResponse(
        output_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename=updated_{file.filename}"
        },
    )