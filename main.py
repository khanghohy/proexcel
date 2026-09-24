import uvicorn
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pathlib import Path

from app.database import init_db
from app.routes import router

app = FastAPI(
    title="Excel Image Sheet with Local SQL",
    description="Ứng dụng bảng tính Excel hỗ trợ chụp/đính kèm ảnh lưu trực tiếp vào cơ sở dữ liệu SQL trên máy tính",
    version="1.0.0"
)

# Initialize database
init_db()

# Include API routes
app.include_router(router)

# Mount static folders (supports both /static and relative ./css, ./js)
ROOT_DIR = Path(__file__).resolve().parent
STATIC_DIR = ROOT_DIR / "static"
CSS_DIR = STATIC_DIR / "css"
JS_DIR = STATIC_DIR / "js"

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
app.mount("/css", StaticFiles(directory=str(CSS_DIR)), name="css")
app.mount("/js", StaticFiles(directory=str(JS_DIR)), name="js")

@app.get("/")
def index():
    return FileResponse(str(STATIC_DIR / "index.html"))

if __name__ == "__main__":
    print("=================================================================")
    print("🚀 Đang khởi động hệ thống Bảng Tính Excel & SQL Database...")
    print("📍 Truy cập ứng dụng tại: http://localhost:8000")
    print("📍 Mạng nội bộ (LAN): http://0.0.0.0:8000")
    print("=================================================================")
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
