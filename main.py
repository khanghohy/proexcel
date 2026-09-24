import uvicorn
from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path

from app.database import init_db
from app.routes import router

app = FastAPI(
    title="Excel Image Sheet with Local SQL",
    description="Ứng dụng bảng tính Excel hỗ trợ chụp/đính kèm ảnh lưu trực tiếp vào cơ sở dữ liệu SQL trên máy tính",
    version="1.0.0"
)

# CORS Middleware to allow requests from GitHub Pages (https://khanghohy.github.io)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Private Network Access header support
@app.middleware("http")
async def add_pna_header(request: Request, call_next):
    if request.method == "OPTIONS":
        response = await call_next(request)
        response.headers["Access-Control-Allow-Private-Network"] = "true"
        response.headers["Access-Control-Allow-Origin"] = "*"
        return response
    response = await call_next(request)
    response.headers["Access-Control-Allow-Private-Network"] = "true"
    return response

# Initialize database
init_db()

# Include API routes
app.include_router(router)

# Mount static folders
ROOT_DIR = Path(__file__).resolve().parent
STATIC_DIR = ROOT_DIR / "static"
CSS_DIR = ROOT_DIR / "css"
JS_DIR = ROOT_DIR / "js"

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
if CSS_DIR.exists():
    app.mount("/css", StaticFiles(directory=str(CSS_DIR)), name="css")
if JS_DIR.exists():
    app.mount("/js", StaticFiles(directory=str(JS_DIR)), name="js")

@app.get("/")
def index():
    return FileResponse(str(ROOT_DIR / "index.html" if (ROOT_DIR / "index.html").exists() else STATIC_DIR / "index.html"))

if __name__ == "__main__":
    print("=================================================================")
    print("🚀 Đang khởi động hệ thống Bảng Tính Excel & SQL Database...")
    print("📍 Truy cập ứng dụng tại: http://localhost:8000")
    print("📍 Mạng nội bộ (LAN): http://0.0.0.0:8000")
    print("=================================================================")
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
