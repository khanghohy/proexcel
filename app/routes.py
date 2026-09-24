import io
from fastapi import APIRouter, UploadFile, File, Form, HTTPException, Response
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Optional, List
from PIL import Image as PILImage
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from app.database import (
    get_all_records,
    create_record,
    update_record,
    delete_record,
    save_image_blob,
    remove_image_blob,
    get_image_blob,
    get_database_info,
    run_custom_sql,
    get_connection
)

router = APIRouter(prefix="/api")

class UpdateRecordRequest(BaseModel):
    name: str

class CreateRecordRequest(BaseModel):
    name: Optional[str] = ""

class Base64UploadRequest(BaseModel):
    data_base64: str
    filename: Optional[str] = "clipboard_image.png"

class SQLQueryRequest(BaseModel):
    query: str

@router.get("/records")
def list_records():
    return {"records": get_all_records()}

@router.post("/records")
def add_record(payload: CreateRecordRequest):
    record = create_record(name=payload.name or "")
    return {"success": True, "record": record}

@router.put("/records/{record_id}")
def update_record_endpoint(record_id: int, payload: UpdateRecordRequest):
    success = update_record(record_id, payload.name)
    if not success:
        raise HTTPException(status_code=404, detail="Không tìm thấy hàng này")
    return {"success": True}

@router.delete("/records/{record_id}")
def delete_record_endpoint(record_id: int):
    success = delete_record(record_id)
    if not success:
        raise HTTPException(status_code=404, detail="Không tìm thấy hàng này")
    return {"success": True}

@router.post("/records/{record_id}/upload-image/{image_type}")
async def upload_image(record_id: int, image_type: str, file: UploadFile = File(...)):
    if image_type not in ("full", "half"):
        raise HTTPException(status_code=400, detail="Loại ảnh không hợp lệ (chỉ full hoặc half)")
    
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="Tập tin rỗng")
        
    mime_type = file.content_type or "image/png"
    filename = file.filename or f"{image_type}_{record_id}.png"
    
    success = save_image_blob(record_id, image_type, contents, filename, mime_type)
    if not success:
        raise HTTPException(status_code=404, detail="Không tìm thấy bản ghi")
        
    return {
        "success": True, 
        "image_url": f"/api/records/{record_id}/image/{image_type}",
        "filename": filename,
        "size": len(contents),
        "mime": mime_type
    }

@router.post("/records/{record_id}/upload-base64/{image_type}")
async def upload_base64_image(record_id: int, image_type: str, payload: Base64UploadRequest):
    import base64
    if image_type not in ("full", "half"):
        raise HTTPException(status_code=400, detail="Loại ảnh không hợp lệ")
    
    raw_str = payload.data_base64
    mime_type = "image/png"
    if "," in raw_str:
        header, raw_str = raw_str.split(",", 1)
        if "data:" in header and ";base64" in header:
            mime_type = header.split("data:")[1].split(";base64")[0]
            
    try:
        data_bytes = base64.b64decode(raw_str)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Dữ liệu ảnh không hợp lệ: {str(e)}")
        
    filename = payload.filename or f"pasted_{image_type}_{record_id}.png"
    success = save_image_blob(record_id, image_type, data_bytes, filename, mime_type)
    if not success:
        raise HTTPException(status_code=404, detail="Không tìm thấy bản ghi")
        
    return {
        "success": True,
        "image_url": f"/api/records/{record_id}/image/{image_type}",
        "filename": filename,
        "size": len(data_bytes),
        "mime": mime_type
    }

@router.delete("/records/{record_id}/image/{image_type}")
def delete_image(record_id: int, image_type: str):
    if image_type not in ("full", "half"):
        raise HTTPException(status_code=400, detail="Loại ảnh không hợp lệ")
    success = remove_image_blob(record_id, image_type)
    return {"success": success}

@router.get("/records/{record_id}/image/{image_type}")
def get_image(record_id: int, image_type: str):
    data = get_image_blob(record_id, image_type)
    if not data:
        raise HTTPException(status_code=404, detail="Không có ảnh")
    
    blob, mime, filename = data
    return Response(
        content=blob,
        media_type=mime,
        headers={
            "Content-Disposition": f'inline; filename="{filename}"',
            "Cache-Control": "no-cache"
        }
    )

@router.get("/db/stats")
def db_stats():
    return get_database_info()

@router.post("/db/query")
def execute_sql(payload: SQLQueryRequest):
    result = run_custom_sql(payload.query)
    return result

@router.get("/db/download")
def download_database():
    from app.database import DB_PATH
    from fastapi.responses import FileResponse
    if not DB_PATH.exists():
        raise HTTPException(status_code=404, detail="Chưa có tệp CSDL")
    return FileResponse(
        str(DB_PATH),
        filename="excel_app.db",
        media_type="application/x-sqlite3"
    )


@router.get("/export/excel")
def export_excel():
    """Generate Excel spreadsheet with text columns and embedded images."""
    records = get_all_records()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dữ Liệu Ảnh"
    
    # Styling definitions
    header_fill = PatternFill(start_color="107C41", end_color="107C41", fill_type="solid") # Excel Green
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    headers = ["STT", "Tên", "Ảnh chụp full", "Ảnh chụp 1 nửa", "Ngày tạo"]
    ws.append(headers)
    
    # Style header row
    ws.row_dimensions[1].height = 28
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
        
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 22
    
    for idx, r in enumerate(records, start=1):
        row_num = idx + 1
        ws.row_dimensions[row_num].height = 80 # Make room for thumbnail image
        
        # STT
        c_stt = ws.cell(row=row_num, column=1, value=idx)
        c_stt.alignment = center_align
        c_stt.font = data_font
        c_stt.border = thin_border
        
        # Name
        c_name = ws.cell(row=row_num, column=2, value=r["name"])
        c_name.alignment = left_align
        c_name.font = data_font
        c_name.border = thin_border
        
        # Full Image
        c_full = ws.cell(row=row_num, column=3)
        c_full.alignment = center_align
        c_full.border = thin_border
        if r["full_image"]["exists"]:
            full_data = get_image_blob(r["id"], "full")
            if full_data:
                try:
                    img_bytes, _, _ = full_data
                    pil_img = PILImage.open(io.BytesIO(img_bytes))
                    # Thumbnail to fit 160x70
                    pil_img.thumbnail((160, 95))
                    img_buffer = io.BytesIO()
                    pil_img.save(img_buffer, format="PNG")
                    img_buffer.seek(0)
                    xl_img = openpyxl.drawing.image.Image(img_buffer)
                    xl_img.width = pil_img.width
                    xl_img.height = pil_img.height
                    col_letter = get_column_letter(3)
                    ws.add_image(xl_img, f"{col_letter}{row_num}")
                except Exception:
                    c_full.value = r["full_image"]["name"] or "Có ảnh"
        else:
            c_full.value = "(Chưa có ảnh)"
            c_full.font = Font(name="Calibri", size=9, italic=True, color="888888")
            
        # Half Image
        c_half = ws.cell(row=row_num, column=4)
        c_half.alignment = center_align
        c_half.border = thin_border
        if r["half_image"]["exists"]:
            half_data = get_image_blob(r["id"], "half")
            if half_data:
                try:
                    img_bytes, _, _ = half_data
                    pil_img = PILImage.open(io.BytesIO(img_bytes))
                    pil_img.thumbnail((160, 95))
                    img_buffer = io.BytesIO()
                    pil_img.save(img_buffer, format="PNG")
                    img_buffer.seek(0)
                    xl_img = openpyxl.drawing.image.Image(img_buffer)
                    xl_img.width = pil_img.width
                    xl_img.height = pil_img.height
                    col_letter = get_column_letter(4)
                    ws.add_image(xl_img, f"{col_letter}{row_num}")
                except Exception:
                    c_half.value = r["half_image"]["name"] or "Có ảnh"
        else:
            c_half.value = "(Chưa có ảnh)"
            c_half.font = Font(name="Calibri", size=9, italic=True, color="888888")
            
        # Created at
        c_time = ws.cell(row=row_num, column=5, value=str(r["created_at"]))
        c_time.alignment = center_align
        c_time.font = data_font
        c_time.border = thin_border

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="bang_du_lieu_anh.xlsx"'}
    )
