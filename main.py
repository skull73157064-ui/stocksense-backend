"""
StockSense 抽圖後端
功能:接收 Excel,抽出內嵌圖片並對應款號,上傳到 Supabase Storage,回傳結果。
"""
import os, re, zipfile, tempfile, subprocess, shutil, uuid, base64
from typing import List, Dict, Optional
from fastapi import FastAPI, UploadFile, File, HTTPException, Header, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, Response
import openpyxl
import httpx
import jwt as pyjwt

app = FastAPI(title="StockSense Image Extractor")

# CORS: 最暴力可靠的做法,直接在每個 response 加 header
@app.middleware("http")
async def add_cors_headers(request: Request, call_next):
    if request.method == "OPTIONS":
        # preflight 直接回 200 + 允許 header
        response = Response(status_code=200)
        response.headers["Access-Control-Allow-Origin"] = "*"
        response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "*"
        response.headers["Access-Control-Max-Age"] = "86400"
        return response
    response = await call_next(request)
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "*"
    return response

# Supabase 設定(從環境變數讀,部署時設定)
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")  # service_role key,後端用
SUPABASE_JWT_SECRET = os.environ.get("SUPABASE_JWT_SECRET", "")    # 用來驗 access token

BUCKET = "product-images"


def verify_token(authorization: Optional[str] = Header(None)) -> Dict:
    """驗證前端傳來的 Supabase JWT,確認是合法登入使用者。"""
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="缺少 Authorization header")
    token = authorization.split(" ", 1)[1]
    try:
        # Supabase JWT 用 HS256 簽,secret 從專案設定取得
        payload = pyjwt.decode(token, SUPABASE_JWT_SECRET, algorithms=["HS256"], audience="authenticated")
        return payload
    except pyjwt.ExpiredSignatureError:
        raise HTTPException(status_code=401, detail="Token 已過期")
    except pyjwt.InvalidTokenError as e:
        raise HTTPException(status_code=401, detail=f"Token 無效: {e}")


def convert_xls_to_xlsx(input_path: str, output_dir: str) -> str:
    """用 LibreOffice 把 .xls 轉成 .xlsx,回傳新檔路徑。"""
    result = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "xlsx", "--outdir", output_dir, input_path],
        capture_output=True, text=True, timeout=120
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice 轉檔失敗: {result.stderr}")
    base = os.path.splitext(os.path.basename(input_path))[0]
    out = os.path.join(output_dir, base + ".xlsx")
    if not os.path.exists(out):
        raise RuntimeError("LibreOffice 轉檔後找不到 .xlsx 檔")
    return out


def extract_images_with_styles(xlsx_path: str) -> List[Dict]:
    """
    從 xlsx 抽圖並對應款號。
    回傳: [{style_no, part, image_bytes, image_ext}, ...]
    """
    results = []
    seen_styles = set()
    z = zipfile.ZipFile(xlsx_path)
    names = z.namelist()
    ws_files = sorted([n for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n)])
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    sheetnames = wb.sheetnames

    for idx, wsfile in enumerate(ws_files):
        sn = sheetnames[idx] if idx < len(sheetnames) else "?"
        if 'A品' not in sn and 'A 品' not in sn:
            continue
        rels_path = f"xl/worksheets/_rels/{os.path.basename(wsfile)}.rels"
        if rels_path not in names:
            continue
        dm = re.search(r'drawing(\d+)\.xml', z.read(rels_path).decode('utf-8', 'ignore'))
        if not dm:
            continue
        draw = f"xl/drawings/drawing{dm.group(1)}.xml"
        if draw not in names:
            continue
        dcontent = z.read(draw).decode('utf-8', 'ignore')
        anchors = re.findall(
            r'<xdr:from>.*?<xdr:row>(\d+)</xdr:row>.*?r:embed="(rId\d+)"',
            dcontent, re.S
        )
        drels_path = f"xl/drawings/_rels/drawing{dm.group(1)}.xml.rels"
        rid2img = {}
        if drels_path in names:
            rid2img = {
                m[0]: m[1] for m in re.findall(
                    r'Id="(rId\d+)"[^>]*Target="\.\./media/([^"]+)"',
                    z.read(drels_path).decode('utf-8', 'ignore')
                )
            }
        ws = wb[sn]
        row2style = {}
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, 4).value
            if v and re.match(r'[A-Z]{2,4}\d', str(v)):
                row2style[r - 1] = str(v).strip()  # openpyxl 1-based, drawing 0-based

        for row, rid in anchors:
            style = row2style.get(int(row))
            img = rid2img.get(rid)
            if style and img and style not in seen_styles:
                ext = img.split('.')[-1].lower()
                if ext == 'jpeg':
                    ext = 'jpg'
                img_data = z.read('xl/media/' + img)
                results.append({
                    "style_no": style,
                    "part": "",
                    "image_bytes": img_data,
                    "image_ext": ext,
                })
                seen_styles.add(style)
    return results


async def upload_to_supabase_storage(image_bytes: bytes, filename: str, access_token: str) -> str:
    """上傳圖片到 Supabase Storage,回傳公開 URL。用使用者的 token,符合 RLS。"""
    url = f"{SUPABASE_URL}/storage/v1/object/{BUCKET}/{filename}"
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.post(
            url,
            content=image_bytes,
            headers={
                "Authorization": f"Bearer {access_token}",
                "Content-Type": f"image/{filename.split('.')[-1]}",
                "x-upsert": "true",
            },
        )
        if r.status_code not in (200, 201):
            raise RuntimeError(f"Storage 上傳失敗 ({r.status_code}): {r.text}")
    return f"{SUPABASE_URL}/storage/v1/object/public/{BUCKET}/{filename}"


@app.get("/")
def health():
    return {"ok": True, "service": "StockSense Image Extractor"}


@app.options("/extract")
def extract_preflight():
    """明確處理 CORS preflight"""
    return {"ok": True}


@app.post("/extract")
async def extract(
    file: UploadFile = File(...),
    authorization: Optional[str] = Header(None),
):
    """主端點:接收 Excel,抽圖,上傳 Storage,回傳結果供前端核對。"""
    # 驗身份(同時拿原始 token 給 Storage 上傳用)
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="缺少 Authorization header")
    access_token = authorization.split(" ", 1)[1]
    payload = verify_token(authorization)

    # 存上傳檔到暫存目錄
    tmpdir = tempfile.mkdtemp(prefix="ssx_")
    try:
        in_path = os.path.join(tmpdir, file.filename)
        content = await file.read()
        with open(in_path, "wb") as f:
            f.write(content)

        # 若是 .xls 先轉 .xlsx
        if file.filename.lower().endswith(".xls"):
            xlsx_path = convert_xls_to_xlsx(in_path, tmpdir)
        elif file.filename.lower().endswith(".xlsx"):
            xlsx_path = in_path
        else:
            raise HTTPException(status_code=400, detail="只接受 .xls 或 .xlsx 檔")

        # 抽圖+對應款號
        images = extract_images_with_styles(xlsx_path)

        # 上傳到 Storage
        results = []
        for img in images:
            unique_name = f"{img['style_no']}_{uuid.uuid4().hex[:8]}.{img['image_ext']}"
            try:
                url = await upload_to_supabase_storage(img["image_bytes"], unique_name, access_token)
                results.append({
                    "style_no": img["style_no"],
                    "image_url": url,
                    "filename": unique_name,
                })
            except Exception as e:
                results.append({
                    "style_no": img["style_no"],
                    "error": str(e),
                })

        return {
            "ok": True,
            "extracted": len(images),
            "uploaded": len([r for r in results if "image_url" in r]),
            "results": results,
        }
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
