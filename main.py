"""
Smart Drive — 智能文件助手 后端
FastAPI + SQLite + DeepSeek (via OpenAI SDK)
"""

import os, re, json, sqlite3
from datetime import datetime
from pathlib import Path
from uuid import uuid4
from dotenv import load_dotenv

# ─── 加载 .env 文件（必须在读取环境变量之前调用） ───
load_dotenv()

# ─── 强制设置 Tesseract 路径（必须在 import pytesseract 之后立即设置） ───
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

import pdfplumber
from PIL import Image
from pptx import Presentation
from fastapi import FastAPI, UploadFile, File, Query, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from openai import OpenAI

# ─── 配置 ───────────────────────────────────────────
UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)
DB_PATH = "files.db"
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_BASE_URL = "https://api.deepseek.com"
DEEPSEEK_MODEL = "deepseek-chat"
print(f"[配置] DEEPSEEK_API_KEY: {'***' + DEEPSEEK_API_KEY[-6:] if DEEPSEEK_API_KEY and len(DEEPSEEK_API_KEY) > 6 else '(未设置)'}")
print(f"[配置] DEEPSEEK_BASE_URL: {DEEPSEEK_BASE_URL}")
print(f"[配置] DEEPSEEK_MODEL:    {DEEPSEEK_MODEL}")

# ─── DeepSeek 客户端（OpenAI SDK） ─────────────────
client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_BASE_URL)

# ─── FastAPI 实例 ──────────────────────────────────
app = FastAPI(title="Smart Drive")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─── 数据库初始化 ──────────────────────────────────
def init_db() -> None:
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS files (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            original_name TEXT    NOT NULL,
            new_name      TEXT    NOT NULL,
            content       TEXT    DEFAULT '',
            keywords      TEXT    DEFAULT '',
            date          TEXT    NOT NULL
        )
    """)
    conn.close()

init_db()


# ─── 文本提取 ──────────────────────────────────────
def extract_text(file_path: str, filename: str) -> str:
    """根据文件扩展名提取全文文本（不经过大模型）"""
    ext = Path(filename).suffix.lower()
    text = ""
    if ext == ".pdf":
        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t: text += t + "\n"
        except Exception as e:
            print(f"[PDF 提取失败] {e}")
            raise HTTPException(status_code=500, detail=f"PDF 提取失败: {e}")
    elif ext in (".ppt", ".pptx"):
        try:
            prs = Presentation(file_path)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            text += para.text + "\n"
        except Exception as e:
            print(f"[PPT 提取失败] {e}")
            raise HTTPException(status_code=500, detail=f"PPT 提取失败: {e}")
    elif ext in (".png",".jpg",".jpeg",".bmp",".tiff",".tif",".gif",".webp"):
        try:
            img = Image.open(file_path)
            text = pytesseract.image_to_string(img, lang="chi_sim+eng")
        except Exception as e:
            print(f"[OCR 提取失败] {e}")
            raise HTTPException(status_code=500, detail=f"OCR 提取失败: {e}")
    elif ext == ".docx":
        try:
            from docx import Document
            doc = Document(file_path)
            for para in doc.paragraphs:
                if para.text:
                    text += para.text + "\n"
        except Exception as e:
            print(f"[DOCX 提取失败] {e}")
            raise HTTPException(status_code=500, detail=f"DOCX 提取失败: {e}")
    elif ext in (".txt",".md",".csv",".json",".xml",".html",".htm",".log",".py",".js",".css"):
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                text = f.read()
        except Exception as e:
            print(f"[文本读取失败] {e}")
            raise HTTPException(status_code=500, detail=f"文本读取失败: {e}")
    return text.strip()


# ─── DeepSeek：生成标题 + 关键词 ───────────────────
def generate_title_and_keywords(text: str) -> tuple:
    """调用 DeepSeek 生成 15 字内标题和 3 个关键词。任何环节失败直接抛 HTTPException。"""
    # ── 空文本检查：不再静默兜底，直接报错 ──
    print(f"[DeepSeek] 步骤1: 检查提取文本, 长度={len(text) if text else 0}")
    if not text or not text.strip():
        print(f"[DeepSeek] ❌ 提取文本为空！无法调用 AI。")
        raise HTTPException(status_code=500, detail="提取文本为空，无法生成标题。请检查文件内容或提取器（Tesseract/PDF/PPT）是否正确安装。")

    # ── API Key 检查 ──
    print(f"[DeepSeek] 步骤2: 检查 API Key, 末6位={DEEPSEEK_API_KEY[-6:] if len(DEEPSEEK_API_KEY) > 6 else '(无)'}")
    if not DEEPSEEK_API_KEY:
        print(f"[DeepSeek] ❌ DEEPSEEK_API_KEY 为空！")
        raise HTTPException(status_code=500, detail="未配置 DEEPSEEK_API_KEY 环境变量。")

    snippet = text[:2000]
    prompt = (
        "请根据以下文件内容，完成两项任务：\n"
        "1. 生成一个不超过15个字的简洁标题\n"
        "2. 提取3个关键词\n\n"
        "严格按以下JSON格式返回，不要添加markdown代码块标记，不要添加任何多余内容：\n"
        '{"title": "标题", "keywords": ["关键词1", "关键词2", "关键词3"]}\n\n'
        f"文件内容：\n{snippet}"
    )
    print(f"[DeepSeek] 步骤3: 调用 {DEEPSEEK_MODEL} @ {DEEPSEEK_BASE_URL}, 发送 {len(snippet)} 字符")

    # ── 调用 API：出错直接抛 ──
    print(f"[DeepSeek] 步骤4: 正在请求 DeepSeek API...")
    response = client.chat.completions.create(
        model=DEEPSEEK_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=200,
    )
    raw = response.choices[0].message.content.strip()
    print(f"[DeepSeek] ✅ API 返回成功, 长度={len(raw)}")
    print(f"[DeepSeek] ✅ 原始返回内容:\n{raw}")

    # ── 兼容 markdown ```json ... ``` 包裹 ──
    cleaned = raw
    if "```" in cleaned:
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
        cleaned = cleaned.strip()
        print(f"[DeepSeek] 去除markdown包裹后:\n{cleaned}")

    # ── 解析 JSON：失败直接抛 ──
    print(f"[DeepSeek] 步骤5: 解析 JSON...")
    data = None
    try:
        data = json.loads(cleaned)
        print(f"[DeepSeek] ✅ 直接 JSON.loads 解析成功")
    except json.JSONDecodeError as e1:
        print(f"[DeepSeek] ⚠️ 直接解析失败: {e1}, 尝试正则提取...")
        m = re.search(r"\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}", cleaned, re.DOTALL)
        if m:
            try:
                data = json.loads(m.group())
                print(f"[DeepSeek] ✅ 正则提取后解析成功")
            except json.JSONDecodeError as e2:
                print(f"[DeepSeek] ❌ 正则提取后仍然解析失败: {e2}")
        else:
            print(f"[DeepSeek] ❌ 正则未找到 JSON 对象")

    if not (data and isinstance(data, dict)):
        print(f"[DeepSeek] ❌ 无法从返回内容中解析出 JSON 对象")
        raise HTTPException(status_code=500, detail=f"DeepSeek 返回内容无法解析为 JSON: {raw[:200]}")

    # ── 提取字段 ──
    title = str(data.get("title", "")).strip()[:15]
    kw = data.get("keywords", [])
    if not isinstance(kw, list):
        kw = [str(kw)]
    keywords = ",".join(str(k).strip() for k in kw[:3] if k)
    if not title:
        print(f"[DeepSeek] ⚠️ 返回的 title 为空")
        raise HTTPException(status_code=500, detail=f"DeepSeek 返回的 title 为空, 原始返回: {raw[:200]}")
    if not keywords:
        keywords = "未分类"
    print(f"[DeepSeek] ✅ 步骤6 完成 → 标题: {title}, 关键词: {keywords}")
    return title, keywords


# ─── 接口：上传文件 ────────────────────────────────
@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    print(f"\n{'='*60}")
    print(f"[上传 步骤1] 收到文件: {file.filename}")
    temp_name = f"temp_{uuid4().hex}_{file.filename}"
    temp_path = UPLOAD_DIR / temp_name
    with open(temp_path, "wb") as f:
        f.write(await file.read())
    print(f"[上传 步骤1] ✅ 已保存临时文件: {temp_path}")

    # 提取文本
    print(f"[上传 步骤2] 开始提取文本...")
    text = extract_text(str(temp_path), file.filename)
    print(f"[上传 步骤2] ✅ 文本提取完成, 长度: {len(text)} 字符")
    if len(text) == 0:
        print(f"[上传 步骤2] ❌ 提取到的文本为空！文件类型: {Path(file.filename).suffix}")
    else:
        print(f"[上传 步骤2] 前 200 字: {text[:200]}")

    # 调用 DeepSeek（任何异常会以 HTTPException 抛出，不再吞掉）
    print(f"[上传 步骤3] 开始调用 DeepSeek 生成标题和关键词...")
    title, keywords = generate_title_and_keywords(text)
    print(f"[上传 步骤3] ✅ AI 返回 → 标题: {title}, 关键词: {keywords}")

    # 重命名
    print(f"[上传 步骤4] 重命名文件...")
    ext = Path(file.filename).suffix
    new_filename = f"{title}{ext}"
    new_path = UPLOAD_DIR / new_filename
    c = 1
    while new_path.exists():
        new_filename = f"{title}_{c}{ext}"
        new_path = UPLOAD_DIR / new_filename
        c += 1
    os.rename(str(temp_path), str(new_path))
    print(f"[上传 步骤4] ✅ {file.filename} → {new_filename}")

    # 写库
    print(f"[上传 步骤5] 写入数据库...")
    date_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT INTO files (original_name,new_name,content,keywords,date) VALUES (?,?,?,?,?)",
        (file.filename, new_filename, text, keywords, date_str),
    )
    conn.commit(); conn.close()
    print(f"[上传 步骤5] ✅ 数据库写入完成, 日期: {date_str}")
    print(f"{'='*60}\n")
    return {"status":"ok","original_name":file.filename,"new_name":new_filename,
            "title":title,"keywords":keywords,"date":date_str}


# ─── 接口：搜索文件（纯本地 LIKE，零大模型调用） ───
@app.get("/search")
async def search_files(query: str = Query(default="")):
    try:
        conn = sqlite3.connect(DB_PATH)
        if query.strip() == "":
            cursor = conn.execute(
                "SELECT id,original_name,new_name,keywords,date FROM files ORDER BY id DESC")
        else:
            like = f"%{query}%"
            cursor = conn.execute("""
                SELECT id,original_name,new_name,keywords,date FROM files
                WHERE original_name LIKE ? OR new_name LIKE ?
                   OR content LIKE ? OR keywords LIKE ? OR date LIKE ?
                ORDER BY id DESC""", (like, like, like, like, like))
        results = [{"id": r[0], "original_name": r[1], "new_name": r[2],
                    "filename": r[2], "keywords": r[3], "date": r[4]}
                   for r in cursor.fetchall()]
        conn.close()
        print(f"[搜索] query='{query}', 返回 {len(results)} 条结果")
        return results
    except Exception as e:
        print(f"[搜索] ❌ 数据库查询出错: {e}")
        return []


# ─── PWA：manifest.json ────────────────────────────
@app.get("/manifest.json")
async def manifest():
    return JSONResponse({
        "name":"智能文件助手","short_name":"SmartDrive",
        "description":"AI 驱动的智能文件管理助手",
        "start_url":"/","display":"standalone",
        "background_color":"#ffffff","theme_color":"#007AFF",
        "icons":[{"src":"data:image/svg+xml,<svg xmlns=%27http://www.w3.org/2000/svg%27 viewBox=%270 0 100 100%27><rect fill=%27%23007AFF%27 width=%27100%27 height=%27100%27 rx=%2720%27/><text x=%2750%27 y=%2768%27 font-size=%2750%27 fill=%27white%27 text-anchor=%27middle%27>📁</text></svg>",
                 "sizes":"any","type":"image/svg+xml","purpose":"any maskable"}]
    })


# ─── PWA：Service Worker ───────────────────────────
@app.get("/sw.js")
async def service_worker():
    from starlette.responses import Response
    sw = """const C='smart-drive-v1',A=['/','/manifest.json'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(A)));self.skipWaiting()});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==C).map(k=>caches.delete(k)))));self.clients.claim()});
self.addEventListener('fetch',e=>{if(e.request.method!=='GET')return;var u=e.request.url;if(u.includes('/search')||u.includes('/upload')||u.includes('/files/')){e.respondWith(fetch(e.request).catch(()=>new Response('Offline',{status:503})));return}e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(resp=>{if(resp.ok){const cl=resp.clone();caches.open(C).then(c=>c.put(e.request,cl))}return resp}).catch(()=>new Response('Offline',{status:503}))))});"""
    return Response(content=sw, media_type="application/javascript")


# ─── 静态文件挂载（供前端下载） ────────────────────
app.mount("/files", StaticFiles(directory=str(UPLOAD_DIR)), name="files")


# ─── 首页 ──────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def read_index():
    with open("index.html", "r", encoding="utf-8") as f:
        return f.read()


# ─── 启动入口 ──────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
