# -*- coding: utf-8 -*-
"""
網頁版:把 Word (.doc/.docx)、圖片 (.jpg/.jpeg/.png) 和 PDF 拖進瀏覽器,合併成一個 PDF
所有檔案都只在自己的電腦上處理,不會上傳到網路。

安裝:
    pip install flask pillow pypdf pywin32
執行:
    python merge_web.py
瀏覽器會自動開啟 http://127.0.0.1:5000
(Windows 需已安裝 Microsoft Word;沒有 Word 的話會改用 LibreOffice)
"""
import atexit
import io
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import uuid
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from PIL import Image, ImageOps
from pypdf import PdfReader, PdfWriter

PORT = 5000
WORD_EXT = {".doc", ".docx"}
IMG_EXT = {".jpg", ".jpeg", ".png"}
PDF_EXT = {".pdf"}
DPI = 200
A4 = (int(8.27 * DPI), int(11.69 * DPI))

WORK = Path(tempfile.mkdtemp(prefix="merge_pdf_"))
atexit.register(lambda: shutil.rmtree(WORK, ignore_errors=True))

app = Flask(__name__)


# ---------- 圖片 -> PDF ----------
def image_to_pdf(src: Path, dst: Path, fit_a4: bool):
    img = ImageOps.exif_transpose(Image.open(src))  # 依拍照方向自動轉正
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, "white")
        bg.paste(img, mask=img.split()[-1])
        img = bg
    else:
        img = img.convert("RGB")

    if fit_a4:
        page = A4 if img.height >= img.width else (A4[1], A4[0])  # 橫圖用橫式 A4
        margin = int(0.4 * DPI)
        box = (page[0] - 2 * margin, page[1] - 2 * margin)
        scale = min(box[0] / img.width, box[1] / img.height)
        img = img.resize((max(1, int(img.width * scale)), max(1, int(img.height * scale))),
                         Image.LANCZOS)
        canvas = Image.new("RGB", page, "white")
        canvas.paste(img, ((page[0] - img.width) // 2, (page[1] - img.height) // 2))
        img = canvas
    img.save(dst, "PDF", resolution=DPI)


# ---------- Word -> PDF ----------
_word = None


def find_libreoffice():
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    for path in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(path):
            return path
    return None


def get_word():
    """取得(或啟動)背景中的 Microsoft Word;不是 Windows 或沒有 Word 時回傳 None"""
    global _word
    if sys.platform != "win32":
        return None
    if _word is None:
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            _word = win32com.client.DispatchEx("Word.Application")
            _word.Visible = False
            _word.DisplayAlerts = 0
        except Exception:
            _word = None
    return _word


@atexit.register
def quit_word():
    global _word
    if _word is not None:
        try:
            _word.Quit()
        except Exception:
            pass
        _word = None


def word_to_pdf(src: Path, dst: Path):
    global _word
    for attempt in range(2):
        word = get_word()
        if word is None:
            break
        try:
            doc = word.Documents.Open(str(src.resolve()), ReadOnly=True)
            try:
                doc.SaveAs(str(dst.resolve()), FileFormat=17)  # 17 = PDF
            finally:
                doc.Close(False)
            return
        except Exception:
            _word = None            # Word 可能被關掉了,重新啟動再試一次
            if attempt == 1:
                raise

    soffice = find_libreoffice()
    if soffice is None:
        raise RuntimeError("找不到 Microsoft Word 或 LibreOffice,無法轉換 Word 檔")
    subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                    "--outdir", str(dst.parent), str(src)],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    produced = dst.parent / (src.stem + ".pdf")
    if not produced.exists():
        raise RuntimeError("LibreOffice 轉換失敗")
    produced.replace(dst)


# ---------- 網頁路由 ----------
@app.get("/")
def index():
    return PAGE


@app.post("/convert")
def convert():
    """接收一個檔案,轉成暫存 PDF,回傳代號與頁數"""
    f = request.files.get("file")
    if f is None:
        return jsonify(error="沒有收到檔案"), 400
    ext = Path(f.filename or "").suffix.lower()
    if ext not in WORD_EXT | IMG_EXT | PDF_EXT:
        return jsonify(error="不支援的檔案類型"), 400

    fid = uuid.uuid4().hex
    src = WORK / f"{fid}_src{ext}"
    dst = WORK / f"{fid}.pdf"
    f.save(src)
    try:
        if ext in PDF_EXT:
            shutil.copyfile(src, dst)                      # PDF 不用轉換,直接使用
        elif ext in IMG_EXT:
            image_to_pdf(src, dst, request.form.get("fit") == "1")
        else:
            word_to_pdf(src, dst)
        pages = len(PdfReader(str(dst)).pages)
    except Exception as e:
        return jsonify(error=str(e) or "轉換失敗"), 500
    finally:
        src.unlink(missing_ok=True)
    return jsonify(id=fid, pages=pages)


@app.post("/merge")
def merge():
    """依照傳入的代號順序合併成一個 PDF"""
    ids = (request.get_json(silent=True) or {}).get("ids", [])
    writer = PdfWriter()
    for fid in ids:
        if not (isinstance(fid, str) and fid.isalnum()):
            continue
        pdf = WORK / f"{fid}.pdf"
        if pdf.exists():
            writer.append(str(pdf))
    if len(writer.pages) == 0:
        return jsonify(error="沒有可合併的檔案"), 400
    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return send_file(buf, mimetype="application/pdf",
                     as_attachment=True, download_name="合併結果.pdf")


# ---------- 網頁畫面 ----------
PAGE = r"""<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Word / 圖片 / PDF 合併</title>
<style>
  :root { --ink:#1f2933; --sub:#6b7780; --line:#d9dee3; --bg:#f4f6f8; --card:#fff;
          --accent:#c0392b; --ok:#1e8e5a; }
  * { box-sizing:border-box; }
  body { margin:0; background:var(--bg); color:var(--ink);
         font-family:"Microsoft JhengHei","PingFang TC","Noto Sans TC",sans-serif; }
  main { max-width:760px; margin:0 auto; padding:32px 20px 60px; }
  h1 { font-size:24px; margin:0 0 4px; }
  .sub { color:var(--sub); font-size:14px; margin:0 0 20px; }
  #drop { border:2px dashed var(--line); border-radius:12px; background:var(--card);
          padding:36px 20px; text-align:center; cursor:pointer; transition:.15s; }
  #drop.over { border-color:var(--accent); background:#fdf3f1; }
  #drop strong { display:block; font-size:17px; margin-bottom:6px; }
  #drop span { color:var(--sub); font-size:13px; }
  .bar { display:flex; flex-wrap:wrap; gap:10px; align-items:center; margin:18px 0 10px; }
  .bar label { font-size:14px; color:var(--sub); margin-right:auto; }
  button { font:inherit; border:1px solid var(--line); background:var(--card); color:var(--ink);
           border-radius:8px; padding:7px 12px; cursor:pointer; }
  button:hover:not(:disabled) { border-color:#aab3bb; }
  button:disabled { opacity:.45; cursor:default; }
  #go { background:var(--accent); border-color:var(--accent); color:#fff;
        font-size:16px; padding:11px 26px; font-weight:600; }
  ul { list-style:none; margin:0; padding:0; }
  li { display:flex; align-items:center; gap:10px; background:var(--card);
       border:1px solid var(--line); border-radius:8px; padding:8px 10px; margin-bottom:6px; }
  li.drag { opacity:.4; }
  .no { width:26px; text-align:right; color:var(--sub); font-size:13px; }
  .tag { font-size:11px; font-weight:700; color:#fff; border-radius:4px; padding:2px 6px;
         background:#2b6cb0; }
  .tag.img { background:#b7791f; }
  .tag.pdf { background:#c0392b; }
  .name { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;
          font-size:14px; cursor:grab; }
  .st { font-size:12px; color:var(--sub); white-space:nowrap; }
  .st.ok { color:var(--ok); } .st.err { color:var(--accent); }
  li button { padding:3px 8px; font-size:13px; }
  #msg { margin-top:16px; font-size:14px; min-height:20px; }
  .foot { display:flex; align-items:center; gap:16px; margin-top:16px; flex-wrap:wrap; }
</style>
</head>
<body>
<main>
  <h1>Word / 圖片 / PDF 合併</h1>
  <p class="sub">檔案只在這台電腦上處理,不會上傳到網路。</p>

  <div id="drop">
    <strong>把檔案拖到這裡</strong>
    <span>或點一下選擇檔案 · 支援 .doc .docx .jpg .jpeg .png .pdf</span>
    <input id="pick" type="file" multiple accept=".doc,.docx,.jpg,.jpeg,.png,.pdf" hidden>
  </div>

  <div class="bar">
    <label><input id="fit" type="checkbox" checked> 圖片縮放到 A4 頁面</label>
    <button id="sort">依檔名排序</button>
    <button id="clear">全部清除</button>
  </div>

  <ul id="list"></ul>

  <div class="foot">
    <button id="go" disabled>合併成 PDF</button>
    <div id="msg"></div>
  </div>
</main>

<script>
const WORD = ['.doc', '.docx'], IMG = ['.jpg', '.jpeg', '.png'], PDF = ['.pdf'];
let items = [], busy = false, dragIdx = null;
const $ = id => document.getElementById(id);
const ext = n => (n.match(/\.[^.]+$/) || [''])[0].toLowerCase();
const natural = (a, b) => a.name.localeCompare(b.name, 'zh-Hant', {numeric: true});

function addFiles(files) {
  const added = [...files]
    .filter(f => [...WORD, ...IMG, ...PDF].includes(ext(f.name)) && !f.name.startsWith('~$'))
    .map(f => ({file: f, name: f.name, img: IMG.includes(ext(f.name)), pdf: PDF.includes(ext(f.name)),
                id: null, fit: null, st: '', cls: ''}))
    .sort(natural);
  const skipped = files.length - added.length;
  items.push(...added);
  $('msg').textContent = skipped ? `已略過 ${skipped} 個不支援的檔案` : '';
  render();
}

function render() {
  const ul = $('list');
  ul.innerHTML = '';
  items.forEach((it, i) => {
    const li = document.createElement('li');
    li.draggable = !busy;
    const no = Object.assign(document.createElement('span'), {className: 'no', textContent: i + 1});
    const tag = Object.assign(document.createElement('span'),
      {className: 'tag' + (it.img ? ' img' : it.pdf ? ' pdf' : ''),
       textContent: it.img ? '圖片' : it.pdf ? 'PDF' : 'Word'});
    const name = Object.assign(document.createElement('span'),
      {className: 'name', textContent: it.name, title: it.name});
    const st = Object.assign(document.createElement('span'),
      {className: 'st ' + it.cls, textContent: it.st});
    const mk = (t, fn, off) => {
      const b = Object.assign(document.createElement('button'), {textContent: t});
      b.disabled = busy || off; b.onclick = fn; return b;
    };
    li.append(no, tag, name, st,
      mk('↑', () => move(i, i - 1), i === 0),
      mk('↓', () => move(i, i + 1), i === items.length - 1),
      mk('✕', () => { items.splice(i, 1); render(); }));
    li.ondragstart = () => { dragIdx = i; li.classList.add('drag'); };
    li.ondragend = () => { dragIdx = null; li.classList.remove('drag'); };
    li.ondragover = e => { if (dragIdx !== null) e.preventDefault(); };
    li.ondrop = e => {
      if (dragIdx === null) return;
      e.preventDefault(); e.stopPropagation();
      move(dragIdx, i);
    };
    ul.append(li);
  });
  $('go').disabled = busy || items.length === 0;
  $('sort').disabled = $('clear').disabled = busy || items.length === 0;
}

function move(from, to) {
  if (to < 0 || to >= items.length || from === to) return;
  items.splice(to, 0, items.splice(from, 1)[0]);
  render();
}

async function run() {
  busy = true;
  const fit = $('fit').checked;
  let failed = 0;
  render();
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (it.id && (!it.img || it.fit === fit)) continue;      // 已轉換過,不用重做
    it.id = null; it.st = '轉換中…'; it.cls = ''; render();
    $('msg').textContent = `轉換中 ${i + 1} / ${items.length}`;
    try {
      const fd = new FormData();
      fd.append('file', it.file, it.name);
      fd.append('fit', fit ? '1' : '0');
      const r = await fetch('/convert', {method: 'POST', body: fd});
      const j = await r.json();
      if (!r.ok) throw new Error(j.error || '轉換失敗');
      it.id = j.id; it.fit = fit; it.st = j.pages + ' 頁'; it.cls = 'ok';
    } catch (e) {
      it.st = '失敗:' + e.message; it.cls = 'err'; failed++;
    }
    render();
  }
  failed = items.filter(it => !it.id).length;
  const ids = items.filter(it => it.id).map(it => it.id);
  try {
    if (!ids.length) throw new Error('沒有任何檔案轉換成功');
    $('msg').textContent = '合併中…';
    const r = await fetch('/merge', {method: 'POST',
      headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ids})});
    if (!r.ok) throw new Error((await r.json()).error || '合併失敗');
    const url = URL.createObjectURL(await r.blob());
    const a = Object.assign(document.createElement('a'), {href: url, download: '合併結果.pdf'});
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
    $('msg').textContent = failed
      ? `已下載,但有 ${failed} 個檔案失敗、未包含在 PDF 中`
      : '完成!PDF 已下載。';
  } catch (e) {
    $('msg').textContent = '錯誤:' + e.message;
  }
  busy = false;
  render();
}

const drop = $('drop');
drop.onclick = () => $('pick').click();
$('pick').onchange = e => { addFiles(e.target.files); e.target.value = ''; };
drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
drop.ondragleave = () => drop.classList.remove('over');
drop.ondrop = e => {
  e.preventDefault(); drop.classList.remove('over');
  if (e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
};
// 避免檔案掉到框外時瀏覽器直接把它打開
window.addEventListener('dragover', e => e.preventDefault());
window.addEventListener('drop', e => e.preventDefault());
$('sort').onclick = () => { items.sort(natural); render(); };
$('clear').onclick = () => { items = []; $('msg').textContent = ''; render(); };
$('go').onclick = run;
render();
</script>
</body>
</html>
"""

if __name__ == "__main__":
    url = f"http://127.0.0.1:{PORT}"
    print(f"已啟動:{url}")
    print("使用期間請不要關閉這個視窗;用完後按 Ctrl+C 或直接關閉視窗即可。")
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    # threaded=False:所有請求都在同一條執行緒,Word 才能穩定重複使用
    app.run(host="127.0.0.1", port=PORT, threaded=False)
