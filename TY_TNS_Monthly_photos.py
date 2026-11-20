# -*- coding: utf-8 -*-
"""
Tree Audit Report - General View Version
Part 1: General View photos
Part 2: Tree-specific photos (optional)
If no Part 2 -> Findings = No general comment...
"""
import os
import sys
import re
import tempfile
import zipfile
import shutil
import webbrowser
from pathlib import Path
from io import BytesIO
from collections import defaultdict
from datetime import datetime

from flask import (
    Flask, render_template_string, request, send_file, redirect, url_for, flash
)
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image

app = Flask(__name__)
app.secret_key = "tree-audit-general-2026"
APP_VERSION = "2026-08-13-general-v1"
print(f"TY_TNS_Monthly_photos R1 {APP_VERSION}")

UPLOAD_FOLDER = Path(tempfile.gettempdir()) / "tree_audit_general_uploads"
UPLOAD_FOLDER.mkdir(exist_ok=True)

OFFICERS = [
    "Davy Chan AC",
    "Fung Cheng AS",
    "Belmont Chong AS",
    "Felix Cheung AI",
    "Lynn Yuen AI",
    "Chun Ngo AI",
    "Kay Hai AI",
    "Pan Tam AI",
]

# Estate Code -> (English name, Chinese name, District)
ESTATES = [
    ("FHT", "Hoi Tai Factory Estate", "開泰工廠大廈", "Tuen Mun"),
    ("CLW", "Cheung Lung Wai Estate", "祥龍圍邨", "North"),
    ("CWA", "Cheung Wah Estate", "祥華邨", "North"),
    ("CHH", "Ching Ho Estate", "清河邨", "North"),
    ("CP", "Choi Po Court", "彩蒲苑", "North"),
    ("CY", "Choi Yuen Estate", "彩園邨", "North"),
    ("FM", "Fai Ming Estate", "暉明邨", "North"),
    ("KAF", "Ka Fuk Estate", "嘉福邨", "North"),
    ("PSW", "Po Shek Wu Estate", "寶石湖邨", "North"),
    ("TPG", "Tin Ping Estate", "天平邨", "North"),
    ("WS", "Wah Sum Estate", "華心邨", "North"),
    ("YSG", "Yung Shing Court", "雍盛苑", "North"),
    ("QH", "Queens Hill Estate", "皇后山邨", "North"),
    ("CNS", "Chun Shek Estate", "秦石邨", "Sha Tin"),
    ("CNY", "Chun Yeung Estate", "駿洋邨", "Sha Tin"),
    ("CHO", "Chung On Estate", "頌安邨", "Sha Tin"),
    ("FUW", "Fung Wo Estate", "豐和邨", "Sha Tin"),
    ("HO", "Heng On Estate", "恆安邨", "Sha Tin"),
    ("HYU", "Hin Yiu Estate", "顯耀邨", "Sha Tin"),
    ("HKFU", "Kam Fung Court", "錦豐苑", "Sha Tin"),
    ("LEO", "Lee On Estate", "利安邨", "Sha Tin"),
    ("LY", "Lek Yuen Estate", "瀝源邨", "Sha Tin"),
    ("LH", "Lung Hang Estate", "隆亨邨", "Sha Tin"),
    ("ML", "Mei Lam Estate", "美林邨", "Sha Tin"),
    ("MTI", "Mei Tin Estate", "美田邨", "Sha Tin"),
    ("PH", "Pok Hong Estate", "博康邨", "Sha Tin"),
    ("SK", "Sha Kok Estate", "沙角邨", "Sha Tin"),
    ("SM", "Shek Mun Estate", "碩門邨", "Sha Tin"),
    ("SCO", "Shui Chuen O Estate", "水泉澳邨", "Sha Tin"),
    ("HSW", "Sui Wo Court", "穗禾苑", "Sha Tin"),
    ("SC", "Sun Chui Estate", "新翠邨", "Sha Tin"),
    ("STW", "Sun Tin Wai Estate", "新田圍邨", "Sha Tin"),
    ("WC", "Wo Che Estate", "禾輋邨", "Sha Tin"),
    ("YAO", "Yan On Estate", "欣安邨", "Sha Tin"),
    ("FH", "Fu Heng Estate", "富亨邨", "Tai Po"),
    ("FSH", "Fu Shin Estate", "富善邨", "Tai Po"),
    ("FTP", "Fu Tip Estate", "富蝶邨", "Tai Po"),
    ("KFK", "Kwong Fuk Estate", "廣福邨", "Tai Po"),
    ("POH", "Po Heung Estate", "寶鄉邨", "Tai Po"),
    ("TWO", "Tai Wo Estate", "太和邨", "Tai Po"),
    ("TY", "Tai Yuen Estate", "大元邨", "Tai Po"),
    ("WTT", "Wan Tau Tong Estate", "運頭塘邨", "Tai Po"),
    ("BE", "Butterfly Estate", "蝴蝶邨", "Tuen Mun"),
    ("CTI", "Ching Tin Estate", "菁田邨", "Tuen Mun"),
    ("FTA", "Fu Tai Estate", "富泰邨", "Tuen Mun"),
    ("KSG", "Kin Sang Estate", "建生邨", "Tuen Mun"),
    ("LKG", "Leung King Estate", "良景邨", "Tuen Mun"),
    ("LKY", "Lung Yat Estate", "龍逸邨", "Tuen Mun"),
    ("OT", "On Ting Estate", "安定邨", "Tuen Mun"),
    ("POT", "Po Tin Estate", "寶田邨", "Tuen Mun"),
    ("SS", "Sam Shing Estate", "三聖邨", "Tuen Mun"),
    ("TH", "Tai Hing Estate", "大興邨", "Tuen Mun"),
    ("TK", "Tin King Estate", "田景邨", "Tuen Mun"),
    ("WOT", "Wo Tin Estate", "和田邨", "Tuen Mun"),
    ("WK", "Wu King Estate", "湖景邨", "Tuen Mun"),
    ("YO", "Yau Oi Estate", "友愛邨", "Tuen Mun"),
    ("YNT", "Yan Tin Estate", "欣田邨", "Tuen Mun"),
    ("HFU", "Hung Fuk Estate", "洪福邨", "Yuen Long"),
    ("LCG", "Long Ching Estate", "朗晴邨", "Yuen Long"),
    ("LSN", "Long Shin Estate", "朗善邨", "Yuen Long"),
    ("SPWA", "Shui Pin Wai Estate", "水邊圍邨", "Yuen Long"),
    ("TCK", "Tin Chak Estate", "天澤邨", "Yuen Long"),
    ("TCG", "Tin Ching Estate", "天晴邨", "Yuen Long"),
    ("TNH", "Tin Heng Estate", "天恒邨", "Yuen Long"),
    ("TS1", "Tin Shui (1) Estate", "天瑞(一)邨", "Yuen Long"),
    ("TS2", "Tin Shui (2) Estate", "天瑞(二)邨", "Yuen Long"),
    ("TTZ", "Tin Tsz Estate", "天慈邨", "Yuen Long"),
    ("TWA", "Tin Wah Estate", "天華邨", "Yuen Long"),
    ("TIA", "Tin Yan Estate", "天恩邨", "Yuen Long"),
    ("TYA", "Tin Yat Estate", "天逸邨", "Yuen Long"),
    ("TNY1", "Tin Yiu (1) Estate", "天耀(一)邨", "Yuen Long"),
    ("TNY2", "Tin Yiu (2) Estate", "天耀(二)邨", "Yuen Long"),
    ("TYT", "Tin Yuet Estate", "天悅邨", "Yuen Long"),
    ("SWA", "San Wai Court", "新圍苑", "Tuen Mun"),
]
ESTATE_BY_CODE = {e[0].upper(): e for e in ESTATES}

OFFICER_EMAIL = {
    "Davy Chan AC": "Davy.chan@baguio.com.hk",
    "Davy Chan": "Davy.chan@baguio.com.hk",
    "Belmont Chong AS": "belmont.chong@baguio.com.hk",
    "Fung Cheng AS": "fung.cheng@baguio.com.hk",
    "Felix Cheung AI": "felix.cheung@baguio.com.hk",
    "Lynn Yuen AI": "lynn.yuen@baguio.com.hk",
    "Chun Ngo AI": "chun.ngo@baguio.com.hk",
    "Kay Hai AI": "kay.hai@baguio.com.hk",
    "Pan Tam AI": "pan.tam@baguio.com.hk",
}

SIGNATURE_MAP = {
    "Belmont Chong AS": "belmont_chong.png",
    "Davy Chan AC": "davy_chan.png",
    "Davy Chan": "davy_chan.png",
    "Felix Cheung AI": "felix_cheung.png",
    "Lynn Yuen AI": "lynn_yuen.png",
    "Pan Tam AI": "pan_tam.png",
    "Kay Hai AI": "kay_hai.png",
    "Fung Cheng AS": "fung_cheng.png",
    "Cheng Fung AS": "fung_cheng.png",
}

NO_FINDING_TEXT = (
    "No general comment on tree works maintenance and horticultural "
    "reminders after regular audit inspection."
)

HTML = r"""
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>TY_TNS_Monthly_photos R1</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Segoe UI", Arial, sans-serif; background: #f0f4f8; color: #222; }
  .header { background: #2c5f2d; color: #fff; padding: 18px 24px; }
  .header h1 { font-size: 1.4rem; }
  .header p { opacity: 0.85; font-size: 0.9rem; margin-top: 4px; }
  .wrap { max-width: 860px; margin: 24px auto; padding: 0 16px; }
  .card { background: #fff; border-radius: 10px; padding: 22px; margin-bottom: 18px;
          box-shadow: 0 2px 8px rgba(0,0,0,.08); }
  .card h2 { font-size: 1.1rem; color: #2c5f2d; margin-bottom: 14px; border-bottom: 2px solid #e8f0e8; padding-bottom: 8px; }
  .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
  label { display: block; font-size: 0.82rem; color: #555; margin-bottom: 3px; }
  input, select, textarea { width: 100%; padding: 8px 10px; border: 1px solid #ccc;
    border-radius: 6px; font-size: 0.95rem; }
  input:focus, select:focus, textarea:focus { outline: none; border-color: #2c5f2d; }
  .drop { border: 2px dashed #a8c5a8; border-radius: 8px; padding: 20px; text-align: center;
    background: #f7fbf7; cursor: pointer; }
  .drop:hover { border-color: #2c5f2d; background: #eef6ee; }
  .drop input { display: none; }
  .hint { font-size: 0.8rem; color: #666; margin-top: 8px; line-height: 1.5; }
  .btn { background: #2c5f2d; color: #fff; border: none; padding: 12px 28px;
    border-radius: 8px; font-size: 1rem; cursor: pointer; width: 100%; margin-top: 8px; }
  .btn:hover { background: #1e4620; }
  .btn:disabled { background: #999; cursor: wait; }
  .flash { padding: 12px 16px; border-radius: 8px; margin-bottom: 14px; }
  .flash.error { background: #fde8e8; color: #9b1c1c; }
  .flash.ok { background: #e6f4ea; color: #1e4620; }
  .part-badge { display: inline-block; background: #2c5f2d; color: #fff; font-size: 0.75rem;
    padding: 2px 8px; border-radius: 4px; margin-right: 6px; }
</style>
</head>
<body>
<div class="header">
  <h1>TY_TNS Monthly Photos</h1>
  <p>TY / TNS Region (North / Sha Tin / Tai Po / Tuen Mun / Yuen Long)</p>
</div>
<div class="wrap">
  {% with messages = get_flashed_messages(with_categories=true) %}
    {% for cat, msg in messages %}
      <div class="flash {{ cat }}">{{ msg }}</div>
    {% endfor %}
  {% endwith %}

  <form method="post" action="{{ url_for('generate') }}" enctype="multipart/form-data" id="mainForm">
    <div class="card">
      <h2>① 報告基本資料</h2>
      <div class="grid">
        <div>
          <label>個案編號 (Case No.)</label>
          <input type="text" name="case_no" value="{{ form.case_no }}">
        </div>
        <div>
          <label>日期</label>
          <input type="date" name="date" value="{{ form.date }}">
        </div>
        <div>
          <label>巡查人員 (Inspection Officer)</label>
          <select name="officer" id="officerSelect" required onchange="updateOfficerEmail()">
            {% for o in officers %}
            <option value="{{ o }}" {% if o == form.officer %}selected{% endif %}>{{ o }}</option>
            {% endfor %}
          </select>
        </div>
        <div>
          <label>Email</label>
          <input type="text" name="email" value="{{ form.email }}">
        </div>
        <div>
          <label>發件公司 (From)</label>
          <input type="text" name="from_" value="{{ form.from_ }}">
        </div>
        <div>
          <label>收件公司 (To)</label>
          <input type="text" name="to" value="{{ form.to }}">
        </div>
        <div>
          <label>副本 (cc)</label>
          <input type="text" name="cc" value="{{ form.cc }}">
        </div>
        <div>
          <label>收件 Email</label>
          <input type="text" name="to_email" value="{{ form.to_email }}">
        </div>
        <div>
          <label>合約編號</label>
          <input type="text" name="contract_no" value="{{ form.contract_no }}">
        </div>
        <div>
          <label>屋邨 (Estate)</label>
          <select name="estate_code" id="estateSelect" onchange="updateEstateFields()">
            {% for code, eng, chi, dist in estates %}
            <option value="{{ code }}" data-eng="{{ eng }}" data-chi="{{ chi }}" data-dist="{{ dist }}"
              {% if code == form.estate_code %}selected{% endif %}>{{ code }} - {{ eng }} {{ chi }}</option>
            {% endfor %}
          </select>
          <input type="hidden" name="estate" id="estateName" value="{{ form.estate }}">
        </div>
        <div>
          <label>地點 (Location)</label>
          <input type="text" name="location" id="locationField" value="{{ form.location }}">
        </div>
        <div>
          <label>Tree Information</label>
          <input type="text" name="tree_info" value="{{ form.tree_info }}">
        </div>
      </div>
    </div>


    <div class="card">
      <h2>相片日期／時間戳（印在相片右下角，像相機時間）</h2>
      <p class="hint" style="margin-bottom:10px">勾選後，日期時間會直接印在上傳的相片上（橙色字，右下角）。</p>
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px">
        <input type="checkbox" name="add_datetime" id="addDt" value="1">
        <label for="addDt" style="margin:0;font-weight:600">在相片上加上日期／時間</label>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">
        <div>
          <label>日期來源</label>
          <select name="date_source">
            <option value="exif" selected>相片拍攝日期（EXIF）</option>
            <option value="form">自訂日期（用上方 Date）</option>
          </select>
        </div>
        <div>
          <label>時間來源</label>
          <select name="time_source">
            <option value="exif" selected>相片拍攝時間（EXIF）</option>
            <option value="none">不顯示時間</option>
            <option value="custom">自訂時間</option>
            <option value="now">現在時間</option>
          </select>
        </div>
        <div>
          <label>自訂時間（選「自訂時間」時填，例如 14:30）</label>
          <input type="text" name="custom_time" placeholder="HH:MM" value="">
        </div>
        <div>
          <label>字體大小</label>
          <select name="dt_size">
            <option value="50">50</option>
            <option value="80" selected>80</option>
            <option value="100">100</option>
            <option value="120">120</option>
          </select>
        </div>
      </div>
    </div>

    <!-- Part 1 removed for TY/TNS -->

    </div>

    <div class="card">
      <h2><span class="part-badge">Part 2</span> 樹木相片（可選）</h2>
      <div class="drop" id="treeDrop"
           onclick="document.getElementById('treeFiles').click()"
           ondragover="event.preventDefault(); this.style.borderColor='#2c5f2d';"
           ondragleave="this.style.borderColor='#a8c5a8';"
           ondrop="handleDrop(event, 'treeFiles', 'treeList')">
        <p style="font-size:1.1rem;margin-bottom:6px">📁 點擊選擇 或 拖入相片／資料夾</p>
        <p class="hint">檔名例如：TC-T0951_Dead fronds.jpg · 可直接把 folder 拖進來</p>
        <input type="file" name="tree_photos" id="treeFiles" accept="image/*" multiple
               onchange="showNames('treeFiles','treeList')">
      </div>
      <button type="button" class="btn" style="background:#5a7a5a;margin-top:8px;padding:8px;"
              onclick="document.getElementById('treeFolder').click()">📂 選擇整個資料夾</button>
      <input type="file" id="treeFolder" webkitdirectory directory multiple
             style="display:none" onchange="folderToInput(this, 'treeFiles', 'treeList')">
      <div id="treeList" class="hint"></div>
      <p class="hint" style="margin-top:10px">
        ✅ 有樹木相片 → Findings 自動列出各樹問題<br>
        ❌ 沒有樹木相片 → Findings 寫成「No general comment on tree works maintenance...」
      </p>
    </div>

    <button type="submit" class="btn" id="submitBtn">產生 Word 報告</button>
  </form>
</div>
<script>
  const officerEmails = {
    "Davy Chan AC": "Davy.chan@baguio.com.hk",
    "Belmont Chong AS": "belmont.chong@baguio.com.hk",
    "Fung Cheng AS": "fung.cheng@baguio.com.hk",
    "Felix Cheung AI": "felix.cheung@baguio.com.hk",
    "Lynn Yuen AI": "lynn.yuen@baguio.com.hk",
    "Chun Ngo AI": "chun.ngo@baguio.com.hk",
    "Kay Hai AI": "kay.hai@baguio.com.hk",
    "Pan Tam AI": "pan.tam@baguio.com.hk"
  };
  function updateOfficerEmail() {
    const o = document.getElementById('officerSelect').value;
    const f = document.querySelector('input[name="email"]');
    if (f && officerEmails[o]) f.value = officerEmails[o];
  }
  function updateEstateFields() {
    const sel = document.getElementById('estateSelect');
    const opt = sel.options[sel.selectedIndex];
    const eng = opt.getAttribute('data-eng') || '';
    const dist = opt.getAttribute('data-dist') || '';
    document.getElementById('estateName').value = eng + ' / ' + dist;
    document.getElementById('locationField').value = eng;
  }
  updateEstateFields();
  function showNames(inputId, listId) {
    const files = document.getElementById(inputId).files;
    const el = document.getElementById(listId);
    if (!files.length) { el.innerHTML = ''; return; }
    el.innerHTML = '已選 ' + files.length + ' 張：' +
      Array.from(files).map(f => f.name).join('、');
  }
  function folderToInput(folderInput, targetId, listId) {
    const dt = new DataTransfer();
    const imgs = Array.from(folderInput.files).filter(f =>
      /\.(jpe?g|png|gif|bmp|webp|tif{1,2})$/i.test(f.name));
    imgs.forEach(f => dt.items.add(f));
    document.getElementById(targetId).files = dt.files;
    showNames(targetId, listId);
  }
  function handleDrop(ev, targetId, listId) {
    ev.preventDefault();
    ev.currentTarget.style.borderColor = '#a8c5a8';
    const dt = new DataTransfer();
    const existing = document.getElementById(targetId).files;
    for (const f of existing) dt.items.add(f);
    const items = ev.dataTransfer.items;
    let pending = 0;
    const done = () => {
      document.getElementById(targetId).files = dt.files;
      showNames(targetId, listId);
    };
    if (!items) {
      for (const f of ev.dataTransfer.files) {
        if (/\.(jpe?g|png|gif|bmp|webp)$/i.test(f.name)) dt.items.add(f);
      }
      done(); return;
    }
    for (const item of items) {
      const entry = item.webkitGetAsEntry && item.webkitGetAsEntry();
      if (entry && entry.isDirectory) {
        pending++;
        readDir(entry, dt, () => { pending--; if (pending<=0) done(); });
      } else if (item.kind === 'file') {
        const f = item.getAsFile();
        if (f && /\.(jpe?g|png|gif|bmp|webp)$/i.test(f.name)) dt.items.add(f);
      }
    }
    if (pending === 0) done();
  }
  function readDir(dirEntry, dt, cb) {
    const reader = dirEntry.createReader();
    const read = () => {
      reader.readEntries(entries => {
        if (!entries.length) { cb(); return; }
        let left = entries.length;
        entries.forEach(ent => {
          if (ent.isFile) {
            ent.file(f => {
              if (/\.(jpe?g|png|gif|bmp|webp)$/i.test(f.name)) dt.items.add(f);
              left--; if (left<=0) read();
            });
          } else if (ent.isDirectory) {
            readDir(ent, dt, () => { left--; if (left<=0) read(); });
          } else {
            left--; if (left<=0) read();
          }
        });
      });
    };
    read();
  }
  document.getElementById('mainForm').addEventListener('submit', function() {
    const btn = document.getElementById('submitBtn');
    btn.disabled = true;
    btn.innerHTML = '⏳ 正在產生報告...';
  });
</script>
</body>
</html>
"""


def set_cell_borders(cell, border_color="000000", border_size="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), border_size)
        el.set(qn("w:color"), border_color)
        tcBorders.append(el)
    tcPr.append(tcBorders)



def set_cell_margins(cell, top=40, bottom=40, left=40, right=40):
    """Inner cell padding in twips (40 = 0.07cm) so photo stays inside border."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    # replace existing
    for old in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(old)
    tcPr.append(tcMar)


def set_run_font(run, font_name="Arial", font_size=11, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    run.font.size = Pt(font_size)
    run.font.bold = bold


def apply_exif_orientation(img):
    """Rotate according to camera EXIF so portrait/landscape match original."""
    try:
        exif = img._getexif()
        if not exif:
            return img
        o = exif.get(274) or exif.get(0x0112)
        if o == 2:
            img = img.transpose(Image.FLIP_LEFT_RIGHT)
        elif o == 3:
            img = img.transpose(Image.ROTATE_180)
        elif o == 4:
            img = img.transpose(Image.FLIP_TOP_BOTTOM)
        elif o == 5:
            img = img.transpose(Image.FLIP_LEFT_RIGHT).transpose(Image.ROTATE_90)
        elif o == 6:
            img = img.transpose(Image.ROTATE_270)
        elif o == 7:
            img = img.transpose(Image.FLIP_LEFT_RIGHT).transpose(Image.ROTATE_270)
        elif o == 8:
            img = img.transpose(Image.ROTATE_90)
    except Exception:
        pass
    return img


def get_exif_datetime(path):
    """Return 'YYYY-MM-DD HH:MM' from EXIF or empty string."""
    try:
        img = Image.open(path)
        exif = img._getexif() or {}
        from PIL import ExifTags
        for tag, val in exif.items():
            name = ExifTags.TAGS.get(tag, tag)
            if name in ("DateTimeOriginal", "DateTime"):
                try:
                    from datetime import datetime as _dt
                    d = _dt.strptime(val, "%Y:%m:%d %H:%M:%S")
                    return d.strftime("%Y-%m-%d %H:%M")
                except Exception:
                    return str(val).replace(":", "-", 2)
    except Exception:
        pass
    return ""


def get_stamp_font(size):
    for fp in [
        "C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/ARIAL.TTF",
        "C:/Windows/Fonts/calibri.ttf", "C:/Windows/Fonts/segoeui.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "arial.ttf",
    ]:
        try:
            from PIL import ImageFont
            return ImageFont.truetype(fp, size)
        except Exception:
            continue
    from PIL import ImageFont
    return ImageFont.load_default()


def stamp_on_image(img, text, font_pt=80):
    if not text:
        return img
    from PIL import ImageDraw, ImageFont
    img = img.convert("RGB")
    draw = ImageDraw.Draw(img)
    w, h = img.size
    try:
        font_pt = int(font_pt)
    except Exception:
        font_pt = 80
    size = max(24, font_pt)
    font = get_stamp_font(size)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    margin = max(8, int(min(w, h) * 0.02))
    x = max(0, w - tw - margin)
    y = max(0, h - th - margin)
    draw.text((x, y), text, fill=(255, 140, 0), font=font)
    return img


def build_stamp_text(path, opts):
    if not opts or not opts.get("add_dt"):
        return ""
    form_date = (opts.get("form_date") or "").replace("/", "-")
    date_source = opts.get("date_source", "exif")
    time_source = opts.get("time_source", "none")
    custom_t = opts.get("custom_time") or ""

    exif_full = get_exif_datetime(path)
    if date_source == "exif" and exif_full:
        date_part = exif_full.split()[0]
    else:
        date_part = form_date or (exif_full.split()[0] if exif_full else "")

    if not date_part:
        return ""

    if time_source == "none" or not time_source:
        return date_part
    if time_source == "exif":
        if exif_full and " " in exif_full:
            return f"{date_part} {exif_full.split()[1]}"
        return date_part
    if time_source == "custom":
        t = custom_t or ""
        return f"{date_part} {t}".strip() if t else date_part
    if time_source == "now":
        from datetime import datetime as _dt
        return f"{date_part} {_dt.now().strftime('%H:%M')}"
    return date_part


def compress_image(src_path, max_side=1400, quality=70, stamp_text="", dt_size=80):
    img = Image.open(src_path)
    img = apply_exif_orientation(img)
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    w, h = img.size
    scale = min(1.0, max_side / max(w, h))
    if scale < 1.0:
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    if stamp_text:
        img = stamp_on_image(img, stamp_text, font_pt=dt_size)
    buf = BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    buf.seek(0)
    return buf, img.size[0], img.size[1]


def parse_tree_filename(filename):
    """Parse various formats:
    SL1-T0391 (2)_Dead branch.JPG
    TC-T0951_Dead fronds.jpg
    TL-T1656 (1).jpg
    """
    name = Path(filename).stem.strip()

    
    # Pattern AT: AT1, AT1_Wilting, AT_Wilting, AT2 (1)_something
    m = re.match(
        r"^AT([\s_\-]*\d+)?[\s_\-]*(.*)$",
        name, re.IGNORECASE,
    )
    if m and name.upper().startswith("AT"):
        num_part = (m.group(1) or "").strip(" _-")
        rest = (m.group(2) or "").strip(" _-")
        num = int(num_part) if num_part.isdigit() else 0
        label = f"AT{num}" if num else "AT"
        finding = rest or label
        return label, finding, num or 0

# Pattern A: EstateCode-T#### (n)_Finding  e.g. SL1-T0391 (2)_Dead branch
    m = re.match(
        r"^((?:[A-Z]{1,6}\d?)-T\d+)\s*(?:[\(\[]\s*(\d+)\s*[\)\]])?\s*[_\-\s]*(.*)$",
        name, re.IGNORECASE,
    )
    if m:
        tree_id = m.group(1).upper()
        idx = int(m.group(2) or 0)
        rest = (m.group(3) or "").strip(" _-")
        m2 = re.match(r"^(.+?)\s*[\(\[]\s*(\d+)\s*[\)\]]\s*$", rest)
        if m2:
            rest, idx = m2.group(1).strip(), int(m2.group(2)) or idx
        m3 = re.match(r"^(.+?)\s*#\s*(\d+)\s*$", rest)
        if m3:
            rest, idx = m3.group(1).strip(), int(m3.group(2)) or idx
        finding = rest.strip(" _-#")
        if finding and not re.search(r"[A-Za-z\u4e00-\u9fff]", finding):
            finding = ""
        return tree_id, finding, idx

    # Pattern B: T#### without estate prefix
    m = re.match(
        r"^(T\d+)\s*(?:[\(\[]\s*(\d+)\s*[\)\]])?\s*[_\-\s]*(.*)$",
        name, re.IGNORECASE,
    )
    if m:
        tree_id = m.group(1).upper()
        idx = int(m.group(2) or 0)
        rest = (m.group(3) or "").strip(" _-")
        finding = rest.strip(" _-#")
        if finding and not re.search(r"[A-Za-z\u4e00-\u9fff]", finding):
            finding = ""
        return tree_id, finding, idx

    return None


def group_tree_photos(file_list):
    groups = defaultdict(lambda: {"finding": "", "photos": []})
    for orig_name, path in file_list:
        parsed = parse_tree_filename(orig_name)
        if not parsed:
            print(f"  [skip] {orig_name}")
            continue
        tid, finding, idx = parsed
        if finding:
            # prefer longer descriptive finding
            if len(finding) >= len(groups[tid]["finding"]):
                groups[tid]["finding"] = finding
        elif not groups[tid]["finding"]:
            groups[tid]["finding"] = "Photo"
        groups[tid]["photos"].append((idx, path, finding))
    for tid in groups:
        groups[tid]["photos"].sort(key=lambda x: (x[0], str(x[1])))
    return groups


def add_photo_cell(cell, path, max_width_cm=6.4, max_height_cm=5.0, stamp_opts=None):
    """Insert photo keeping orientation; smaller than cell so it stays inside border."""
    set_cell_borders(cell, border_color="000000", border_size="8")
    set_cell_margins(cell, 60, 60, 60, 60)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    try:
        st = build_stamp_text(path, stamp_opts) if stamp_opts else ""
        dt_size = (stamp_opts or {}).get("dt_size", 80)
        buf, px_w, px_h = compress_image(path, stamp_text=st, dt_size=dt_size)
        aspect = px_w / max(px_h, 1)
        box_aspect = max_width_cm / max_height_cm
        if aspect >= box_aspect:
            use_w, use_h = max_width_cm, max_width_cm / aspect
        else:
            use_h, use_w = max_height_cm, max_height_cm * aspect
        run = p.add_run()
        run.add_picture(buf, width=Cm(use_w), height=Cm(use_h))
    except Exception as e:
        p.add_run(f"[Error: {e}]")


def add_tree_photo_cell(cell, path, stamp_opts=None):
    """Part 2 tree photos: fit in 7cm x 9.33cm box, keep orientation"""
    add_photo_cell(cell, path, max_width_cm=6.6, max_height_cm=8.8, stamp_opts=stamp_opts)


def find_signature(officer):
    sig_file = SIGNATURE_MAP.get(officer, "")
    if not sig_file:
        # try partial match
        for k, v in SIGNATURE_MAP.items():
            if officer.lower().startswith(k.split()[0].lower()):
                sig_file = v
                break
    if not sig_file:
        return None
    candidates = []
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            candidates.append(Path(meipass) / "signatures" / sig_file)
        candidates.append(Path(sys.executable).parent / "signatures" / sig_file)
    candidates += [
        Path(__file__).resolve().parent / "signatures" / sig_file,
        Path.cwd() / "signatures" / sig_file,
    ]
    for c in candidates:
        if c.exists():
            return c
    return None



def ensure_docx_template():
    """Make sure python-docx can find default.docx when frozen."""
    try:
        import docx
        from pathlib import Path
        import sys, shutil
        pkg = Path(docx.__file__).parent / "templates" / "default.docx"
        if pkg.exists():
            return
        # frozen: look next to exe or in _MEIPASS
        bases = []
        if getattr(sys, "frozen", False):
            bases.append(Path(sys._MEIPASS))
            bases.append(Path(sys.executable).parent)
        bases.append(Path(__file__).resolve().parent)
        for b in bases:
            cand = b / "docx" / "templates" / "default.docx"
            if cand.exists():
                pkg.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(cand, pkg)
                return
    except Exception:
        pass


def make_blank_docx():
    """Create a minimal valid blank docx in temp so we never need package default.docx."""
    import zipfile
    from io import BytesIO
    from pathlib import Path
    import tempfile

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t></w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>
    </w:sectPr>
  </w:body>
</w:document>"""

    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document_xml)
    path = Path(tempfile.gettempdir()) / "blank_audit_template.docx"
    path.write_bytes(buf.getvalue())
    return str(path)

def generate_docx(meta, general_photos, tree_groups, stamp_opts=None):
    doc = Document(make_blank_docx())
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    # === Header table ===
    table = doc.add_table(rows=5, cols=4)
    table.cell(0, 0).merge(table.cell(0, 3))
    table.cell(0, 0).text = "Audit Report on Routine Maintenance / Completion Report"
    for run in table.cell(0, 0).paragraphs[0].runs:
        set_run_font(run, bold=True, font_size=12)
    table.cell(0, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    table.cell(1, 0).text = "From:"
    table.cell(1, 1).text = meta["from_"]
    table.cell(1, 2).text = "To:"
    table.cell(1, 3).text = meta["to"]

    table.cell(2, 0).text = "Inspection\nOfficer:"
    # Officer display like "Pan Tam / AI"
    officer_display = meta["officer"]
    table.cell(2, 1).text = officer_display
    table.cell(2, 2).text = "cc:"
    table.cell(2, 3).text = meta["cc"]

    table.cell(3, 0).text = "Email:"
    table.cell(3, 1).text = meta["email"]
    table.cell(3, 2).text = "Email:"
    table.cell(3, 3).text = meta.get("to_email", "")

    table.cell(4, 0).text = "Date:"
    table.cell(4, 1).text = meta["date"]
    table.cell(4, 2).text = "Estate/\nDistrict:"
    table.cell(4, 3).text = meta["estate"]

    for row in table.rows:
        for cell in row.cells:
            set_cell_borders(cell)
            for p in cell.paragraphs:
                for run in p.runs:
                    set_run_font(run, font_size=10)

    doc.add_paragraph()

    # === Info table ===
    info = doc.add_table(rows=3, cols=2)
    info.cell(0, 0).width = Cm(4)
    info.cell(0, 1).width = Cm(12)
    info.cell(0, 0).text = "Referral\nCase No."
    info.cell(0, 1).text = meta["case_no"] or "N/A"
    info.cell(1, 0).text = "Location:"
    info.cell(1, 1).text = meta["location"]
    info.cell(2, 0).text = "Findings:"

    if tree_groups:
        findings_lines = []
        for tid, ginfo in tree_groups.items():
            findings_lines.append(f"{tid}_{ginfo['finding']}")
        info.cell(2, 1).text = "\n".join(findings_lines)
    else:
        info.cell(2, 1).text = NO_FINDING_TEXT

    for row in info.rows:
        for cell in row.cells:
            set_cell_borders(cell)
            for p in cell.paragraphs:
                for run in p.runs:
                    set_run_font(run, font_size=10)
    for r in range(3):
        for p in info.cell(r, 0).paragraphs:
            for run in p.runs:
                run.bold = True

    doc.add_paragraph()

    # === Photos – General View ===
    if general_photos:
        cap = doc.add_table(rows=1, cols=1)
        set_cell_borders(cap.cell(0, 0))
        p = cap.cell(0, 0).paragraphs[0]
        run = p.add_run("Photos – General View")
        set_run_font(run, bold=True, font_size=11)

        for row_start in range(0, len(general_photos), 2):
            row_paths = general_photos[row_start:row_start + 2]
            img_table = doc.add_table(rows=1, cols=2)
            img_table.autofit = False
            for j in range(2):
                cell = img_table.cell(0, j)
                cell.width = Cm(8.0)
                if j < len(row_paths):
                    add_photo_cell(cell, row_paths[j], stamp_opts=stamp_opts)
                else:
                    set_cell_borders(cell, border_color="FFFFFF")

    # === Photos – each tree (start on new page after General View) ===
    if general_photos and tree_groups:
        p_break = doc.add_paragraph()
        run_break = p_break.add_run()
        run_break.add_break(WD_BREAK.PAGE)

    PHOTO_TABLE_W = 16.0
    # Page 1: header + first tree.
    # From the second tree: if two consecutive trees each have ≤2 photos, put them on the same page.
    # Otherwise start a new page so caption + photos stay together.
    trees_on_current_page = 0
    prev_photos = 0
    prev_n_photos = 0
    for n, (tid, ginfo) in enumerate(tree_groups.items(), 1):
        n_photos = len(ginfo.get("photos") or [])
        if n == 1:
            trees_on_current_page = 1
            prev_photos = n_photos
        else:
            # Always start a new page after the first tree
            if n == 2:
                need_break = True
            else:
                # On later pages: allow two small trees (≤2 photos each) to share a page
                if prev_photos <= 2 and n_photos <= 2 and trees_on_current_page < 2:
                    need_break = False
                else:
                    need_break = True
            if need_break:
                pb = doc.add_paragraph()
                pb.add_run().add_break(WD_BREAK.PAGE)
                trees_on_current_page = 0
            trees_on_current_page += 1
            prev_photos = n_photos

        cap = doc.add_table(rows=1, cols=1)
        cap.autofit = False
        cap.allow_autofit = False
        cap.cell(0, 0).width = Cm(PHOTO_TABLE_W)
        set_cell_borders(cap.cell(0, 0), border_color="000000", border_size="8")
        p = cap.cell(0, 0).paragraphs[0]
        run = p.add_run(f"Photos – {tid}_{ginfo['finding']}")
        set_run_font(run, bold=True, font_size=11)

        triples = ginfo["photos"]
        # Simple left-right pairing for all photos (including AT)
        photo_paths = [item[1] for item in triples]

        def add_pair(left_path, right_path):
            img_table = doc.add_table(rows=1, cols=2)
            img_table.autofit = False
            c0 = img_table.cell(0, 0)
            c1 = img_table.cell(0, 1)
            c0.width = Cm(8.0)
            c1.width = Cm(8.0)
            if left_path:
                add_tree_photo_cell(c0, left_path, stamp_opts=stamp_opts)
            else:
                set_cell_borders(c0, border_color="FFFFFF")
            if right_path:
                add_tree_photo_cell(c1, right_path, stamp_opts=stamp_opts)
            else:
                set_cell_borders(c1, border_color="FFFFFF")

        for ri in range(0, len(photo_paths), 2):
            left = photo_paths[ri]
            right = photo_paths[ri + 1] if ri + 1 < len(photo_paths) else None
            add_pair(left, right)
        gap = doc.add_paragraph()
        gap.paragraph_format.space_before = Pt(10)
        gap.paragraph_format.space_after = Pt(10)

    # === Footer ===
    doc.add_paragraph()
    foot = doc.add_table(rows=4, cols=2)
    foot.cell(0, 0).text = "Audited by:"
    # Short name without rank for signature block
    short_name = meta["officer"].split(" AC")[0].split(" AS")[0].split(" AI")[0].strip()
    foot.cell(0, 1).text = short_name
    foot.cell(1, 0).text = "Signature:"
    sig_cell = foot.cell(1, 1)
    sig_cell.paragraphs[0].clear()
    sig_path = find_signature(meta["officer"])
    if sig_path:
        try:
            run = sig_cell.paragraphs[0].add_run()
            run.add_picture(str(sig_path), width=Inches(1.8))
        except Exception:
            sig_cell.paragraphs[0].add_run("(signed)")
    else:
        sig_cell.paragraphs[0].add_run("(signed)")
    foot.cell(2, 0).text = "AS Contractor:"
    foot.cell(2, 1).text = meta["from_"]
    foot.cell(3, 0).text = "Contract No.:"
    foot.cell(3, 1).text = meta["contract_no"]

    for row in foot.rows:
        for cell in row.cells:
            set_cell_borders(cell)
            for p in cell.paragraphs:
                for run in p.runs:
                    set_run_font(run, font_size=10)
    for r in range(4):
        for p in foot.cell(r, 0).paragraphs:
            for run in p.runs:
                run.bold = True

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


@app.route("/")
def index():
    today = datetime.now().strftime("%Y-%m-%d")
    form = {
        "case_no": "AS-TY-2026-001",
        "date": today,
        "officer": "Pan Tam AI",
        "email": "pan.tam@baguio.com.hk",
        "from_": "Baguio Landscaping Services Limited",
        "to": "Muni Arborist Limited",
        "cc": "TMU",
        "to_email": "",
        "contract_no": "CB20240067",
        "estate_code": "BE",
        "estate": "Butterfly Estate / Tuen Mun",
        "location": "Butterfly Estate",
        "tree_info": "Site audit inspection on estate trees condition (General items)",
    }
    return render_template_string(HTML, form=form, officers=OFFICERS, estates=ESTATES)



def region_abbrev(estate_code, estate_text=""):
    code = (estate_code or "").upper()
    if code in ESTATE_BY_CODE:
        dist = ESTATE_BY_CODE[code][3]
        d = dist.upper()
        if "NORTH" in d or "SHA TIN" in d or "TAI PO" in d:
            return "TNS"
        return "TY"
    txt = (estate_text or "").upper()
    if any(k in txt for k in ("NORTH", "SHA TIN", "TAI PO")):
        return "TNS"
    return "TY"


def generate_referral_xlsx(meta, tree_groups):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Referral"
    headers = [
        "Source of Referral", "Estate", "Case No.",
        "Date Received", "Tree ID", "Observation",
    ]
    ws.append(headers)
    for c in ws[1]:
        c.font = c.font.copy(bold=True)

    loc = (meta.get("location") or "").strip()
    est = (meta.get("estate") or "").strip()
    estate_eng = loc if loc and "/" not in loc else re.split(r"[\u4e00-\u9fff/]", est)[0].strip()
    estate_eng = re.sub(r"\s*/\s*.*$", "", estate_eng).strip() or loc or est
    case_no = meta.get("case_no") or ""
    date_req = (meta.get("date") or "").replace("/", "-")

    if tree_groups:
        for tid, ginfo in tree_groups.items():
            desc = (ginfo.get("finding") or "").strip()
            if desc.lower() in ("photo", tid.lower()):
                desc = ""
            ws.append(["AS", estate_eng, case_no, date_req, tid, desc])
    else:
        ws.append(["AS", estate_eng, case_no, date_req, "", ""])

    for col in ws.columns:
        maxlen = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(maxlen + 2, 16), 42)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf



def generate_pdf(meta, tree_groups, stamp_opts=None):
    """Simple PDF version matching the Word report content."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from io import BytesIO
    from pathlib import Path

    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    story = []

    title = Paragraph("<b>Audit Report on Routine Maintenance / Completion Report</b>", styles["Heading2"])
    story.append(title)
    story.append(Spacer(1, 8))

    data = [
        ["From:", meta.get("from_", ""), "To:", meta.get("to", "")],
        ["Inspection Officer:", meta.get("officer", ""), "cc:", meta.get("cc", "")],
        ["Email:", meta.get("email", ""), "Email:", meta.get("to_email", "")],
        ["Date:", meta.get("date", ""), "Estate/District:", meta.get("estate", "")],
    ]
    t = Table(data, colWidths=[4*cm, 6*cm, 3*cm, 4.5*cm])
    t.setStyle(TableStyle([("GRID", (0,0), (-1,-1), 0.5, colors.black), ("FONTSIZE", (0,0), (-1,-1), 9)]))
    story.append(t)
    story.append(Spacer(1, 8))

    findings = []
    for tid, ginfo in tree_groups.items():
        findings.append(f"{tid}_{ginfo.get('finding','')}")
    info_data = [
        ["Referral Case No.", meta.get("case_no", "")],
        ["Location:", meta.get("location", "")],
        ["Findings:", "\n".join(findings)],
    ]
    it = Table(info_data, colWidths=[4*cm, 13*cm])
    it.setStyle(TableStyle([("GRID", (0,0), (-1,-1), 0.5, colors.black), ("FONTSIZE", (0,0), (-1,-1), 9), ("VALIGN", (0,0), (-1,-1), "TOP")]))
    story.append(it)
    story.append(Spacer(1, 12))

    for n, (tid, ginfo) in enumerate(tree_groups.items(), 1):
        n_photos = len(ginfo.get("photos") or [])
        n_photos = len(ginfo.get("photos") or [])
        if n == 2 or (n > 2 and not (prev_n_photos <= 2 and n_photos <= 2)):
            story.append(PageBreak())
        prev_n_photos = n_photos
        cap = Paragraph(f"<b>Photos – {tid}_{ginfo.get('finding','')}</b>", styles["Heading3"])
        story.append(cap)
        story.append(Spacer(1, 6))
        paths = [item[1] for item in ginfo.get("photos", [])]
        for i in range(0, len(paths), 2):
            row = []
            for j in range(2):
                if i + j < len(paths):
                    try:
                        img = RLImage(str(paths[i+j]), width=7.5*cm, height=6*cm)
                        row.append(img)
                    except Exception:
                        row.append(Paragraph("[image error]", styles["Normal"]))
                else:
                    row.append("")
            pt = Table([row], colWidths=[8*cm, 8*cm])
            pt.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("ALIGN", (0,0), (-1,-1), "CENTER")]))
            story.append(pt)
            story.append(Spacer(1, 6))

    story.append(Spacer(1, 12))
    foot = [
        ["Audited by:", meta.get("officer", "").split(" AC")[0].split(" AS")[0].split(" AI")[0].strip()],
        ["Signature:", ""],
        ["AS Contractor:", meta.get("from_", "")],
        ["Contract No.:", meta.get("contract_no", "")],
    ]
    ft = Table(foot, colWidths=[4*cm, 13*cm])
    ft.setStyle(TableStyle([("GRID", (0,0), (-1,-1), 0.5, colors.black), ("FONTSIZE", (0,0), (-1,-1), 9)]))
    story.append(ft)

    doc.build(story)
    buf.seek(0)
    return buf

@app.route("/generate", methods=["POST"])
def generate():
    stamp_opts = {
        "add_dt": request.form.get("add_datetime") == "1",
        "date_source": request.form.get("date_source", "exif"),
        "time_source": request.form.get("time_source", "none"),
        "form_date": request.form.get("date", "").strip(),
        "custom_time": request.form.get("custom_time", "").strip(),
        "dt_size": int(request.form.get("dt_size", "80") or 80),
    }
    estate_code = request.form.get("estate_code", "").strip().upper()
    estate_name = request.form.get("estate", "").strip()
    location = request.form.get("location", "").strip()
    if estate_code and estate_code in ESTATE_BY_CODE:
        eng, chi, dist = ESTATE_BY_CODE[estate_code][1], ESTATE_BY_CODE[estate_code][2], ESTATE_BY_CODE[estate_code][3]
        if not estate_name:
            estate_name = f"{eng} / {dist}"
        if not location:
            location = eng

    meta = {
        "case_no": request.form.get("case_no", "").strip(),
        "date": request.form.get("date", "").replace("-", "/"),
        "officer": request.form.get("officer", "").strip(),
        "email": request.form.get("email", "").strip(),
        "from_": request.form.get("from_", "").strip(),
        "to": request.form.get("to", "").strip(),
        "cc": request.form.get("cc", "").strip(),
        "to_email": request.form.get("to_email", "").strip(),
        "contract_no": request.form.get("contract_no", "").strip(),
        "estate": estate_name,
        "location": location,
        "tree_info": request.form.get("tree_info", "").strip(),
        "estate_code": estate_code,
    }

    session_dir = UPLOAD_FOLDER / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    session_dir.mkdir(parents=True, exist_ok=True)

    # Part 1: general photos
    general_photos = []
    for f in request.files.getlist("general_photos"):
        if f and f.filename:
            dest = session_dir / f"gen_{f.filename}"
            f.save(str(dest))
            general_photos.append(dest)

    # Part 2: tree photos
    tree_file_list = []
    for f in request.files.getlist("tree_photos"):
        if f and f.filename:
            dest = session_dir / f.filename
            f.save(str(dest))
            tree_file_list.append((f.filename, dest))

    tree_groups = group_tree_photos(tree_file_list) if tree_file_list else {}

    # Auto-detect estate from General View filenames e.g. SL1_General view.jpg
    # Especially useful when Part 2 is empty
    def detect_code_from_name(fname):
        stem = Path(fname).stem.strip()
        # SL1_General view / KSW_General view / TC_xxx
        m = re.match(r"^([A-Za-z]{1,6}\d?)[_\-\s]", stem)
        if m:
            return m.group(1).upper()
        m = re.match(r"^([A-Za-z]{1,6}\d?)$", stem)
        if m:
            return m.group(1).upper()
        return None

    detected = None
    for p in general_photos:
        # original name may be gen_SL1_General view.jpg
        raw = p.name
        if raw.startswith("gen_"):
            raw = raw[4:]
        code = detect_code_from_name(raw)
        if code and code in ESTATE_BY_CODE:
            detected = code
            break
    # Tree photo prefixes take priority (SL2-T0391 -> SL2)
    if tree_groups:
        for tid in tree_groups:
            m = re.match(r"^([A-Z]{1,6}\d?)-T\d+", tid, re.I)
            if m:
                code = m.group(1).upper()
                if code in ESTATE_BY_CODE:
                    detected = code
                    break

    if detected:
        eng, chi, dist = ESTATE_BY_CODE[detected][1], ESTATE_BY_CODE[detected][2], ESTATE_BY_CODE[detected][3]
        meta["estate_code"] = detected
        meta["estate"] = f"{eng} / {dist}"
        meta["location"] = eng
        print(f"Auto-detected estate: {detected} - {eng} / {dist}")

    if not tree_groups:  # TY/TNS: only Part 2 required
        shutil.rmtree(session_dir, ignore_errors=True)
        flash("請至少上傳 General View 或樹木相片", "error")
        return redirect(url_for("index"))

    try:
        buf = generate_docx(meta, [], tree_groups, stamp_opts=stamp_opts)  # no Part 1
        case = meta["case_no"] or "Audit"
        safe_case = re.sub(r'[\\/:*?"<>|]', "_", case)
        # Estate English name for filename
        est_raw = (meta.get("location") or meta.get("estate") or "").split("/")[0].strip()
        safe_est = re.sub(r'[\\/:*?"<>|]', "_", est_raw) or "Estate"
        out_name = f"{safe_case}_{safe_est}_Audit Report on Routine Maintenance.docx"
        docx_path = Path(tempfile.gettempdir()) / out_name
        docx_path.write_bytes(buf.getvalue() if hasattr(buf, "getvalue") else buf.read())
        return send_file(
            str(docx_path),
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        try:
            flash(f"產生失敗: {e}", "error")
        except Exception:
            pass
        return redirect(url_for("index"))


if __name__ == "__main__":
    import socket
    import threading
    import time

    def find_free_port(start=5001, end=5020):
        for p in range(start, end + 1):
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                try:
                    s.bind(("127.0.0.1", p))
                    return p
                except OSError:
                    continue
        return start

    port = find_free_port()
    url = f"http://127.0.0.1:{port}"
    server_ready = threading.Event()

    def start_flask():
        try:
            app.run(host="127.0.0.1", port=port, debug=False, threaded=True, use_reloader=False)
        except Exception as e:
            print("Flask error:", e)

    threading.Thread(target=start_flask, daemon=True).start()

    # Wait longer for the server to start
    for _ in range(60):
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                server_ready.set()
                break
        except OSError:
            time.sleep(0.4)

    if server_ready.is_set():
        try:
            webbrowser.open(url)
        except Exception:
            pass
    else:
        print("Warning: server did not start on", url)
        print("Try closing other programs using port 5001-5010 and run again.")

    # Desktop status window (works in EXE, no extra install)
    try:
        import tkinter as tk
        root = tk.Tk()
        root.title("TY_TNS Monthly Photos")
        root.geometry("460x200")
        root.resizable(False, False)
        tk.Label(root, text="TY_TNS Monthly Photos", font=("Segoe UI", 14, "bold")).pack(pady=(18, 6))
        tk.Label(root, text="瀏覽器已開啟：  " + url, font=("Segoe UI", 10)).pack()
        tk.Label(
            root,
            text="請保持此視窗開啟。\n關閉此視窗即停止程式。\n其他同事請在自己電腦雙擊同一個 EXE。",
            font=("Segoe UI", 10),
            justify="center",
        ).pack(pady=10)
        tk.Button(root, text="開啟網頁", command=lambda: webbrowser.open(url), width=16).pack(pady=4)
        root.protocol("WM_DELETE_WINDOW", root.destroy)
        root.mainloop()
    except Exception:
        print("=" * 50)
        print("  TY_TNS Monthly Photos")
        print("  Open:", url)
        print("  Keep this window open.")
        print("=" * 50)
        try:
            import time
            while True:
                time.sleep(3600)
        except KeyboardInterrupt:
            pass
