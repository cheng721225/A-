import io
import os
from datetime import datetime

# 自動檢查並安裝雲端缺少的套件
import subprocess
import sys


def install_package(package):
  try:
    __import__(package)
  except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])


install_package("docx")
install_package("PIL")
install_package("streamlit")

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from PIL import Image
import streamlit as st

st.set_page_config(
    page_title="建築工地互動式缺失查驗系統", page_icon="🏗️", layout="wide"
)

st.title("🏗️ 建築工地室內缺失查驗系統")
st.markdown(
    "操作流程：**1. 選擇戶別與樓層** ➔ **2. 點擊按鈕新增缺失編號** ➔"
    " **3. 於下方表格填寫細節並上傳照片** ➔ **4. 一鍵匯出 Word 報告**"
)

if "markers" not in st.session_state:
  st.session_state.markers = []
if "defect_details" not in st.session_state:
  st.session_state.defect_details = {}
if "current_unit" not in st.session_state:
  st.session_state.current_unit = "A1戶別"

st.sidebar.header("📌 1. 查驗位置設定")
unit_choice = st.sidebar.selectbox(
    "選擇戶別", ["A1戶別", "A2戶別", "A3戶別", "B1戶別", "B2戶別", "B3戶別"]
)

if unit_choice != st.session_state.current_unit:
  st.session_state.current_unit = unit_choice
  st.session_state.markers = []
  st.session_state.defect_details = {}
  st.rerun()

floor_choice = st.sidebar.selectbox(
    "選擇樓層",
    [
        "2F",
        "3F",
        "4F",
        "5F",
        "6F",
        "7F",
        "8F",
        "9F",
        "10F",
        "11F",
        "12F",
        "13F",
        "14F",
        "15F",
    ],
)

project_info = {
    "project_name": st.sidebar.text_input(
        "工程名稱", value="弘峻草福段A區新建工程"
    ),
    "location": f"{unit_choice} ({floor_choice})",
    "inspector": st.sidebar.text_input("查驗人員", value=""),
    "supervisor": st.sidebar.text_input("施工負責人", value=""),
}

st.sidebar.divider()
if st.sidebar.button("🗑️ 清除所有標註與表格"):
  st.session_state.markers = []
  st.session_state.defect_details = {}
  st.rerun()

st.subheader(f"🗺️ 2. 平面圖檢視與新增區 ({project_info['location']})")

plan_filename = f"{unit_choice.replace('戶別', '')}_plan.jpg"

col_map, col_ctrl = st.columns([2, 1])

with col_map:
  if os.path.exists(plan_filename):
    st.image(
        plan_filename,
        caption=f"{unit_choice} 平面圖",
        use_container_width=True,
    )
  else:
    st.warning(
        f"⚠️ 尚未偵測到 `{unit_choice}` 的平面圖檔案 (`{plan_filename}`)。"
    )

with col_ctrl:
  st.markdown("### 📍 缺失編號管理")
  if st.button("➕ 新增一個缺失編號", type="primary", use_container_width=True):
    next_id = max([m["id"] for m in st.session_state.markers], default=0) + 1
    st.session_state.markers.append({"id": next_id})
    st.session_state.defect_details[next_id] = {
        "space": "牆面",
        "space_custom": "",
        "content": "",
        "vendor": "油漆",
        "vendor_custom": "",
        "photo": None,
        "remark": "",
    }
    st.rerun()

  st.markdown(f"**目前已建立缺失數：** `{len(st.session_state.markers)}` 處")

  if st.session_state.markers:
    st.markdown("---")
    del_target = st.selectbox(
        "選擇要刪除的編號", [m["id"] for m in st.session_state.markers]
    )
    if st.button("🗑️ 刪除選定編號", type="secondary"):
      st.session_state.markers = [
          x for x in st.session_state.markers if x["id"] != del_target
      ]
      if del_target in st.session_state.defect_details:
        del st.session_state.defect_details[del_target]
      st.rerun()

st.divider()

st.subheader("📋 3. 缺失查驗明細填報表")

if not st.session_state.markers:
  st.info("👈 請點擊上方右側的「新增一個缺失編號」按鈕來建立填報欄位！")
else:
  with st.form("defect_table_form"):
    for m in st.session_state.markers:
      m_id = m["id"]
      if m_id not in st.session_state.defect_details:
        st.session_state.defect_details[m_id] = {
            "space": "牆面",
            "space_custom": "",
            "content": "",
            "vendor": "油漆",
            "vendor_custom": "",
            "photo": None,
            "remark": "",
        }

      detail = st.session_state.defect_details[m_id]

      st.markdown(f"#### 🔍 缺失編號：【 #{m_id} 】")
      c1, c2, c3 = st.columns(3)

      with c1:
        space_options = [
            "天花板",
            "牆面",
            "地板",
            "廚具",
            "衛浴",
            "開關插座",
            "給排水",
            "自訂",
        ]
        curr_space = (
            detail["space"]
            if detail["space"] in space_options
            else "自訂"
        )
        space_sel = st.selectbox(
            f"缺失位置 #{m_id}",
            space_options,
            index=space_options.index(curr_space),
            key=f"space_{m_id}",
        )
        if space_sel == "自訂":
          detail["space_custom"] = st.text_input(
              f"手動輸入位置 #{m_id}",
              value=detail.get("space_custom", ""),
              key=f"space_c_{m_id}",
          )
          detail["space"] = detail["space_custom"]
        else:
          detail["space"] = space_sel

      with c2:
        vendor_options = [
            "水電",
            "泥作",
            "磁磚",
            "油漆",
            "點工",
            "衛浴",
            "廚具",
            "自訂",
        ]
        curr_vendor = (
            detail["vendor"]
            if detail["vendor"] in vendor_options
            else "自訂"
        )
        vendor_sel = st.selectbox(
            f"缺失廠商 #{m_id}",
            vendor_options,
            index=vendor_options.index(curr_vendor),
            key=f"vendor_{m_id}",
        )
        if vendor_sel == "自訂":
          detail["vendor_custom"] = st.text_input(
              f"手動輸入廠商 #{m_id}",
              value=detail.get("vendor_custom", ""),
              key=f"vendor_c_{m_id}",
          )
          detail["vendor"] = detail["vendor_custom"]
        else:
          detail["vendor"] = vendor_sel

      with c3:
        detail["photo"] = st.file_uploader(
            f"缺失照片 #{m_id} (支援手機拍照)",
            type=["jpg", "jpeg", "png"],
            key=f"photo_{m_id}",
        )

      detail["content"] = st.text_input(
          f"缺失內容描述 #{m_id}",
          value=detail["content"],
          key=f"content_{m_id}",
          placeholder="例如：乳膠漆有明顯刷痕",
      )
      detail["remark"] = st.text_input(
          f"備註 #{m_id}",
          value=detail["remark"],
          key=f"remark_{m_id}",
          placeholder="例如：需全面砂紙磨平重漆",
      )
      st.markdown("---")

    submitted_form = st.form_submit_button(
        "🚀 確認填報內容並產出 Word 報告", type="primary", use_container_width=True
    )

if "submitted_form" in locals() and submitted_form:
  doc = Document()

  for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

  title_p = doc.add_paragraph()
  title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
  title_run = title_p.add_run(
      f"【{project_info['project_name']}】室內缺失查驗紀錄表"
  )
  title_run.bold = True
  title_run.font.size = Pt(18)

  meta_table = doc.add_table(rows=2, cols=4)
  meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
  meta_data = [
      (
          "工程名稱",
          project_info["project_name"],
          "查驗日期",
          datetime.today().strftime("%Y年%m月%d日"),
      ),
      (
          "施作位置",
          project_info["location"],
          "查驗人員 / 主管",
          f"{project_info['inspector']} / {project_info['supervisor']}",
      ),
  ]
  for r_idx, row_content in enumerate(meta_data):
    for c_idx, text_val in enumerate(row_content):
      meta_table.cell(r_idx, c_idx).text = str(text_val)

  doc.add_paragraph()

  if os.path.exists(plan_filename):
    doc.add_paragraph(
        f"一、 戶別平面圖 ({project_info['location']})"
    ).runs[0].bold = True
    doc.add_picture(plan_filename, width=Inches(5.0))
    doc.add_paragraph()

  doc.add_paragraph("二、 缺失查驗明細表").runs[0].bold = True
  table = doc.add_table(rows=1, cols=5)
  table.alignment = WD_TABLE_ALIGNMENT.CENTER

  headers = ["編號", "缺失位置", "缺失內容", "缺失廠商", "備註"]
  hdr_cells = table.rows[0].cells
  for i, header_text in enumerate(headers):
    hdr_cells[i].text = header_text
    hdr_cells[i].paragraphs[0].runs[0].bold = True

  for m in st.session_state.markers:
    m_id = m["id"]
    detail = st.session_state.defect_details.get(m_id, {})
    row_cells = table.add_row().cells
    row_cells[0].text = f"#{m_id}"
    row_cells[1].text = detail.get("space", "")
    row_cells[2].text = detail.get("content", "")
    row_cells[3].text = detail.get("vendor", "")
    row_cells[4].text = detail.get("remark", "")

  doc.add_paragraph()
  doc.add_paragraph("三、 現場缺失照片紀錄").runs[0].bold = True

  for m in st.session_state.markers:
    m_id = m["id"]
    detail = st.session_state.defect_details.get(m_id, {})
    photo_file = detail.get("photo")
    if photo_file is not None:
      doc.add_paragraph(
          f"編號 #{m_id} 缺失照片 ({detail.get('space', '')} -"
          f" {detail.get('vendor', '')}):"
      )
      doc.add_picture(io.BytesIO(photo_file.read()), width=Inches(3.0))
      doc.add_paragraph()

  output_io = io.BytesIO()
  doc.save(output_io)
  output_io.seek(0)

  st.success("🎉 Word 報告已成功產出！請點擊下方按鈕下載：")
  st.download_button(
      label="📥 下載室內缺失查驗紀錄表 (.docx)",
      data=output_io,
      file_name=f"室內缺失查驗紀錄_{project_info['location'].replace(' ', '_')}_{datetime.today().strftime('%Y%m%d')}.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
  )