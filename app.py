import io
import os
from datetime import datetime
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from PIL import Image, ImageDraw
import streamlit as st
from streamlit_image_coordinates import streamlit_image_coordinates

st.set_page_config(
    page_title="建築工地互動式缺失查驗系統", page_icon="🏗️", layout="wide"
)

st.title("🏗️ 建築工地互動式平面圖缺失查驗系統")
st.markdown(
    "操作流程：**1. 選擇戶別與樓層** ➔ **2. 直接點擊平面圖任意位置新增缺失編號** ➔"
    " **3. 於下方表格填寫細節並上傳照片** ➔ **4. 一鍵匯出 Word 報告**"
)

# 初始化全域總資料庫：以「戶別_樓層」作為 Key 來獨立儲存各個區塊的 markers 與 defect_details
if "project_data" not in st.session_state:
  st.session_state.project_data = {}

st.sidebar.header("📌 1. 查驗位置設定")
unit_choice = st.sidebar.selectbox(
    "選擇戶別", ["A1戶別", "A2戶別", "A3戶別", "B1戶別", "B2戶別", "B3戶別"]
)

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

# 組合成當前唯一的空間識別 Key (例如：A1戶別_3F)
current_key = f"{unit_choice}_{floor_choice}"

# 如果這個空間還沒有在資料庫建立過，初始化它的結構
if current_key not in st.session_state.project_data:
  st.session_state.project_data[current_key] = {
      "markers": [],
      "defect_details": {},
  }

project_info = {
    "project_name": st.sidebar.text_input(
        "工程名稱", value="弘峻草福段A區新建工程"
    ),
    "location": f"{unit_choice} ({floor_choice})",
    "inspector": st.sidebar.text_input("查驗人員", value=""),
    "supervisor": st.sidebar.text_input("施工負責人", value=""),
}

st.sidebar.divider()
if st.sidebar.button("🗑️ 清除目前【" + current_key + "】的所有標註與表格"):
  st.session_state.project_data[current_key] = {
      "markers": [],
      "defect_details": {},
  }
  st.rerun()

st.subheader(f"🗺️ 2. 平面圖互動標註區 ({project_info['location']})")
st.markdown(
    "💡 **提示：直接用滑鼠點擊下方平面圖的任意位置，即可自動在該處長出紅點與編號！**"
)

plan_filename = f"{unit_choice.replace('戶別', '')}_plan.jpg"

if os.path.exists(plan_filename):
  original_img = Image.open(plan_filename).convert("RGB")
  orig_w, orig_h = original_img.size

  # 強制將圖片寬高縮小為 0.5 倍（確保畫面瀏覽順暢）
  new_w, new_h = int(orig_w * 0.5), int(orig_h * 0.5)
  base_img = original_img.resize((new_w, new_h), Image.Resampling.LANCZOS)

  draw_img = base_img.copy()
  draw = ImageDraw.Draw(draw_img)

  # 取得當前戶別與樓層專屬的 markers
  current_markers = st.session_state.project_data[current_key]["markers"]
  current_details = st.session_state.project_data[current_key]["defect_details"]

  # 在圖上繪製已存在的標註紅點與編號
  for m in current_markers:
    mx, my = m["x"], m["y"]
    r = 12
    draw.ellipse(
        [mx - r, my - r, mx + r, my + r], fill="red", outline="white", width=2
    )
    draw.text((mx - 4, my - 7), str(m["id"]), fill="white")

  # 顯示互動地圖並接收點擊座標
  coord = streamlit_image_coordinates(
      draw_img, key=f"map_{current_key}"
  )

  if coord is not None:
    click_x, click_y = coord["x"], coord["y"]
    last_marker = current_markers[-1] if current_markers else None
    if (
        last_marker is None
        or last_marker["x"] != click_x
        or last_marker["y"] != click_y
    ):
      next_id = max([m["id"] for m in current_markers], default=0) + 1
      current_markers.append({"id": next_id, "x": click_x, "y": click_y})
      current_details[next_id] = {
          "space": "牆面",
          "space_custom": "",
          "content": "",
          "vendor": "油漆",
          "vendor_custom": "",
          "photo": None,
          "remark": "",
      }
      st.rerun()

  # 控制面板區塊（置於平面圖下方）
  st.markdown("---")
  c_ctrl1, c_ctrl2 = st.columns([2, 2])
  with c_ctrl1:
    st.markdown(
        f"**📍 目前【{current_key}】已建立缺失數：** `{len(current_markers)}` 處"
    )
  with c_ctrl2:
    if current_markers:
      del_target = st.selectbox(
          "選擇要刪除的編號", [m["id"] for m in current_markers], key=f"del_{current_key}"
      )
      if st.button("🗑️ 刪除選定編號", type="secondary"):
        st.session_state.project_data[current_key]["markers"] = [
            x for x in current_markers if x["id"] != del_target
        ]
        if del_target in current_details:
          del current_details[del_target]
        st.rerun()

else:
  st.warning(
      f"⚠️ 尚未偵測到 `{unit_choice}` 的平面圖檔案 (`{plan_filename}`)。"
      f" 請確認有將對應的 JPG 檔上傳到 GitHub 專案根目錄中。"
  )

st.divider()

st.subheader(f"📋 3. 缺失查驗明細填報表 ({project_info['location']})")

current_markers = st.session_state.project_data[current_key]["markers"]
current_details = st.session_state.project_data[current_key]["defect_details"]

if not current_markers:
  st.info(f"👈 目前【{current_key}】尚無缺失，請直接點擊上方平面圖新增！")
else:
  with st.form(f"defect_table_form_{current_key}"):
    for m in current_markers:
      m_id = m["id"]
      if m_id not in current_details:
        current_details[m_id] = {
            "space": "牆面",
            "space_custom": "",
            "content": "",
            "vendor": "油漆",
            "vendor_custom": "",
            "photo": None,
            "remark": "",
        }

      detail = current_details[m_id]

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
            key=f"space_{current_key}_{m_id}",
        )
        if space_sel == "自訂":
          detail["space_custom"] = st.text_input(
              f"手動輸入位置 #{m_id}",
              value=detail.get("space_custom", ""),
              key=f"space_c_{current_key}_{m_id}",
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
            "防水",
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
            key=f"vendor_{current_key}_{m_id}",
        )
        if vendor_sel == "自訂":
          detail["vendor_custom"] = st.text_input(
              f"手動輸入廠商 #{m_id}",
              value=detail.get("vendor_custom", ""),
              key=f"vendor_c_{current_key}_{m_id}",
          )
          detail["vendor"] = detail["vendor_custom"]
        else:
          detail["vendor"] = vendor_sel

      with c3:
        detail["photo"] = st.file_uploader(
            f"缺失照片 #{m_id} (支援手機拍照)",
            type=["jpg", "jpeg", "png"],
            key=f"photo_{current_key}_{m_id}",
        )

      detail["content"] = st.text_input(
          f"缺失內容描述 #{m_id}",
          value=detail["content"],
          key=f"content_{current_key}_{m_id}",
          placeholder="例如：乳膠漆有明顯刷痕",
      )
      detail["remark"] = st.text_input(
          f"備註 #{m_id}",
          value=detail["remark"],
          key=f"remark_{current_key}_{m_id}",
          placeholder="例如：需全面砂紙磨平重漆",
      )
      st.markdown("---")

    submitted_form = st.form_submit_button(
        f"🚀 確認填報內容並產出【{project_info['location']}】Word 報告",
        type="primary",
        use_container_width=True,
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
        f"一、 戶別平面圖與標註位置 ({project_info['location']})"
    ).runs[0].bold = True
    img_byte_arr = io.BytesIO()
    draw_img.save(img_byte_arr, format="JPEG")
    img_byte_arr.seek(0)
    doc.add_picture(img_byte_arr, width=Inches(5.0))
    doc.add_paragraph()

  doc.add_paragraph("二、 缺失查驗明細表").runs[0].bold = True
  table = doc.add_table(rows=1, cols=6)
  table.alignment = WD_TABLE_ALIGNMENT.CENTER

  headers = ["編號", "缺失位置", "缺失內容", "缺失廠商", "現場照片", "備註"]
  hdr_cells = table.rows[0].cells
  for i, header_text in enumerate(headers):
    hdr_cells[i].text = header_text
    hdr_cells[i].paragraphs[0].runs[0].bold = True

  for m in current_markers:
    m_id = m["id"]
    detail = current_details.get(m_id, {})
    row_cells = table.add_row().cells

    row_cells[0].text = f"#{m_id}"
    row_cells[1].text = detail.get("space", "")
    row_cells[2].text = detail.get("content", "")
    row_cells[3].text = detail.get("vendor", "")

    photo_file = detail.get("photo")
    if photo_file is not None:
      photo_para = row_cells[4].paragraphs[0]
      photo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
      run = photo_para.add_run()
      run.add_picture(io.BytesIO(photo_file.read()), width=Inches(1.2))
      row_cells[4].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    else:
      row_cells[4].text = "無照片"

    row_cells[5].text = detail.get("remark", "")

  output_io = io.BytesIO()
  doc.save(output_io)
  output_io.seek(0)

  st.success("🎉 Word 報告已成功產出！請點擊下方按鈕直接下載檔案：")
  st.download_button(
      label=f"📥 下載 {project_info['location']} 缺失查驗紀錄表 (.docx)",
      data=output_io,
      file_name=f"室內缺失查驗紀錄_{project_info['location'].replace(' ', '_')}_{datetime.today().strftime('%Y%m%d')}.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
  )
