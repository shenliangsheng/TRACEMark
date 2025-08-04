# app.py  —— 商标请款单网页版（可识别商标+类别）
import os
import re
import io
import zipfile
import datetime
import streamlit as st
import pdfplumber
from docx import Document
from openpyxl import load_workbook
from collections import defaultdict

# ---------- PDF 解析 ----------
def extract_pdf_data(pdf_file):
    data = {
        "申请人": "N/A",
        "统一社会信用代码": "N/A",
        "日期": "N/A",
        "商标列表": []
    }

    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join(
            page.extract_text().replace("　", " ").replace("\xa0", " ").strip()
            for page in pdf.pages
        )

    # 申请人
    m = re.search(r"申请人名称\(中文\)：\s*([^（(]+)", full_text)
    if not m:
        m = re.search(r"答辩人名称\(中文\)：\s*([^（(]+)", full_text)
    if m:
        data["申请人"] = m.group(1).strip()

    # 统一社会信用代码
    m = re.search(r"统一社会信用代码[:：]\s*([0-9A-Z]+)", full_text, re.I)
    if m:
        data["统一社会信用代码"] = m.group(1).strip()

    # 日期
    m = re.search(r"(\d{4}年\s*\d{1,2}月\s*\d{1,2}日)", full_text)
    if m:
        data["日期"] = m.group(1).replace(" ", "")

    # 商标+类别
    # 每遇到“商标代理委托书”就提取商标名，并把紧挨着的类别一起拿
    tm_blocks = re.finditer(r"商标代理委托书[\s\S]*?代理\s*([^\n]+?)\s*商标[\s\S]*?事宜", full_text, re.I)
    for block in tm_blocks:
        tm_name = block.group(1).strip()
        # 在商标名附近 500 字符内找类别
        near = full_text[block.start():block.end()+500]
        cats = re.findall(r"类别[:：]?\s*(\d{1,3})", near)
        if not cats:
            cats = ["待填写"]
        for c in cats:
            data["商标列表"].append({"商标名称": tm_name, "类别": c})

    return data

# ---------- 金额大写 ----------
def number_to_upper(amount):
    CN_NUM = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']
    CN_UNIT = ['元', '拾', '佰', '仟', '万', '拾万', '佰万', '仟万', '亿']
    s = str(int(amount))[::-1]
    result = []
    for i, ch in enumerate(s):
        digit = int(ch)
        unit = CN_UNIT[i] if i < len(CN_UNIT) else ''
        if digit != 0:
            result.append(f"{CN_NUM[digit]}{unit}")
    return ''.join(reversed(result)) + "元整"

# ---------- 生成 Word ----------
def create_word_doc(data, agent_fee, categories):
    base = os.path.dirname(__file__)
    doc = Document(os.path.join(base, "请款单.docx"))
    num_items = len(categories)
    total_official = num_items * 270
    total_agent = num_items * agent_fee
    total_subtotal = total_official + total_agent
    total_upper = number_to_upper(total_subtotal)

    # 替换段落
    for p in doc.paragraphs:
        p.text = (
            p.text.replace("{申请人}", data["申请人"])
                  .replace("{事宜类型}", "商标注册申请")
                  .replace("{日期}", data["日期"])
                  .replace("{总官费}", str(total_official))
                  .replace("{总代理费}", str(total_agent))
                  .replace("{总计}", str(total_subtotal))
                  .replace("{大写}", total_upper)
        )

    # 表格
    table = doc.tables[0]
    if len(table.rows) > 2:
        tbl = table.rows[1]._element
        tbl.getparent().remove(tbl)
    for idx, item in enumerate(categories, 1):
        row = table.add_row().cells
        row[0].text = str(idx)
        row[1].text = item["商标名称"]
        row[2].text = "商标注册申请"
        row[3].text = item["类别"]
        row[4].text = "¥270"
        row[5].text = f"¥{agent_fee}"
        row[6].text = f"¥{270 + agent_fee}"

    # 合计行
    total_row = table.add_row().cells
    total_row[0].merge(total_row[3])
    total_row[0].text = "合计"
    total_row[4].text = f"¥{total_official}"
    total_row[5].text = f"¥{total_agent}"
    total_row[6].text = f"¥{total_subtotal}"

    filename = f"请款单（{data['申请人']}-{total_subtotal}-{data['日期']}）.docx"
    path = os.path.join("output", filename)
    os.makedirs("output", exist_ok=True)
    doc.save(path)
    return path

# ---------- Streamlit UI ----------
st.set_page_config(page_title="商标请款单生成器")
st.title("📄 商标请款单生成器")
st.info("上传 PDF → 填写代理费 → 一键下载 Word + Excel")

files = st.file_uploader("选择 PDF（可多选）", type="pdf", accept_multiple_files=True)
if not files:
    st.stop()

all_applicants = []
generated_docs = []

for idx, pdf_file in enumerate(files):
    data = extract_pdf_data(pdf_file)
    applicant = data["申请人"]
    st.markdown(f"---\n### 👤 {applicant}")

    # 如果 PDF 没解析出商标/类别，提示用户手动补
    if not data["商标列表"]:
        st.warning("未识别到商标信息，请手动补充")
        tm = st.text_input("商标名称", key=f"tm_{idx}")
        cat = st.text_input("类别", key=f"cat_{idx}")
        data["商标列表"] = [{"商标名称": tm, "类别": cat}] if tm else []

    agent_fee = st.number_input("代理费（元/项）", 0, 5000, 800, key=f"fee_{idx}")
    categories = []
    for tm in data["商标列表"]:
        cat_input = st.text_input(f"{tm['商标名称']} 类别", tm["类别"], key=f"cat_{idx}_{tm['商标名称']}")
        for c in cat_input.split(","):
            if c.strip():
                categories.append({"商标名称": tm["商标名称"], "类别": c.strip()})

    if st.button("生成并下载 Word", key=f"go_{idx}"):
        fname = create_word_doc(data, agent_fee, categories)
        with open(fname, "rb") as f:
            st.download_button(
                label="📥 下载请款单",
                data=f.read(),
                file_name=os.path.basename(fname),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        generated_docs.append(fname)
        all_applicants.append({
            "申请人": applicant,
            "统一社会信用代码": data["统一社会信用代码"],
            "日期": data["日期"],
            "总官费": len(categories) * 270,
            "总代理费": len(categories) * agent_fee,
            "总计": len(categories) * (270 + agent_fee)
        })

# 打包下载 Excel
if all_applicants:
    if st.button("📦 打包下载全部文件"):
        wb = load_workbook(os.path.join(os.path.dirname(__file__), "发票申请表.xlsx"))
        ws = wb.active
        row = 2
        for a in all_applicants:
            ws[f'B{row}'] = a["申请人"]
            ws[f'C{row}'] = a["统一社会信用代码"]
            ws[f'G{row}'] = a["总官费"]
            ws[f'H{row}'] = a["总官费"]
            ws[f'I{row}'] = a["总计"]
            ws[f'Q{row}'] = a["日期"]
            row += 1
            ws[f'B{row}'] = a["申请人"]
            ws[f'C{row}'] = a["统一社会信用代码"]
            ws[f'G{row}'] = a["总代理费"]
            ws[f'H{row}'] = a["总代理费"]
            ws[f'I{row}'] = a["总计"]
            ws[f'Q{row}'] = a["日期"]
            row += 1
        excel_name = f"发票申请表-{datetime.date.today()}.xlsx"
        wb.save(os.path.join("output", excel_name))

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for f in generated_docs + [excel_name]:
                zf.write(os.path.join("output", f), f)
        zip_buffer.seek(0)
        st.download_button(
            label="⬇️ 下载全部文件",
            data=zip_buffer,
            file_name="商标请款单+发票申请.zip",
            mime="application/zip"
        )
