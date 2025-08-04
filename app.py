# app.py  —— 商标请款单网页版（Streamlit Cloud 即用）
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

# ---------------- 工具函数（已适配网页） ----------------
def extract_pdf_data(pdf_file):
    """
    从 PDF 提取申请人、信用代码、日期、商标+类别
    """
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text += page.extract_text().replace("　", " ").replace("\xa0", " ")

    applicant = "N/A"
    unified_credit_code = "N/A"
    final_date = "N/A"
    trademarks_with_categories = []

    # 申请人
    m = re.search(r"申请人名称\(中文\)：\s*(.*?)\s*\(", text)
    if not m:
        m = re.search(r"答辩人名称\(中文\)：\s*(.*?)\s*\(", text)
    applicant = m.group(1).strip() if m else "N/A"

    # 统一社会信用代码
    m = re.search(r"统一社会信用代码：\s*([0-9A-Z]+)", text)
    unified_credit_code = m.group(1).strip() if m else "N/A"

    # 日期
    m = re.search(r"(\d{4}年\s*\d{1,2}月\s*\d{1,2}日)", text)
    final_date = m.group(1).replace(" ", "") if m else "N/A"

    # 商标+类别
    tm_blocks = re.findall(r"商标代理委托书.*?代理\s*(.*?)\s*商标\s*的\s*如下", text, re.DOTALL)
    for tm in tm_blocks:
        tm = tm.strip()
        if not tm:
            continue
        cats = re.findall(r"类别：(\d+)", text)
        if not cats:
            cats = ["待填写"]
        for c in cats:
            trademarks_with_categories.append({"商标名称": tm, "类别": c})

    return {
        "申请人": applicant,
        "统一社会信用代码": unified_credit_code,
        "日期": final_date,
        "商标列表": trademarks_with_categories,
        "事宜类型": "商标注册申请"
    }

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

def create_word_doc(data, agent_fee, categories):
    BASE = os.path.dirname(__file__)
    doc = Document(os.path.join(BASE, "请款单.docx"))
    num_items = len(categories)
    total_official = num_items * 270
    total_agent = num_items * agent_fee
    total_subtotal = total_official + total_agent
    total_upper = number_to_upper(total_subtotal)

    # 段落占位符
    for p in doc.paragraphs:
        p.text = p.text.replace("{申请人}", data["申请人"]) \
                       .replace("{事宜类型}", "商标注册申请") \
                       .replace("{日期}", data["日期"]) \
                       .replace("{总官费}", str(total_official)) \
                       .replace("{总代理费}", str(total_agent)) \
                       .replace("{总计}", str(total_subtotal)) \
                       .replace("{大写}", total_upper)

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

    filename = f"请款单（{data['申请人']}-{total_subtotal}-{data['日期']}）.docx"
    path = os.path.join("output", filename)
    os.makedirs("output", exist_ok=True)
    doc.save(path)
    return path

# ---------------- Streamlit 界面 ----------------
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

    agent_fee = st.number_input("代理费（元/项）", 0, 5000, 800, key=f"fee_{idx}")
    categories = []
    for tm in data["商标列表"]:
        cats = st.text_input(f"{tm['商标名称']} 类别（逗号分隔）", tm["类别"], key=f"cat_{idx}_{tm['商标名称']}")
        for c in cats.split(","):
            if c.strip():
                categories.append({"商标名称": tm["商标名称"], "类别": c.strip()})

    if st.button("生成该客户 Word", key=f"go_{idx}"):
        fname = create_word_doc(data, agent_fee, categories)
        generated_docs.append(fname)
        all_applicants.append({
            "申请人": applicant,
            "统一社会信用代码": data["统一社会信用代码"],
            "日期": data["日期"],
            "总官费": len(categories) * 270,
            "总代理费": len(categories) * agent_fee,
            "总计": len(categories) * (270 + agent_fee)
        })
        st.success("已生成 ✅")

# 在你原来按钮的位置
if st.button("生成该客户Word", key=f"gen_{idx}"):
    fname = create_word_doc(data, agent_fee, categories)  # 先生成
    # 立即提供下载
    with open(fname, "rb") as f:
        st.download_button(
            label="📥 下载请款单",
            data=f.read(),
            file_name=os.path.basename(fname),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
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
    st.download_button("⬇️ 下载全部文件", data=zip_buffer, file_name="商标请款单+发票申请.zip", mime="application/zip")


