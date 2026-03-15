#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AI求职助手 - Streamlit Web UI"""

import os
import re
import smtplib
import tempfile
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import openai
import openpyxl
import PyPDF2
import streamlit as st

# ── 页面配置 ──────────────────────────────────────────────
st.set_page_config(page_title="AI求职助手", page_icon="📨", layout="wide")

# OpenAI兼容API配置
API_KEY = os.getenv("API_KEY", "")
API_BASE = os.getenv("API_BASE", "https://api.siliconflow.cn/v1")  # 默认用SiliconFlow
MODEL_NAME = os.getenv("MODEL_NAME", "Qwen/Qwen2.5-7B-Instruct")

if not API_KEY:
    st.error("未配置API Key，请联系管理员")
    st.stop()

client = openai.OpenAI(api_key=API_KEY, base_url=API_BASE)

st.title("📨 AI求职助手")
st.caption("上传简历和岗位Excel，自动生成定制邮件并发送")

# ── 侧边栏：用户配置 ──────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 个人配置")
    name = st.text_input("姓名", placeholder="张三")
    school = st.text_input("学校", placeholder="清华大学")
    major = st.text_input("专业", placeholder="计算机科学")
    grade = st.text_input("年级", placeholder="大三")
    intern_period = st.text_input("可实习时间", placeholder="2025年7月至10月")
    grad_year = st.text_input("毕业时间", placeholder="2026年6月")

    st.divider()
    st.header("📧 邮箱配置")
    sender_email = st.text_input("发件邮箱", placeholder="xxxxxx@qq.com")
    auth_code = st.text_input("授权码", type="password", placeholder="QQ邮箱授权码")

    smtp_options = {
        "QQ邮箱 (@qq.com)": ("smtp.qq.com", 465),
        "163邮箱 (@163.com)": ("smtp.163.com", 465),
        "Gmail (@gmail.com)": ("smtp.gmail.com", 465),
        "Outlook (@outlook.com)": ("smtp.office365.com", 587),
    }
    smtp_choice = st.selectbox("邮箱类型", list(smtp_options.keys()))
    smtp_host, smtp_port = smtp_options[smtp_choice]

# ── 主区域：文件上传 ──────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    resume_file = st.file_uploader("📄 上传简历 PDF", type=["pdf"])
with col2:
    excel_file = st.file_uploader("📊 上传岗位 Excel", type=["xlsx", "xls"])

if not resume_file or not excel_file:
    st.info("请上传简历PDF和岗位Excel后继续")
    st.stop()

# ── 解析Excel ─────────────────────────────────────────────
@st.cache_data
def load_jobs(excel_bytes):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(excel_bytes)
        tmp_path = f.name
    wb = openpyxl.load_workbook(tmp_path)
    ws = wb.active
    jobs = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        if len(row) < 4:
            continue
        jd_text, email = row[2], row[3]
        if not jd_text or not email:
            continue
        jobs.append({"jd": str(jd_text).strip(), "email": str(email).strip()})
    os.unlink(tmp_path)
    return jobs


# ── 提取简历亮点 ──────────────────────────────────────────
@st.cache_data
def extract_resume_highlights(pdf_bytes):
    """用AI从PDF提取3-5个核心亮点"""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        f.write(pdf_bytes)
        tmp_path = f.name
    
    try:
        reader = PyPDF2.PdfReader(tmp_path)
        text = "\n".join([page.extract_text() for page in reader.pages])
    finally:
        os.unlink(tmp_path)
    
    prompt = f"""从以下简历中提取3-5个最核心的亮点，每个亮点一句话，用于求职邮件。
要求：
1. 突出量化成果（数字、项目数量、金额等）
2. 突出相关经验（实习、项目、技能）
3. 每个亮点15-25字

简历内容：
{text[:3000]}

请直接输出亮点列表，每行一个，格式：
- 亮点1
- 亮点2
- 亮点3"""

    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[{"role": "user", "content": prompt}],
        max_tokens=500,
        temperature=0.7
    )
    return response.choices[0].message.content.strip()


# ── AI解析JD ──────────────────────────────────────────────
def ai_parse_jd(jd_text, name, school, major, grad_year, intern_period):
    """用AI解析JD，提取公司、职位、邮件主题格式、简历命名格式"""
    prompt = f"""解析以下招聘JD，提取关键信息。

JD内容：
{jd_text[:2000]}

请提取：
1. 公司名称（去掉【】等符号）
2. 职位名称
3. 邮件主题格式（如果JD中有明确要求，比如"邮件主题：xxx"，则提取；否则返回null）
4. 简历文件命名格式（如果JD中有明确要求，比如"简历命名：xxx"，则提取；否则返回null）

用户信息：
- 姓名：{name}
- 学校：{school}
- 专业：{major}
- 毕业时间：{grad_year}
- 可实习时间：{intern_period}

返回JSON格式：
{{
  "company": "公司名",
  "position": "职位名",
  "subject_format": "邮件主题格式（如有要求）或null",
  "resume_format": "简历命名格式（如有要求）或null"
}}"""

    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[{"role": "user", "content": prompt}],
        max_tokens=300,
        temperature=0.7
    )
    
    import json
    # 提取JSON（去掉markdown代码块标记）
    text = response.choices[0].message.content.strip()
    if "```json" in text:
        text = text.split("```json")[1].split("```")[0].strip()
    elif "```" in text:
        text = text.split("```")[1].split("```")[0].strip()
    
    result = json.loads(text)
    
    # 如果没有格式要求，生成默认格式
    if not result.get("subject_format") or result["subject_format"] == "null":
        result["subject_format"] = f"求职申请-{name}-{school}-{result['position']}"
    else:
        # 替换占位符
        s = result["subject_format"]
        for k, v in {
            "姓名": name, "学校": school, "专业": major,
            "毕业时间": grad_year, "可入职时间": intern_period.split("至")[0] if "至" in intern_period else intern_period,
        }.items():
            s = s.replace(k, v)
        result["subject_format"] = s
    
    if not result.get("resume_format") or result["resume_format"] == "null":
        result["resume_format"] = f"{name}_简历_{result['position']}_{result['company']}.pdf"
    else:
        s = result["resume_format"]
        for k, v in {
            "姓名": name, "学校": school, "专业": major,
            "毕业时间": grad_year, "到岗时间": intern_period.split("至")[0] if "至" in intern_period else intern_period,
        }.items():
            s = s.replace(k, v)
        if not s.endswith(".pdf"):
            s += ".pdf"
        result["resume_format"] = s
    
    return result


# ── AI生成邮件正文 ────────────────────────────────────────
def ai_generate_body(company, position, jd_text, resume_highlights, name, school, major, grade, intern_period):
    """用AI生成个性化邮件正文"""
    prompt = f"""为求职邮件生成正文，要求：
1. 100字以内
2. 中文
3. 结构：问候 + 自我介绍 + 1-2个核心匹配点 + 可实习时间 + 期待交流
4. 核心匹配点要结合JD要求和简历亮点

公司：{company}
职位：{position}
JD摘要：{jd_text[:500]}

简历亮点：
{resume_highlights}

个人信息：
- 姓名：{name}
- 学校：{school}
- 专业：{major}
- 年级：{grade}
- 可实习时间：{intern_period}

直接输出邮件正文，不要任何额外说明。"""

    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[{"role": "user", "content": prompt}],
        max_tokens=300,
        temperature=0.7
    )
    return response.choices[0].message.content.strip()


# ── 生成预览 ──────────────────────────────────────────────
if not all([name, school, major, grade, intern_period, grad_year]):
    st.warning("请在左侧填写完整个人信息")
    st.stop()

jobs = load_jobs(excel_file.read())
st.success(f"读取到 {len(jobs)} 个岗位")

max_rows = st.slider("处理岗位数量", 1, len(jobs), min(5, len(jobs)))
jobs = jobs[:max_rows]

# 提取简历亮点
with st.spinner("AI正在分析简历..."):
    resume_highlights = extract_resume_highlights(resume_file.read())

st.info(f"简历亮点：\n{resume_highlights}")

# 生成所有岗位数据
if "previews" not in st.session_state:
    with st.spinner("AI正在解析JD并生成邮件..."):
        previews = []
        for job in jobs:
            parsed = ai_parse_jd(job["jd"], name, school, major, grad_year, intern_period)
            body = ai_generate_body(
                parsed["company"], parsed["position"], job["jd"],
                resume_highlights, name, school, major, grade, intern_period
            )
            previews.append({
                "company": parsed["company"],
                "position": parsed["position"],
                "email": job["email"],
                "subject": parsed["subject_format"],
                "resume_name": parsed["resume_format"],
                "body": body,
                "sent": False,
            })
        st.session_state.previews = previews
else:
    previews = st.session_state.previews

# ── 预览表格（可编辑）──────────────────────────────────────
st.divider()
st.subheader("📋 预览与编辑")

for i, p in enumerate(previews):
    with st.expander(f"{'✅' if p['sent'] else '📧'} {i+1}. {p['company']} — {p['position']}", expanded=not p['sent']):
        col1, col2 = st.columns([3, 1])
        with col1:
            st.write(f"**收件人：** `{p['email']}`")
        with col2:
            if not p['sent']:
                if st.button("发送此岗位", key=f"send_{i}", type="primary"):
                    if not sender_email or not auth_code:
                        st.error("请先填写邮箱配置")
                    else:
                        try:
                            resume_bytes = resume_file.read()
                            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
                                f.write(resume_bytes)
                                tmp_pdf = f.name

                            msg = MIMEMultipart()
                            msg["From"] = sender_email
                            msg["To"] = p["email"]
                            msg["Subject"] = p["subject"]
                            msg.attach(MIMEText(p["body"], "plain", "utf-8"))

                            with open(tmp_pdf, "rb") as f:
                                part = MIMEBase("application", "octet-stream")
                                part.set_payload(f.read())
                            encoders.encode_base64(part)
                            part.add_header("Content-Disposition", "attachment",
                                            filename=("utf-8", "", p["resume_name"]))
                            msg.attach(part)

                            with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
                                server.login(sender_email, auth_code)
                                server.sendmail(sender_email, p["email"], msg.as_string())

                            os.unlink(tmp_pdf)
                            st.session_state.previews[i]["sent"] = True
                            st.success(f"✅ 已发送到 {p['email']}")
                            st.rerun()
                        except Exception as e:
                            st.error(f"发送失败：{e}")
            else:
                st.success("已发送")
        
        # 可编辑字段
        p["subject"] = st.text_input("邮件主题", p["subject"], key=f"subj_{i}")
        p["resume_name"] = st.text_input("附件名称", p["resume_name"], key=f"resume_{i}")
        p["body"] = st.text_area("邮件正文", p["body"], height=200, key=f"body_{i}")

# ── 批量发送区域 ──────────────────────────────────────────
st.divider()
st.subheader("🚀 批量发送")

if not sender_email or not auth_code:
    st.warning("请在左侧填写发件邮箱和授权码")
    st.stop()

unsent = [p for p in previews if not p["sent"]]
if not unsent:
    st.info("所有岗位已发送完毕")
    st.stop()

send_to_self = st.checkbox("📬 先发给自己测试（不发给HR）", value=False)

if st.button(f"批量发送剩余 {len(unsent)} 个岗位", type="primary", use_container_width=True):
    resume_bytes = resume_file.read()

    progress = st.progress(0)
    status_box = st.empty()

    for idx, i in enumerate([previews.index(p) for p in unsent]):
        p = previews[i]
        status_box.info(f"正在发送 {idx+1}/{len(unsent)}：{p['company']}...")

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
            f.write(resume_bytes)
            tmp_pdf = f.name

        to_addr = sender_email if send_to_self else p["email"]

        try:
            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = to_addr
            msg["Subject"] = p["subject"]
            msg.attach(MIMEText(p["body"], "plain", "utf-8"))

            with open(tmp_pdf, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", "attachment",
                            filename=("utf-8", "", p["resume_name"]))
            msg.attach(part)

            with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
                server.login(sender_email, auth_code)
                server.sendmail(sender_email, to_addr, msg.as_string())

            st.session_state.previews[i]["sent"] = True
            time.sleep(1.5)
        except Exception as e:
            status_box.error(f"❌ {p['company']} 发送失败：{e}")
            time.sleep(2)
        finally:
            os.unlink(tmp_pdf)

        progress.progress((idx + 1) / len(unsent))

    status_box.empty()
    st.success("批量发送完成！")
    st.rerun()
