#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AI求职助手 - Streamlit Web UI"""

import os
import re
import shutil
import smtplib
import tempfile
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import openpyxl
import streamlit as st

# ── 页面配置 ──────────────────────────────────────────────
st.set_page_config(page_title="AI求职助手", page_icon="📨", layout="centered")

st.title("📨 AI求职助手")
st.caption("上传简历和岗位Excel，自动生成定制邮件并发送")

# ── 侧边栏：用户配置 ──────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 个人配置")
    name = st.text_input("姓名", value="张衡旭")
    school = st.text_input("学校", value="浙江大学")
    major = st.text_input("专业", value="金融学")
    grade = st.text_input("年级", value="大三")
    intern_period = st.text_input("可实习时间", value="2025年7月至10月")
    grad_year = st.text_input("毕业时间", value="2026年6月")

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


def parse_jd(jd_text):
    first_line = jd_text.split("\n")[0].strip()
    company = re.split(r"[-–—]", first_line)[0].strip()
    company = re.sub(r"[【】\[\]「」]", "", company).strip()

    pos_patterns = [
        r"招聘(.{2,20}?)(?:实习生|助理|分析师)",
        r"[-–—]\s*(.{2,20}?)(?:实习生|助理|分析师)",
        r"【(.{2,20}?)(?:实习生|助理|分析师)",
    ]
    position = "投资实习生"
    for p in pos_patterns:
        m = re.search(p, first_line)
        if m:
            position = m.group(1).strip() + "实习生"
            break

    subject_fmt = None
    for pattern in [r"邮件(?:主题|标题)[：:]\s*(.+)", r"邮件及简历命名格式[：:]\s*(.+)"]:
        m = re.search(pattern, jd_text)
        if m:
            raw = m.group(1).strip().split("\n")[0]
            raw = re.split(r"简历(?:标题|命名)[：:]", raw)[0].strip()
            subject_fmt = raw
            break

    resume_fmt = None
    for pattern in [r"简历(?:命名|标题)[：:]\s*(.+)", r"邮件&简历命名[】\]]\s*(.+)"]:
        m = re.search(pattern, jd_text)
        if m:
            s = m.group(1).strip().split("\n")[0]
            resume_fmt = re.sub(r"[（(].*", "", s).strip()
            break

    return company, position, subject_fmt, resume_fmt


def detect_focus(jd_text):
    jd_lower = jd_text.lower()
    if any(k in jd_lower for k in ["blockchain", "defi", "crypto", "web3", "链上"]):
        return "web3"
    if any(k in jd_lower for k in ["business development", "bd", "合作"]):
        return "bd"
    if any(k in jd_lower for k in ["quant", "量化", "python", "数据分析"]):
        return "quant"
    return "investment"


HIGHLIGHTS = {
    "web3": "在0xU社区深度参与DeFi/链上交易，并担任WhaleRyder BD，具备Web3实战经验",
    "investment": "在水木清华基金完成投后实习，跟踪150+项目，撰写200+页运营报告",
    "quant": "具备Python/R面板数据分析经验，熟练使用Wind/CSMAR，完成绿色信贷量化研究",
    "bd": "担任WhaleRyder BD及Finternet峰会志愿者组长，具备多语言商务沟通能力",
}


def make_subject(subject_fmt, company, position, name, school, major, grad_year, intern_period):
    if not subject_fmt:
        return f"求职申请-{name}-{school}-{position}"
    s = subject_fmt.split("\n")[0]
    for k, v in {
        "最早可入职时间": intern_period.split("至")[0] if "至" in intern_period else intern_period,
        "可入职时间": intern_period.split("至")[0] if "至" in intern_period else intern_period,
        "最早可": intern_period.split("至")[0] if "至" in intern_period else intern_period,
        "姓名": name, "名字": name,
        "学校": school, "本科学校": school,
        "专业": major, "岗位": position, "职位": position,
        "毕业时间": grad_year, "毕业年份": grad_year[:4],
        "一周几天": "5天", "硕士学校（若有）": "",
    }.items():
        s = s.replace(k, v)
    return re.sub(r"[-–—]{2,}", "-", s).strip()


def make_resume_name(resume_fmt, company, position, name, school, major, grad_year, intern_period):
    if not resume_fmt:
        return f"{name}_简历_{position}_{company}.pdf"
    s = resume_fmt.split("\n")[0]
    s = re.sub(r"[（(].*", "", s).strip()
    start = intern_period.split("至")[0] if "至" in intern_period else intern_period
    for k, v in {
        "姓名": name, "名字": name,
        "学校": school, "本科学校": school,
        "本科/研究生学校及专业": f"{school}{major}",
        "专业": major, "岗位": position, "职位": position, "投递岗位": position,
        "毕业时间": grad_year, "到岗时间": start,
        "周实习X天": "周实习5天", "实习X个月": "实习4个月",
        "年级": grade, "FA分析师助理": position, "人才战略": position,
    }.items():
        s = s.replace(k, v)
    s = re.sub(r"[/\\:*?\"<>|]", "-", s)
    s = re.sub(r"\s+", "", s)
    if not s.endswith(".pdf"):
        s += ".pdf"
    return s


def make_body(company, position, focus, name, school, major, grade, intern_period):
    highlight = HIGHLIGHTS.get(focus, HIGHLIGHTS["investment"])
    return (
        f"您好！\n\n"
        f"我是{school}{major}{grade}学生{name}，对贵司{position}岗位非常感兴趣。"
        f"{highlight}，与岗位要求高度匹配。"
        f"可于{intern_period}全职实习，期待进一步交流。\n\n"
        f"{name}"
    )


# ── 生成预览 ──────────────────────────────────────────────
jobs = load_jobs(excel_file.read())
st.success(f"读取到 {len(jobs)} 个岗位")

max_rows = st.slider("处理岗位数量", 1, len(jobs), min(5, len(jobs)))
jobs = jobs[:max_rows]

# 生成所有岗位数据
previews = []
for job in jobs:
    company, position, subject_fmt, resume_fmt = parse_jd(job["jd"])
    focus = detect_focus(job["jd"])
    previews.append({
        "company": company,
        "position": position,
        "email": job["email"],
        "subject": make_subject(subject_fmt, company, position, name, school, major, grad_year, intern_period),
        "resume_name": make_resume_name(resume_fmt, company, position, name, school, major, grad_year, intern_period),
        "body": make_body(company, position, focus, name, school, major, grade, intern_period),
        "focus": focus,
    })

# ── 预览表格 ──────────────────────────────────────────────
st.divider()
st.subheader("📋 预览")

for i, p in enumerate(previews):
    with st.expander(f"{i+1}. {p['company']} — {p['position']}"):
        st.write(f"**收件人：** `{p['email']}`")
        st.write(f"**主题：** {p['subject']}")
        st.write(f"**附件：** {p['resume_name']}")
        st.text_area("邮件正文", p["body"], height=160, key=f"body_{i}", disabled=True)

# ── 发送区域 ──────────────────────────────────────────────
st.divider()
st.subheader("🚀 发送")

if not sender_email or not auth_code:
    st.warning("请在左侧填写发件邮箱和授权码")
    st.stop()

send_to_self = st.checkbox("📬 先发给自己测试（不发给HR）", value=True)

if st.button("开始发送", type="primary", use_container_width=True):
    resume_bytes = resume_file.read()

    progress = st.progress(0)
    status_box = st.empty()
    results = []

    for i, p in enumerate(previews):
        status_box.info(f"正在发送 {i+1}/{len(previews)}：{p['company']}...")

        # 写临时PDF
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

            results.append({"company": p["company"], "status": "✅ 已发送"})
            time.sleep(1)
        except Exception as e:
            results.append({"company": p["company"], "status": f"❌ {e}"})
        finally:
            os.unlink(tmp_pdf)

        progress.progress((i + 1) / len(previews))

    status_box.empty()
    st.success("发送完成！")
    for r in results:
        st.write(f"{r['status']} — {r['company']}")
