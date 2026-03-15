#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AI求职助手 - 读取Excel岗位列表，生成定制邮件并发送"""

import os
import re
import shutil
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import openpyxl

from config import (
    QQ_EMAIL, QQ_AUTH_CODE, RESUME_PDF, EXCEL_FILE,
    NAME, SCHOOL, MAJOR, INTERN_PERIOD,
    SMTP_HOST, SMTP_PORT
)

# ── 输出目录 ──────────────────────────────────────────────
OUTPUT_DIR = "outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ── 1. 解析Excel ──────────────────────────────────────────
def load_jobs(excel_file, max_rows=None):
    """读取Excel，返回岗位列表"""
    wb = openpyxl.load_workbook(excel_file)
    ws = wb.active
    jobs = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        date_val, jd_text, email = row[1], row[2], row[3]
        if not jd_text or not email:
            continue
        jobs.append({
            "date": date_val,
            "jd": str(jd_text).strip(),
            "email": str(email).strip(),
        })
        if max_rows and len(jobs) >= max_rows:
            break
    return jobs


# ── 2. 解析JD ─────────────────────────────────────────────
def parse_jd(jd_text):
    """从JD文本中提取公司名、职位名、邮件主题格式、简历命名格式"""
    # 公司名：取第一行破折号前的部分，或第一行
    first_line = jd_text.split("\n")[0].strip()
    company = re.split(r"[-–—]", first_line)[0].strip()
    # 去掉【】等符号
    company = re.sub(r"[【】\[\]]", "", company).strip()

    # 职位名：尝试匹配常见模式
    pos_patterns = [
        r"招聘(.{2,20}?)(?:实习生|助理|分析师)",
        r"[-–—]\s*(.{2,20}?)(?:实习生|助理|分析师)",
        r"【(.{2,20}?)(?:实习生|助理|分析师)",
    ]
    position = "投资实习生"
    for p in pos_patterns:
        m = re.search(p, first_line)
        if m:
            position = m.group(1).strip() + ("实习生" if "实习生" in jd_text[:100] else "")
            break

    # 邮件主题格式（只取到换行或"简历"字样前）
    subject_fmt = None
    for pattern in [r"邮件(?:主题|标题)[：:]\s*(.+)", r"邮件及简历命名格式[：:]\s*(.+)"]:
        m = re.search(pattern, jd_text)
        if m:
            raw = m.group(1).strip().split("\n")[0]
            # 截断"简历标题："后面的内容
            raw = re.split(r"简历(?:标题|命名)[：:]", raw)[0].strip()
            subject_fmt = raw
            break

    # 简历命名格式
    resume_fmt = None
    for pattern in [r"简历(?:命名|标题)[：:]\s*(.+)", r"简历及邮件(?:命名|标题)[：:]\s*(.+)",
                    r"邮件&简历命名[】\]]\s*(.+)"]:
        m = re.search(pattern, jd_text)
        if m:
            resume_fmt = m.group(1).strip().split("\n")[0]
            break

    return company, position, subject_fmt, resume_fmt


# ── 3. 判断方向 ────────────────────────────────────────────
def detect_focus(jd_text):
    jd_lower = jd_text.lower()
    if any(k in jd_lower for k in ["blockchain", "defi", "crypto", "web3", "nft", "链上"]):
        return "web3"
    if any(k in jd_lower for k in ["business development", "bd", "partnership", "合作"]):
        return "bd"
    if any(k in jd_lower for k in ["quant", "python", "量化", "数据分析", "modeling"]):
        return "quant"
    return "investment"  # 默认投研


FOCUS_HIGHLIGHTS = {
    "web3": "在0xU社区深度参与DeFi/链上交易，并担任WhaleRyder BD，具备Web3实战经验",
    "investment": "在水木清华基金完成投后实习，跟踪150+项目，撰写200+页运营报告",
    "quant": "具备Python/R面板数据分析经验，熟练使用Wind/CSMAR，完成绿色信贷量化研究",
    "bd": "担任WhaleRyder BD及Finternet峰会志愿者组长，具备多语言商务沟通能力",
}


# ── 4. 生成邮件正文（100字以内）────────────────────────────
def generate_email_body(company, position, focus):
    highlight = FOCUS_HIGHLIGHTS.get(focus, FOCUS_HIGHLIGHTS["investment"])
    body = (
        f"您好！\n\n"
        f"我是{SCHOOL}{MAJOR}大三学生{NAME}，对贵司{position}岗位非常感兴趣。"
        f"{highlight}，与岗位要求高度匹配。"
        f"可于{INTERN_PERIOD}全职实习，期待进一步交流。\n\n"
        f"{NAME}"
    )
    return body


# ── 5. 生成邮件主题 ───────────────────────────────────────
def generate_subject(subject_fmt, company, position):
    if not subject_fmt:
        return f"求职申请-{NAME}-{SCHOOL}-{position}"
    # 只取第一行，截断多余内容
    s = subject_fmt.split("\n")[0].strip()
    # 截断"入职时间"后面的换行残留
    s = re.split(r"入职时间\s*$", s)[0] + ("入职时间2025年7月" if s.endswith("入职时间") else "")
    replacements = {
        "最早可入职时间": "2025年7月", "可入职时间": "2025年7月",
        "最早可\n入职时间": "2025年7月", "最早可": "2025年7月",
        "姓名": NAME, "名字": NAME,
        "学校": SCHOOL, "本科学校": SCHOOL,
        "专业": MAJOR,
        "岗位": position, "职位": position,
        "公司": company,
        "毕业时间": "2026年6月", "毕业年份": "2026",
        "一周几天": "5天",
        "硕士学校（若有）": "",
    }
    for k, v in replacements.items():
        s = s.replace(k, v)
    s = re.sub(r"[-–—]{2,}", "-", s)
    return s.strip()


# ── 6. 生成简历文件名 ─────────────────────────────────────
def generate_resume_name(resume_fmt, company, position):
    if not resume_fmt:
        return f"{NAME}_简历_{position}_{company}.pdf"
    # 只取第一行，避免把JD后续内容带入文件名
    s = resume_fmt.split("\n")[0].strip()
    # 截断括号内的说明文字（如"（可脱敏）"）
    s = re.sub(r"[（(].*", "", s).strip()
    replacements = {
        "姓名": NAME, "名字": NAME,
        "学校": SCHOOL, "本科/研究生学校及专业": f"{SCHOOL}{MAJOR}",
        "本科学校": SCHOOL,
        "专业": MAJOR,
        "岗位": position, "职位": position,
        "投递岗位": position,
        "毕业时间": "2026年6月",
        "到岗时间": "2025年7月",
        "周实习X天": "周实习5天",
        "实习X个月": "实习4个月",
        "年级": "大三",
        "FA分析师助理": position,
        "人才战略": position,
    }
    for k, v in replacements.items():
        s = s.replace(k, v)
    # 清理非法文件名字符
    s = re.sub(r"[/\\:*?\"<>|]", "-", s)
    s = re.sub(r"\s+", "", s)
    if not s.endswith(".pdf"):
        s += ".pdf"
    return s


# ── 7. 发送邮件 ───────────────────────────────────────────
def send_email(to_addr, subject, body, pdf_path, pdf_name):
    msg = MIMEMultipart()
    msg["From"] = QQ_EMAIL
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    # 附加PDF（用RFC5987编码处理中文文件名）
    from email.header import Header
    with open(pdf_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    encoded_name = Header(pdf_name, "utf-8").encode()
    part.add_header("Content-Disposition", "attachment", filename=("utf-8", "", pdf_name))
    msg.attach(part)

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(QQ_EMAIL, QQ_AUTH_CODE)
        server.sendmail(QQ_EMAIL, to_addr, msg.as_string())


# ── 8. 生成HTML预览 ───────────────────────────────────────
def save_preview(job_dir, to_addr, subject, body, pdf_name):
    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>body{{font-family:sans-serif;max-width:700px;margin:40px auto;color:#333}}
.field{{margin:8px 0}}.label{{font-weight:bold;color:#555}}
.body-box{{background:#f9f9f9;padding:16px;border-radius:6px;white-space:pre-wrap;line-height:1.8}}
</style></head><body>
<h2>邮件预览</h2>
<div class="field"><span class="label">收件人：</span>{to_addr}</div>
<div class="field"><span class="label">主题：</span>{subject}</div>
<div class="field"><span class="label">附件：</span>{pdf_name}</div>
<hr>
<div class="body-box">{body}</div>
</body></html>"""
    with open(os.path.join(job_dir, "email_preview.html"), "w", encoding="utf-8") as f:
        f.write(html)


# ── 主流程 ────────────────────────────────────────────────
def run(mode="preview", max_rows=None):
    """
    mode: "preview" 只生成文件不发送
          "send"    生成并发送
    max_rows: 限制处理行数，None=全部
    """
    print(f"\n{'='*50}")
    print(f"AI求职助手启动  模式={mode}  时间={datetime.now().strftime('%H:%M:%S')}")
    print(f"{'='*50}\n")

    jobs = load_jobs(EXCEL_FILE, max_rows=max_rows)
    print(f"共读取 {len(jobs)} 个岗位\n")

    results = []
    for i, job in enumerate(jobs, 1):
        company, position, subject_fmt, resume_fmt = parse_jd(job["jd"])
        focus = detect_focus(job["jd"])
        subject = generate_subject(subject_fmt, company, position)
        body = generate_email_body(company, position, focus)
        pdf_name = generate_resume_name(resume_fmt, company, position)
        to_addr = job["email"]

        # 创建岗位目录
        safe_company = re.sub(r"[^\w\u4e00-\u9fff]", "_", company)[:20]
        job_dir = os.path.join(OUTPUT_DIR, f"job_{i:02d}_{safe_company}")
        os.makedirs(job_dir, exist_ok=True)

        # 复制简历并重命名
        pdf_dest = os.path.join(job_dir, pdf_name)
        shutil.copy2(RESUME_PDF, pdf_dest)

        # 生成HTML预览
        save_preview(job_dir, to_addr, subject, body, pdf_name)

        status = "已生成"
        error = None

        if mode == "send":
            try:
                send_email(to_addr, subject, body, pdf_dest, pdf_name)
                status = "已发送"
                print(f"[{i:02d}] ✅ {company} → {to_addr[:4]}***")
                time.sleep(1.5)
            except Exception as e:
                status = f"发送失败: {e}"
                error = str(e)
                print(f"[{i:02d}] ❌ {company} → {error}")
        else:
            print(f"[{i:02d}] 📄 {company} | {position} | {to_addr[:4]}***")

        results.append({
            "n": i, "company": company, "position": position,
            "email": to_addr, "subject": subject,
            "pdf_name": pdf_name, "focus": focus,
            "status": status, "body": body,
        })

    # 生成汇总报告
    save_report(results, mode)
    print(f"\n完成！输出目录：{OUTPUT_DIR}/")
    return results


def save_report(results, mode):
    lines = [
        f"# 求职助手汇总报告",
        f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}  模式：{mode}",
        f"共处理：{len(results)} 个岗位\n",
        "| # | 公司 | 职位 | 邮箱 | 简历文件名 | 状态 |",
        "|---|------|------|------|-----------|------|",
    ]
    for r in results:
        email_masked = r["email"][:4] + "***"
        lines.append(f"| {r['n']} | {r['company']} | {r['position']} | {email_masked} | {r['pdf_name']} | {r['status']} |")

    lines.append("\n---\n## 各岗位邮件正文\n")
    for r in results:
        lines.append(f"### {r['n']}. {r['company']} - {r['position']}")
        lines.append(f"**主题**：{r['subject']}\n")
        lines.append(f"```\n{r['body']}\n```\n")

    with open(os.path.join(OUTPUT_DIR, "summary_report.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


if __name__ == "__main__":
    import sys
    mode = sys.argv[1] if len(sys.argv) > 1 else "preview"
    max_rows = int(sys.argv[2]) if len(sys.argv) > 2 else None
    run(mode=mode, max_rows=max_rows)
