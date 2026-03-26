import requests
from bs4 import BeautifulSoup
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.charset import Charset, QP
from dotenv import load_dotenv
import os

# =========================================
# 1. 你的查詢 URL（你原來那個）
# =========================================
URL = "https://web.pcc.gov.tw/prkms/tender/common/basic/readTenderBasic?pageSize=&firstSearch=true&searchType=basic&isBinding=N&isLogIn=N&level_1=on&orgName=&orgId=&tenderName=%E7%A6%AE%E5%88%B8&tenderId=&tenderType=TENDER_DECLARATION&tenderWay=TENDER_WAY_ALL_DECLARATION&dateType=isNow&tenderStartDate=2026%2F03%2F19&tenderEndDate=2026%2F03%2F25&radProctrgCate=RAD_PROCTRG_CATE_2&policyAdvocacy="

# =========================================
# 2. Outlook SMTP 設定
# =========================================
load_dotenv()

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
OUTLOOK_EMAIL = os.getenv("OUTLOOK_EMAIL")
OUTLOOK_PASSWORD = os.getenv("OUTLOOK_PASSWORD")
TO_EMAIL = os.getenv("TO_EMAIL")

print("Email:", OUTLOOK_EMAIL)
print("Password length:", len(OUTLOOK_PASSWORD) if OUTLOOK_PASSWORD else "None or empty")

# =========================================
# 3. 建立 session 抓 HTML（確保抓到結果，不是查詢頁）
# =========================================
def fetch_html():
    session = requests.Session()
    headers = {"User-Agent": "Mozilla/5.0", "Referer": URL}

    session.get(URL, headers=headers)  # 建立 session
    r = session.get(URL, headers=headers)
    r.raise_for_status()
    return r.text


# =========================================
# 4. 抓「id=tpam」的結果表格（不修改 HTML！）
# =========================================
def extract_table_html(html):
    soup = BeautifulSoup(html, "html.parser")

    table = soup.find("table", {"id": "tpam"})
    if not table:
        return "<p>⚠ 未找到結果表格</p>"

    # ------------------------------
    # 1) 移除排序連結 (?開頭)
    # ------------------------------
    for a in table.find_all("a"):
        href = a["href"].strip()

        if href.startswith("?"):
            a.unwrap()
            continue

        # if not "/tpam?pk=" in href:
        #     # 組完整 URL
        #     full_url = "https://web.pcc.gov.tw" + href

        #     # 建立新的 <a>
        #     new_a = soup.new_tag("a", href=full_url)
        #     new_a.string = a.get_text(strip=True)

        #     # 替換舊的 <a>
        #     a.replace_with(new_a)

    # ------------------------------
    # 2) 找出所有真正的標案連結："/tpam?pk="
    # ------------------------------
    for a in table.find_all("a", href=True):
        href = a["href"].strip()

        if "/tpam?pk=" in href:
            
            span = a.find("span")
            case_name = ""
            if span and span.script and span.script.string:
                import re
                m = re.search(r'Geps3\.CNS\.pageCode2Img\("(.+?)"\)', span.script.string)
                if m:
                    case_name = m.group(1)
            if not case_name:
                case_name = a.get_text(strip=True)
            if not case_name:
                continue

            # 組完整 URL
            full_url = "https://web.pcc.gov.tw" + href

            # 建立新的 <a>
            new_a = soup.new_tag("a", href=full_url)
            new_a.string = case_name

            # 替換舊的 <a>
            a.replace_with(new_a)

    # ------------------------------
    # 3) 移除所有 <u>
    # ------------------------------
    for u in table.find_all("u"):
        u.unwrap()

    # ------------------------------
    # 4) 修正 h1 字體
    # ------------------------------
    for h1 in table.find_all("h1"):
        h1.unwrap()

    # ------------------------------
    # 5) 外框
    # ------------------------------
    table["style"] = "border:2px solid black;border-collapse:collapse; font-size:14px;"
    for cell in table.find_all(["td", "th"]):
        cell["style"] = cell.get("style", "") + "border:1px solid black;padding:4px;"

    return str(table)

# =========================================
# 5. 寄信（Header 強制 UTF‑8，避免 Windows 亂碼）
# =========================================
def send_email(html_table):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = "PCC 標案查詢結果（禮券 / 財務）"
    msg["From"] = OUTLOOK_EMAIL
    msg["To"] = TO_EMAIL

    # 設定 charset（避免 Windows 亂碼）
    charset = Charset("utf-8")
    charset.header_encoding = QP
    charset.body_encoding = QP
    msg.set_charset(charset)

    html_body = f"""
    <html>
    <body>
        <h2>PCC 標案查詢結果：禮券（財務）</h2>
        {html_table}
        <br><br>
        <p style="font-size:12px;color:#888;">此信件由自動排程寄出。</p>
    </body>
    </html>
    """

    msg.attach(MIMEText(html_body, "html", "utf-8"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(OUTLOOK_EMAIL, OUTLOOK_PASSWORD)
        server.sendmail(OUTLOOK_EMAIL, TO_EMAIL, msg.as_string())

    print("send email success")


# =========================================
# 6. 主程式
# =========================================
if __name__ == "__main__":
    html = fetch_html()
    table_html = extract_table_html(html)
    send_email(table_html)
