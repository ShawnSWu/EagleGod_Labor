import os
import time
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

# Load environment variables
load_dotenv()

LOGIN_URL = os.getenv("LOGIN_URL")
ACCOUNT = os.getenv("ACCOUNT")
PASSWORD = os.getenv("PASSWORD")
CHECKIN_URL = os.getenv("CHECKIN_URL", "https://erp6.aoacloud.com.tw/HR/HRHB003S00.aspx")

# Email Settings
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_APP_PASSWORD = os.getenv("EMAIL_APP_PASSWORD")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")

# Email Templates
EMAIL_SUBJECT_SUCCESS = "⚔️ 【社畜形態・覺醒】勞動神鷹看見你了"
EMAIL_BODY_SUCCESS = """打卡完成。

🦅 勞動神鷹已注視此行為，
你的出勤被記錄於今日的時間軸。

請安心開始勞動。"""

EMAIL_SUBJECT_FAILURE = "☠️ 【社畜形態・覺醒失敗】你就是勞工之光"
EMAIL_BODY_FAILURE = """打卡未完成。

✨ 你仍是「勞工之光」，
但尚未被勞動諸神正式承認。

請重新嘗試覺醒。"""

def send_email(subject, body):
    """
    Sends an email notification via Gmail SMTP.
    """
    if not all([EMAIL_USER, EMAIL_APP_PASSWORD, RECEIVER_EMAIL]):
        print("郵件設定不完整，跳過郵件發送。")
        return

    try:
        msg = MIMEText(body, 'plain', 'utf-8')
        msg['Subject'] = Header(subject, 'utf-8')
        msg['From'] = f"EagleGod_Labor 勞動神鷹 <{EMAIL_USER}>"
        msg['To'] = RECEIVER_EMAIL

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL_USER, EMAIL_APP_PASSWORD)
            server.sendmail(EMAIL_USER, [RECEIVER_EMAIL], msg.as_string())
        print(f"通知郵件已寄送至 {RECEIVER_EMAIL}")
    except Exception as e:
        print(f"郵件發送失敗: {e}")

def get_logged_in_page(p, headless=True):
    print(f"啟動瀏覽器 (無頭模式: {headless})...")
    browser = p.chromium.launch(headless=headless)
    context = browser.new_context()
    page = context.new_page()

    try:
        print(f"正在前往登入頁面: {LOGIN_URL}")
        page.goto(LOGIN_URL)
        page.wait_for_selector("#login_name")
        page.fill("#login_name", ACCOUNT)
        page.fill("#password", PASSWORD)
        page.click("#loginBtn")
        page.wait_for_url("**/Default.aspx", timeout=15000)
        print("登入成功！")
        return browser, page
    except Exception as e:
        print(f"登入過程中發生錯誤: {e}")
        browser.close()
        raise e

def navigate_to_checkin(page):
    print(f"正在前往打卡頁面: {CHECKIN_URL}")
    page.goto(CHECKIN_URL)
    page.wait_for_selector("#btnclock1")
    return page
