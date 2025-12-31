import os
import time
from datetime import datetime, timedelta, timezone
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
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

# Email Templates - Clock In
EMAIL_SUBJECT_IN_SUCCESS = "⚔️ 【社畜形態・覺醒】勞動神鷹看見你了"
EMAIL_BODY_IN_SUCCESS = """打卡成功。

🦅 勞動神鷹已注視此行為，
你的出勤被記錄於今日的時間軸。

請安心開始勞動。"""

EMAIL_SUBJECT_IN_FAILURE = "☠️ 【社畜形態・覺醒失敗】你就是勞工之光"
EMAIL_BODY_IN_FAILURE = """打卡未完成。

✨ 你仍是「勞工之光」，
但尚未被勞動諸神正式承認。

請重新嘗試覺醒。"""

# Email Templates - Clock Out
EMAIL_SUBJECT_OUT_SUCCESS = "🕊️ 【回歸凡人・解脫】勞動神鷹准許你離去"
EMAIL_BODY_OUT_SUCCESS = """簽退成功。

🦅 勞動神鷹收回了他的注視，
你重新掌握了你的時間。

請享受你的自由時光。"""

EMAIL_SUBJECT_OUT_FAILURE = "🌑 【封印解除失敗】神鷹拒絕放行"
EMAIL_BODY_OUT_FAILURE = """簽退未完成。

⛓️ 枷鎖依然沉重，
神鷹認為你的勞動尚不足以換取自由。

請嘗試再次突破。"""

# 2026 台灣國定假日 (僅列出週一至週五需放假的日子)
TAIWAN_HOLIDAYS_2026 = [
    "2026-01-01", # 元旦
    "2026-02-16", # 春節 (除夕)
    "2026-02-17", # 春節 (初一)
    "2026-02-18", # 春節 (初二)
    "2026-02-19", # 春節 (初三)
    "2026-02-20", # 春節 (補假)
    "2026-02-27", # 二二八補假 (2/28是週六)
    "2026-04-03", # 兒童節補假 (4/4是週六)
    "2026-04-06", # 清明節補假 (4/5是週日)
    "2026-05-01", # 勞動節
    "2026-06-19", # 端午節
    "2026-09-25", # 中秋節
    "2026-09-28", # 教師節 (週一)
    "2026-10-09", # 國慶日補假 (10/10是週六)
    "2026-10-26", # 台灣光復節補假 (10/25是週日)
    "2026-12-25", # 行憲紀念日 (週五)
]

def is_workday():
    """
    判斷今天是否為工作日 (以台北時間 UTC+8 為準)。
    1. 排除週六 (5) 與週日 (6)
    2. 排除 TAIWAN_HOLIDAYS_2026 名單中的日子
    """
    # 強制設定為台北時間 (UTC+8)
    tz_taiwan = timezone(timedelta(hours=8))
    now = datetime.now(tz_taiwan)
    today_str = now.strftime("%Y-%m-%d")
    weekday = now.weekday()  # Monday is 0, Sunday is 6

    # 1. 判斷是否為週末
    if weekday >= 5:
        print(f"今天是 {today_str} (週{'六' if weekday==5 else '日'})，休假不打卡。")
        return False

    # 2. 判斷是否為國定假日
    if today_str in TAIWAN_HOLIDAYS_2026:
        print(f"今天是 {today_str} (國定假日)，休假不打卡。")
        return False

    print(f"今天是 {today_str}，為工作日，準備執行打卡。")
    return True

def send_email(subject, body, image_path=None):
    """
    Sends an email notification via Gmail SMTP (HTML format with optional embedded image).
    """
    if not all([EMAIL_USER, EMAIL_APP_PASSWORD, RECEIVER_EMAIL]):
        print("郵件設定不完整，跳過郵件發送。")
        return

    try:
        # Create message container
        msg = MIMEMultipart('related')
        msg['Subject'] = Header(subject, 'utf-8')
        msg['From'] = f"EagleGod_Labor 勞動神鷹 <{EMAIL_USER}>"
        msg['To'] = RECEIVER_EMAIL

        # Build HTML body
        # Convert newline to <br> for HTML and use CID for image
        html_body = body.replace("\n", "<br>")
        img_html = ""
        if image_path and os.path.exists(image_path):
            img_html = '<br><br><img src="cid:status_image" style="max-width: 100%; height: auto; border-radius: 10px; box-shadow: 0 4px 8px rgba(0,0,0,0.2);">'
        
        full_html = f"""
        <html>
            <body style="font-family: 'Microsoft JhengHei', sans-serif; line-height: 1.6; color: #333;">
                <div style="padding: 20px; border: 1px solid #eee; border-radius: 10px;">
                    {html_body}
                    {img_html}
                </div>
            </body>
        </html>
        """
        
        msg.attach(MIMEText(full_html, 'html', 'utf-8'))

        # Attach image if provided
        if image_path and os.path.exists(image_path):
            with open(image_path, 'rb') as f:
                img = MIMEImage(f.read())
                img.add_header('Content-ID', '<status_image>')
                img.add_header('Content-Disposition', 'inline', filename=os.path.basename(image_path))
                msg.attach(img)

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
