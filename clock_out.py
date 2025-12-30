import time
import random
from playwright.sync_api import sync_playwright
from utils import (
    get_logged_in_page, 
    navigate_to_checkin, 
    send_email, 
    EMAIL_SUBJECT_SUCCESS, 
    EMAIL_BODY_SUCCESS, 
    EMAIL_SUBJECT_FAILURE, 
    EMAIL_BODY_FAILURE
)
import sys

def clock_out(headless=True):
    with sync_playwright() as p:
        browser = None
        status_msg = "等待回應超時"
        success = False
        try:
            browser, page = get_logged_in_page(p, headless=headless)
            navigate_to_checkin(page)
            
            # 加入隨機延遲 (1 到 600 秒)
            delay = random.randint(1, 600)
            print(f"隨機延遲 {delay} 秒後執行簽退...")
            time.sleep(delay)

            print("正在點擊 [簽退] 按鈕...")
            page.click("#btnclock2")
            
            # 偵測 SweetAlert 彈窗
            try:
                print("等待系統回應訊息...")
                page.wait_for_selector(".sweet-alert", timeout=15000)
                status_msg = page.inner_text(".sweet-alert h2")
                print(f"系統訊息: {status_msg}")
                
                if any(k in status_msg for k in ["成功", "重複", "已簽到", "已簽退"]):
                    success = True
                
                if page.is_visible(".confirm"):
                    page.click(".confirm")
            except Exception as e:
                print(f"等待彈窗訊息時發生異常: {e}")
                status_msg = "未捕捉到彈窗訊息，請檢查系統記錄"

            print("簽退程序執行完畢。")
            
            if success:
                success_body = EMAIL_BODY_SUCCESS.replace("請安心開始勞動。", "請安心結束勞動，回歸自由。")
                body = f"{success_body}\n\n(系統訊息: {status_msg}, 延遲: {delay}秒)"
                send_email(EMAIL_SUBJECT_SUCCESS, body)
            else:
                body = f"{EMAIL_BODY_FAILURE}\n\n(系統訊息: {status_msg}, 延遲: {delay}秒)"
                send_email(EMAIL_SUBJECT_FAILURE, body)

        except Exception as e:
            error_msg = f"{EMAIL_BODY_FAILURE}\n\n(錯誤原因: {e})"
            print(f"簽退發生錯誤: {e}")
            send_email(EMAIL_SUBJECT_FAILURE, error_msg)
        finally:
            if browser:
                browser.close()
                print("瀏覽器已關閉。")

if __name__ == "__main__":
    is_headless = "--head" not in sys.argv
    clock_out(headless=is_headless)
