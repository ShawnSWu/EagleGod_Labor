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

def clock_in(headless=True):
    with sync_playwright() as p:
        browser = None
        status_msg = "未知狀態"
        success = False
        try:
            browser, page = get_logged_in_page(p, headless=headless)
            navigate_to_checkin(page)
            
            # 設定彈窗監聽器
            def handle_dialog(dialog):
                nonlocal status_msg, success
                status_msg = dialog.message
                print(f"系統訊息: {status_msg}")
                # 簡單判斷關鍵字來決定是否成功
                if "成功" in status_msg or "重複" in status_msg or "已簽到" in status_msg:
                    success = True
                dialog.accept()

            page.on("dialog", handle_dialog)

            # 加入隨機延遲 (1 到 600 秒)
            delay = random.randint(1, 600)
            print(f"隨機延遲 {delay} 秒後執行簽到...")
            time.sleep(delay)

            print("正在點擊 [簽到] 按鈕...")
            page.click("#btnclock1")
            
            page.wait_for_timeout(3000)
            print("簽到程序執行完畢。")
            
            if success:
                body = f"{EMAIL_BODY_SUCCESS}\n\n(系統訊息: {status_msg}, 延遲: {delay}秒)"
                send_email(EMAIL_SUBJECT_SUCCESS, body)
            else:
                body = f"{EMAIL_BODY_FAILURE}\n\n(系統訊息: {status_msg}, 延遲: {delay}秒)"
                send_email(EMAIL_SUBJECT_FAILURE, body)

        except Exception as e:
            error_msg = f"{EMAIL_BODY_FAILURE}\n\n(錯誤原因: {e})"
            print(f"簽到發生錯誤: {e}")
            send_email(EMAIL_SUBJECT_FAILURE, error_msg)
        finally:
            if browser:
                browser.close()
                print("瀏覽器已關閉。")

if __name__ == "__main__":
    is_headless = "--head" not in sys.argv
    clock_in(headless=is_headless)
