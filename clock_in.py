import time
import random
from playwright.sync_api import sync_playwright
from utils import (
    get_logged_in_page, 
    navigate_to_checkin, 
    send_email, 
    is_workday,
    EMAIL_SUBJECT_IN_SUCCESS, 
    EMAIL_BODY_IN_SUCCESS, 
    EMAIL_SUBJECT_IN_FAILURE, 
    EMAIL_BODY_IN_FAILURE
)
import sys

def clock_in(headless=True):
    with sync_playwright() as p:
        browser = None
        status_msg = "等待回應超時"
        success = False
        try:
            # 檢查是否為工作日
            if not is_workday():
                return
            
            # 加入隨機延遲 (1 到 300 秒)
            # 移到最前面，防止登入後因延遲導致 Session 過期
            delay = random.randint(1, 300)
            print(f"隨機延遲 {delay} 秒後開始執行簽到流程...")
            time.sleep(delay)

            browser, page = get_logged_in_page(p, headless=headless)
            navigate_to_checkin(page)
            
            print("正在點擊 [簽到] 按鈕...")
            page.click("#btnclock1")
            
            # 改用等待 .sweet-alert 元素出現，而不是 dialog 監聽器
            try:
                print("等待系統回應訊息...")
                # 等待 SweetAlert 彈窗出現 (最多等 15 秒)
                page.wait_for_selector(".sweet-alert", timeout=15000)
                
                # 抓取彈窗訊息 (通常在 h2 或 p 標籤中)
                status_msg = page.inner_text(".sweet-alert h2")
                print(f"系統訊息: {status_msg}")
                
                # 判斷關鍵字
                if any(k in status_msg for k in ["成功", "重複", "已簽到", "已簽退"]):
                    success = True
                
                # 點擊彈窗的確定按鈕 (如果有)
                if page.is_visible(".confirm"):
                    page.click(".confirm")
                    
            except Exception as e:
                print(f"等待彈窗訊息時發生異常: {e}")
                # 有些系統可能直接跳轉或沒跳彈窗，這裡做個備選預案
                status_msg = "未捕捉到彈窗訊息，請檢查系統記錄"

            print("簽到程序執行完畢。")
            
            if success:
                body = f"{EMAIL_BODY_IN_SUCCESS}\n\n(系統訊息: {status_msg}, 延遲: {delay}秒)"
                send_email(EMAIL_SUBJECT_IN_SUCCESS, body, image_path="assets/clock_in.png")
            else:
                body = f"{EMAIL_BODY_IN_FAILURE}\n\n(系統訊息: {status_msg}, 延遲: {delay}秒)"
                send_email(EMAIL_SUBJECT_IN_FAILURE, body)

        except Exception as e:
            error_msg = f"{EMAIL_BODY_IN_FAILURE}\n\n(錯誤原因: {e})"
            print(f"簽到發生錯誤: {e}")
            send_email(EMAIL_SUBJECT_IN_FAILURE, error_msg)
        finally:
            if browser:
                browser.close()
                print("瀏覽器已關閉。")

if __name__ == "__main__":
    is_headless = "--head" not in sys.argv
    clock_in(headless=is_headless)
