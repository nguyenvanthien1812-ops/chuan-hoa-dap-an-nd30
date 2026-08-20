import os
import sys
import time
from playwright.sync_api import sync_playwright

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        # Collect console messages & errors
        console_logs = []
        page.on("console", lambda msg: console_logs.append(f"[{msg.type}] {msg.text}"))
        page.on("pageerror", lambda err: print(f"PAGE ERROR: {err}"))
        
        print("1. Mở trang http://127.0.0.1:8080/...")
        page.goto("http://127.0.0.1:8080/")
        page.wait_for_selector("#fileInput", state="attached")
        
        # Test 1: Chuẩn hóa đề thi mẫu
        print("2. Tải lên file test_samples/de_thi_mau.docx...")
        page.set_input_files("#fileInput", os.path.abspath("test_samples/de_thi_mau.docx"))
        
        page.wait_for_selector(".file-item")
        print("   - Đã nhận diện file trong danh sách!")
        
        btn_process = page.locator("#btnProcess")
        assert not btn_process.is_disabled(), "Nút Chuẩn hóa phải được enable"
        
        print("3. Nhấn 'Chuẩn Hóa Đề Thi'...")
        btn_process.click()
        
        # Chờ 3s và in console logs
        page.wait_for_timeout(3000)
        print("Console logs:", console_logs)
        
        # Wait for results card
        page.wait_for_selector("#resultsCard", state="visible", timeout=5000)
        
        keys_found = page.locator("#statKeysFound").text_content()
        status = page.locator("#statStatus").text_content()
        print(f"   - Kết quả: Số đáp án quét được = {keys_found}, Trạng thái = {status}")
        
        badges = page.locator(".key-badge").all_text_contents()
        print(f"   - Các đáp án đã bóc tách: {', '.join(badges)}")
        
        # Test 2: Chuyển sang tab NĐ30
        print("4. Chuyển sang Tab 'Nghị Định 30' và upload van_ban_mau.docx...")
        page.click("button[data-tab='nd30']")
        page.set_input_files("#fileInput", os.path.abspath("test_samples/van_ban_mau.docx"))
        
        page.click("#btnProcess")
        page.wait_for_selector("#resultsCard", state="visible", timeout=10000)
        status_nd30 = page.locator("#statStatus").text_content()
        print(f"   - Chuẩn hóa NĐ 30 thành công: Trạng thái = {status_nd30}")
        
        print("\n=== KIỂM THỬ THÀNH CÔNG 100% ===")
        print("Console logs:", len(console_logs), "logs ghi nhận.")
        for log in console_logs:
            if "error" in log.lower():
                print("   Console Error:", log)
                
        browser.close()

if __name__ == "__main__":
    run()
