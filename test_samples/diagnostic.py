import base64
import json
from playwright.sync_api import sync_playwright

def run():
    with open("test_samples/de_thi_mau.docx", "rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")
        
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto("http://127.0.0.1:8080/")
        
        js_code = f"""async () => {{
            try {{
                const b64Data = "{b64}";
                const binary = atob(b64Data);
                const bytes = new Uint8Array(binary.length);
                for (let i = 0; i < binary.length; i++) {{
                    bytes[i] = binary.charCodeAt(i);
                }}
                const zip = await JSZip.loadAsync(bytes.buffer);
                
                const normRes = await window.Normalizer.normalizeDocx(zip, {{
                    auto_key: true,
                    bold_as_correct: true,
                    choice_layout: 'auto'
                }});
                return {{ success: true, normRes }};
            }} catch (e) {{
                return {{ success: false, error: e.toString(), stack: e.stack }};
            }}
        }}"""
        
        res = page.evaluate(js_code)
        print("DIAGNOSTIC RESULT:", json.dumps(res, indent=2, ensure_ascii=False))
        browser.close()

if __name__ == "__main__":
    run()
