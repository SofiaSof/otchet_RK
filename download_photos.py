import subprocess
import time
import json
import os
import re
import requests
import websocket

import openpyxl


def get_page_target_ws():
    chrome_proc = subprocess.Popen([
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "--headless=new",
        "--remote-debugging-port=9222",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    for _ in range(30):
        try:
            resp = requests.get("http://localhost:9222/json", timeout=5)
            for t in resp.json():
                if t["type"] == "page":
                    ws_url = t["webSocketDebuggerUrl"]
                    ws = websocket.create_connection(ws_url, suppress_origin=True, timeout=30)
                    return ws, chrome_proc
        except Exception:
            time.sleep(0.5)

    chrome_proc.terminate()
    raise RuntimeError("Could not connect to Chrome CDP")


def navigate_and_wait(ws, url, wait_seconds=8):
    ws.send(json.dumps({"id": 1, "method": "Page.navigate", "params": {"url": url}}))
    ws.recv()
    time.sleep(wait_seconds)

    ws.send(json.dumps({"id": 2, "method": "Runtime.evaluate",
                         "params": {"expression": "document.readyState", "returnByValue": True}}))
    resp = json.loads(ws.recv())
    state = resp.get("result", {}).get("result", {}).get("value")
    if state != "complete":
        for _ in range(20):
            time.sleep(1)
            ws.send(json.dumps({"id": 3, "method": "Runtime.evaluate",
                                 "params": {"expression": "document.readyState", "returnByValue": True}}))
            resp = json.loads(ws.recv())
            state = resp.get("result", {}).get("result", {}).get("value")
            if state == "complete":
                break


def find_photo_urls(ws):
    script = """
    (function() {
        var results = [];
        var seen = {};
        document.querySelectorAll("img").forEach(function(img) {
            var src = img.src || "";
            var isPhoto = (
                src.includes("onlineapi.russoutdoor.ru") ||
                src.includes("russoutdoor.ru/photo") ||
                src.includes("CampaignGuaranteedPhotoReport") ||
                src.includes("GetPosterContent") ||
                (img.className && img.className.includes("thumbnail"))
            );
            if (isPhoto && src && !src.includes(".svg") && !seen[src]) {
                seen[src] = true;
                results.push(src);
            }
        });
        return JSON.stringify(results);
    })()
    """
    ws.send(json.dumps({"id": 10, "method": "Runtime.evaluate",
                         "params": {"expression": script, "returnByValue": True}}))
    resp = json.loads(ws.recv())
    return json.loads(resp.get("result", {}).get("result", {}).get("value", "[]"))


def download_image(url, folder, session):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://online.russoutdoor.ru/"
    }
    try:
        r = session.get(url, timeout=30, headers=headers, allow_redirects=True)
        if r.status_code == 200 and len(r.content) > 1000:
            ext = ".jpg"
            if "png" in r.headers.get("content-type", "") or url.endswith(".png"):
                ext = ".png"
            elif "webp" in r.headers.get("content-type", "") or url.endswith(".webp"):
                ext = ".webp"

            name = url.split("/")[-1].split("?")[0]
            name = re.sub(r'[<>:"/\\|?*]', '_', name)
            if not name or "." not in name:
                name = f"photo_{abs(hash(url))}{ext}"

            filepath = os.path.join(folder, name)
            with open(filepath, "wb") as f:
                f.write(r.content)
            return filepath
    except Exception as e:
        print(f"    Download error: {e}")
    return None


def sanitize_folder(name):
    name = re.sub(r'[<>:"/\\|?*]', '_', str(name))
    return name.strip() or "unnamed"


def main():
    excel_path = "test/Otchet_Samocat.xlsx"
    row_num = 8

    wb = openpyxl.load_workbook(excel_path)
    ws_excel = wb.active

    folder_name = sanitize_folder(ws_excel.cell(row=row_num, column=11).value or "")
    if not folder_name or folder_name == "unnamed":
        print(f"Row {row_num}: column K is empty")
        return

    url = None
    cell_s = ws_excel.cell(row=row_num, column=19)
    if cell_s.hyperlink:
        url = cell_s.hyperlink.target
    if not url:
        url = cell_s.value
    if not url:
        print(f"Row {row_num}: column S has no URL")
        return

    print(f"Row: {row_num}")
    print(f"URL: {url}")
    print(f"Folder: {folder_name}")

    output_folder = os.path.join("test", "photos", folder_name)
    os.makedirs(output_folder, exist_ok=True)

    ws_cdp = None
    chrome_proc = None
    session = requests.Session()

    try:
        ws_cdp, chrome_proc = get_page_target_ws()
        navigate_and_wait(ws_cdp, url, wait_seconds=6)
        photo_urls = find_photo_urls(ws_cdp)

        print(f"\nFound {len(photo_urls)} photo(s):")
        for u in photo_urls:
            print(f"  {u}")

        downloaded = 0
        for i, photo_url in enumerate(photo_urls, 1):
            filepath = download_image(photo_url, output_folder, session)
            if filepath:
                print(f"  [{i}] Saved: {os.path.basename(filepath)} ({len(open(filepath,'rb').read())} bytes)")
                downloaded += 1
            else:
                print(f"  [{i}] Failed: {photo_url}")

        print(f"\nDone. Downloaded {downloaded}/{len(photo_urls)} photos to:\n{output_folder}")

    finally:
        if ws_cdp:
            ws_cdp.close()
        if chrome_proc:
            chrome_proc.terminate()
            chrome_proc.wait()


if __name__ == "__main__":
    main()
