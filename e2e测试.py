# -*- coding: utf-8 -*-
"""
真实浏览器 E2E 测试
  启动真实 Flask 服务器 → 用 Chromium 打开页面 →
  模拟传 PO/校验自动预填 → 点生成 → 校验产出 PDF →
  传多文件 → 点合并 → 校验合并 PDF。
运行：  python3 e2e测试.py
"""
import os
import sys
import time
import socket
import subprocess
from pathlib import Path

import pypdf
from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent
PARENT = BASE.parent
OUTDIR = BASE / "输出单据"
PORT = 5057
URL = f"http://127.0.0.1:{PORT}"

PASS, FAIL = [], []
def check(name, cond):
    (PASS if cond else FAIL).append(name)
    print(("  ✅ " if cond else "  ❌ ") + name)


def wait_port(port, timeout=15):
    t0 = time.time()
    while time.time() - t0 < timeout:
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", port)) == 0:
                return True
        time.sleep(0.3)
    return False


def main():
    # 清掉上次的E2E产物
    for p in list(OUTDIR.glob("E2E*")) + list(OUTDIR.glob("*_img.pdf")):
        p.unlink()

    env = {**os.environ, "NO_BROWSER": "1", "FLASK_PORT": str(PORT)}
    server = subprocess.Popen([sys.executable, "app.py"], cwd=BASE, env=env,
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        if not wait_port(PORT):
            print("❌ 服务器没起来"); sys.exit(1)
        print(f"服务器已启动 {URL}\n")

        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            page = browser.new_page()
            errors = []
            page.on("console", lambda m: errors.append(m.text) if m.type == "error" else None)
            page.on("pageerror", lambda e: errors.append(str(e)))

            page.goto(URL)
            print("【场景1】打开页面")
            check("页面标题正确", "外贸单据" in page.title())
            check("左侧出现『生成发票』卡片", page.locator("text=生成发票").count() > 0)
            check("右侧出现『合并清关』卡片", page.locator("text=合并清关资料包").count() > 0)
            check("发票号输入框存在", page.locator('[data-k="发票号"]').count() == 1)

            print("\n【场景2】上传 Sika PO → 自动预填表单")
            page.set_input_files("#poInput", str(PARENT / "PO_4506776081.pdf"))
            # 等后端识别完、JS把数量填进去
            page.wait_for_function(
                """() => document.querySelector('[data-k="数量KGS"]').value !== ''""",
                timeout=10000)
            qty = page.input_value('[data-k="数量KGS"]')
            ref = page.input_value('[data-k="ReferenceNo"]')
            hs = page.input_value('[data-k="品名及HS编码"]')
            inco = page.input_value('[data-k="贸易术语"]')
            check(f"数量自动填入 (={qty})", qty == "26400")
            check(f"ReferenceNo自动填入 (={ref})", ref == "4506776081")
            check("品名/HS自动填入", "HS CODE" in hs)
            check(f"贸易术语自动填入 (={inco})", "CIF" in inco)
            check("提示『已自动预填』", "已自动预填" in page.locator("#genMsg").inner_text())

            print("\n【场景3】补发票号 → 点生成 → 校验产出")
            page.fill('[data-k="发票号"]', "E2E_SIKA")
            page.fill('[data-k="体积CBM"]', "30")
            page.click("#genBtn")
            page.wait_for_function(
                """() => document.querySelector('#genMsg').textContent.includes('生成成功')""",
                timeout=30000)
            check("页面提示『生成成功』", "生成成功" in page.locator("#genMsg").inner_text())
            check("出现下载链接", page.locator("#genLinks a").count() >= 1)
            check("PDF预览iframe已加载", "E2E_SIKA" in (page.locator("#genFrame").get_attribute("src") or ""))
            pdf = OUTDIR / "E2E_SIKA_PI-CI-PL.pdf"
            xlsx = OUTDIR / "E2E_SIKA_PI-CI-PL.xlsx"
            check("Excel文件已生成", xlsx.exists())
            check("PDF文件已生成", pdf.exists())
            if pdf.exists():
                check("PDF为3页(PI/CI/PL)", len(pypdf.PdfReader(str(pdf)).pages) == 3)

            print("\n【场景4】传多文件(含图片) → 合并清关包")
            jpg = next(PARENT.glob("*.jpg"))
            page.set_input_files("#mgInput", [
                str(PARENT / "A91GX06249_保单_11114046800275568876.pdf"),
                str(PARENT / "2248中成-性能单10000个.pdf"),
                str(jpg),  # 图片：测试图片合并
            ])
            page.wait_for_function(
                """() => document.querySelectorAll('#mgList li').length === 3""", timeout=8000)
            check("文件列表显示3个文件(含图片)", page.locator("#mgList li").count() == 3)
            # 测试删除：删掉再确认剩2，然后重新加回图片
            page.locator("#mgList li").nth(2).locator("button", has_text="✕").click()
            check("删除后列表剩2个", page.locator("#mgList li").count() == 2)
            page.set_input_files("#mgInput", [str(jpg)])
            page.wait_for_function("""() => document.querySelectorAll('#mgList li').length === 3""", timeout=8000)
            page.fill("#mgName", "E2E_清关包")
            page.click("#mgBtn")
            page.wait_for_function(
                """() => document.querySelector('#mgMsg').textContent.includes('合并完成')""",
                timeout=30000)
            mtext = page.locator("#mgMsg").inner_text()
            check("页面提示『合并完成』", "合并完成" in mtext)
            check("图片未被跳过(已支持图片合并)", "跳过" not in mtext)
            mpdf = OUTDIR / "E2E_清关包.pdf"
            check("合并PDF已生成", mpdf.exists())
            if mpdf.exists():
                check("合并PDF页数>=3(含图片页)", len(pypdf.PdfReader(str(mpdf)).pages) >= 3)

            print("\n【场景5】危化品三件套：改字段 → 生成 → 校验")
            page.fill('[data-c="出口国"]', "越南")
            page.fill('[data-c="合同号"]', "E2EHX001")
            page.fill('[data-c="桶数"]', "100")
            page.fill('[data-c="总重量"]', "20000")
            page.click("#chemBtn")
            page.wait_for_function(
                """() => document.querySelector('#chemMsg').textContent.includes('已生成')""",
                timeout=40000)
            check("提示『已生成』", "已生成" in page.locator("#chemMsg").inner_text())
            check("出现3组下载链接区块", page.locator("#chemLinks b").count() == 3)
            check("下载链接>=3条", page.locator("#chemLinks a").count() >= 3)
            import subprocess as _sp
            ok_replace = False
            sf = OUTDIR / "E2EHX001_符合性声明.docx"
            check("符合性声明docx已生成", sf.exists())
            if sf.exists():
                t = _sp.run(["textutil","-convert","txt","-stdout",str(sf)],capture_output=True,text=True).stdout
                ok_replace = ("越南" in t and "共100桶20000千克" in t and "印度" not in t)
            check("符合性声明内容已正确替换(越南/100桶/20000,且无残留印度)", ok_replace)
            check("责任声明docx已生成", (OUTDIR/"E2EHX001_责任自负声明.docx").exists())
            check("商检厂检单docx已生成", (OUTDIR/"E2EHX001_商检厂检单.docx").exists())

            print("\n【场景6】前端无JS报错")
            real_errs = [e for e in errors if "padding" not in e.lower()]
            check(f"无控制台错误 ({len(real_errs)}个)", len(real_errs) == 0)
            if real_errs:
                for e in real_errs[:5]:
                    print("     JS错误:", e[:120])

            browser.close()
    finally:
        server.terminate()
        try: server.wait(timeout=5)
        except Exception: server.kill()

    print("\n" + "=" * 44)
    print(f" E2E 结果：通过 {len(PASS)} / 失败 {len(FAIL)}")
    if FAIL:
        print(" 失败项：")
        for f in FAIL: print("   -", f)
    print("=" * 44)
    sys.exit(1 if FAIL else 0)


if __name__ == "__main__":
    main()
