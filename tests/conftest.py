import functools
import http.server
import json
import socketserver
import threading
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SITE = ROOT / "site"


@pytest.fixture(scope="session")
def site_data():
    return json.loads((SITE / "data" / "nike-financials.json").read_text(encoding="utf-8"))


@pytest.fixture(scope="session")
def server():
    class QuietHandler(http.server.SimpleHTTPRequestHandler):
        def log_message(self, *args):
            pass

    handler = functools.partial(QuietHandler, directory=str(SITE))
    httpd = socketserver.TCPServer(("127.0.0.1", 0), handler)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    yield f"http://127.0.0.1:{httpd.server_address[1]}/"
    httpd.shutdown()


@pytest.fixture(scope="session")
def browser():
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        try:
            b = p.chromium.launch(channel="chrome")
        except Exception:
            b = p.chromium.launch()
        yield b
        b.close()


@pytest.fixture
def page(browser, server):
    ctx = browser.new_context(viewport={"width": 1440, "height": 1000}, accept_downloads=True)
    pg = ctx.new_page()
    pg.console_errors = []
    pg.on("console", lambda m: pg.console_errors.append(m.text) if m.type == "error" else None)
    pg.on("pageerror", lambda e: pg.console_errors.append(str(e)))

    def open_(query=""):
        pg.goto(server + query)
        pg.wait_for_selector('html[data-ready="true"]', timeout=15000)
        return pg

    pg.open = open_
    yield pg
    ctx.close()
