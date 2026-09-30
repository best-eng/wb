#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Браузерный режим: запросы к WB идут изнутри настоящей страницы.

Зачем. Wildberries закрывает card.wb.ru для внешних клиентов, но сайт сам
эти данные получает. Если запрос выполняется в контексте открытой вкладки
wildberries.ru, он неотличим от обычного — со всеми куками, заголовками
и токенами, которые страница выставляет себе сама.

Как. Берём уже установленный Edge (или Chrome), запускаем с портом отладки
и разговариваем с ним по CDP — протоколу, встроенному в любой Chromium.
Playwright не нужен: он тянет Node-драйвер и плохо пакуется в exe,
а тут хватает одного websocket-клиента.

Скорость. Запрос выполняется через fetch() на странице и принимает пачку
артикулов сразу, поэтому 866 штук — это три десятка запросов, а не 866
загрузок страниц.
"""

from __future__ import annotations

import json
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Sequence
from urllib.parse import urlencode, urlparse
from urllib.request import urlopen

WB_MAIN = "https://www.wildberries.ru/"
GEO_URL = "https://user-geo-data.wildberries.ru/get-geo-info"
CARD_URL = "https://card.wb.ru/cards/v2/detail"

# Пачка меньше, чем в обычном режиме: ответ едет через отладочный канал,
# и раздувать сообщения незачем.
BROWSER_BATCH = 30

WINDOWS_BROWSERS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]

LINUX_BROWSERS = [
    "/opt/pw-browsers/chromium-1194/chrome-linux/chrome",
    "/usr/bin/chromium",
    "/usr/bin/chromium-browser",
    "/usr/bin/google-chrome",
]


def find_browser(explicit: str = "") -> str:
    """Ищет установленный Chromium-браузер. Пусто — не нашли."""
    if explicit:
        return explicit if Path(explicit).exists() else ""

    candidates = WINDOWS_BROWSERS if os.name == "nt" else LINUX_BROWSERS
    for path in candidates:
        if Path(path).exists():
            return path

    for name in ("msedge", "chrome", "chromium", "chromium-browser"):
        found = shutil.which(name)
        if found:
            return found
    return ""


def free_port() -> int:
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


class CDPError(RuntimeError):
    pass


class CDP:
    """Минимальный клиент протокола отладки Chromium."""

    def __init__(self, ws_url: str, timeout: int = 60) -> None:
        import websocket  # локальный импорт: режим опционален

        self.ws = websocket.create_connection(
            ws_url, timeout=timeout, max_size=64 * 1024 * 1024
        )
        self._id = 0

    def call(self, method: str, params: dict | None = None, timeout: int = 60) -> dict:
        self._id += 1
        message_id = self._id
        self.ws.send(json.dumps({"id": message_id, "method": method,
                                 "params": params or {}}))

        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                raw = self.ws.recv()
            except Exception as exc:
                raise CDPError(f"связь с браузером оборвалась: {exc}") from exc
            if not raw:
                continue
            message = json.loads(raw)
            if message.get("id") != message_id:
                continue                      # события нас не интересуют
            if "error" in message:
                raise CDPError(str(message["error"]))
            return message.get("result", {})
        raise CDPError(f"браузер не ответил на {method}")

    def evaluate(self, expression: str, timeout: int = 60) -> Any:
        """Выполняет JS на странице и возвращает результат по значению."""
        result = self.call("Runtime.evaluate", {
            "expression": expression,
            "awaitPromise": True,
            "returnByValue": True,
        }, timeout=timeout)

        if result.get("exceptionDetails"):
            text = result["exceptionDetails"].get("text", "ошибка JS")
            raise CDPError(text)
        return (result.get("result") or {}).get("value")

    def close(self) -> None:
        try:
            self.ws.close()
        except Exception:
            pass


class BrowserClient:
    """Совместим по интерфейсу с WBClient, поэтому основной цикл не меняется."""

    def __init__(self, log, browser_path: str = "", headless: bool = True,
                 delay: float = 0.2) -> None:
        self.log = log
        self.browser_path = browser_path
        self.headless = headless
        self.delay = delay

        self.batch_size = BROWSER_BATCH
        self.dead_streak = 0
        self.dest = ""
        self.attempts: list = []
        self.warm_status = "не выполнялся"
        self.warm_cookies = ""

        self.process: subprocess.Popen | None = None
        self.profile_dir: str = ""
        self.cdp: CDP | None = None

    # -- запуск ------------------------------------------------------------

    def start(self) -> bool:
        exe = find_browser(self.browser_path)
        if not exe:
            self.log("Не найден браузер. Нужен Microsoft Edge или Google Chrome.")
            self.log("Edge есть в любой Windows 10 и 11; если он удалён,")
            self.log("укажите путь вручную в wb_config.json, ключ browser_path.")
            return False

        self.log(f"Браузер: {exe}")
        port = free_port()
        self.profile_dir = tempfile.mkdtemp(prefix="wb_browser_")

        args = [
            exe,
            f"--remote-debugging-port={port}",
            f"--user-data-dir={self.profile_dir}",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-background-networking",
            "--disable-extensions",
            "--window-size=1280,900",
            # Chromium с версии 111 отклоняет подключение к отладочному
            # порту, если origin не разрешён явно.
            "--remote-allow-origins=*",
        ]
        if self.headless:
            args.append("--headless=new")
        if os.name != "nt":
            args.append("--no-sandbox")
        args.append(WB_MAIN)

        try:
            self.process = subprocess.Popen(
                args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
        except Exception as exc:
            self.log(f"Не удалось запустить браузер: {exc}")
            return False

        ws_url = self._wait_for_page(port)
        if not ws_url:
            self.log("Браузер запустился, но не отдал вкладку для управления.")
            return False

        try:
            self.cdp = CDP(ws_url)
        except Exception as exc:
            self.log(f"Не удалось подключиться к браузеру: {exc}")
            return False

        return self._open_storefront()

    def _wait_for_page(self, port: int, timeout: int = 40) -> str:
        """Ждём, пока браузер поднимет отладочный порт и создаст вкладку."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            if self.process and self.process.poll() is not None:
                self.log("Браузер завершился сразу после запуска.")
                return ""
            try:
                with urlopen(f"http://127.0.0.1:{port}/json/list", timeout=3) as response:
                    targets = json.loads(response.read().decode("utf-8", "replace"))
                for target in targets:
                    if target.get("type") == "page" and target.get("webSocketDebuggerUrl"):
                        return target["webSocketDebuggerUrl"]
            except Exception:
                pass
            time.sleep(0.5)
        return ""

    def _open_storefront(self) -> bool:
        """Дожидаемся загрузки витрины: именно она выдаёт куки и токены."""
        assert self.cdp is not None
        deadline = time.time() + 60
        self.log("Открываю витрину wildberries.ru…")
        reported = 0.0

        while time.time() < deadline:
            # Молчание на минуту выглядит как зависание, поэтому отмечаемся.
            waited = 60 - (deadline - time.time())
            if waited - reported >= 15:
                reported = waited
                self.log(f"  всё ещё жду загрузку… {int(waited)} с")

            try:
                state = self.cdp.evaluate("document.readyState", timeout=15)
                host = self.cdp.evaluate("location.hostname", timeout=15)
            except CDPError:
                time.sleep(1)
                continue

            expected = urlparse(WB_MAIN).hostname or ""
            on_site = str(host) == expected or str(host).endswith("." + expected)
            if state in ("interactive", "complete") and on_site:
                # Странице нужно время, чтобы выставить свои куки.
                time.sleep(2)
                try:
                    cookies = self.cdp.evaluate("document.cookie", timeout=15) or ""
                except CDPError:
                    cookies = ""
                names = sorted({c.split("=")[0].strip()
                                for c in str(cookies).split(";") if "=" in c})
                self.warm_status = "открыта"
                self.warm_cookies = ", ".join(names)
                self.log(f"Витрина открыта, куки: {self.warm_cookies or 'нет'}")
                return True
            time.sleep(1)

        self.log("Витрина wildberries.ru не открылась за отведённое время.")
        self.log("Проверьте, открывается ли сайт в обычном браузере.")
        return False

    # -- запросы -----------------------------------------------------------

    def _page_fetch(self, url: str, timeout: int = 60) -> tuple[Any, str]:
        """Выполняет fetch() внутри страницы — то есть от имени самого сайта."""
        if self.cdp is None:
            return None, "браузер не запущен"

        script = """
        (async () => {
            try {
                const r = await fetch(%s, {credentials: 'include'});
                const t = await r.text();
                return {ok: true, status: r.status, body: t};
            } catch (e) {
                return {ok: false, status: 0, body: String(e)};
            }
        })()
        """ % json.dumps(url)

        try:
            result = self.cdp.evaluate(script, timeout=timeout)
        except CDPError as exc:
            return None, f"браузер: {exc}"

        if not isinstance(result, dict):
            return None, "браузер вернул неожиданный ответ"
        if not result.get("ok"):
            return None, f"fetch не удался: {str(result.get('body'))[:100]}"

        status = result.get("status")
        body = result.get("body") or ""
        if status != 200:
            self.log.detail(f"{url} -> {status}; тело: {body[:200]}")
            return None, f"HTTP {status}"

        try:
            return json.loads(body), ""
        except Exception:
            return None, "ответ не JSON"

    def detect_dest(self) -> str:
        """Код региона спрашиваем у WB через ту же страницу."""
        from urllib.parse import parse_qs

        url = GEO_URL + "?" + urlencode({
            "currency": "RUB", "latitude": "55.7522",
            "longitude": "37.6156", "locale": "ru",
        })
        payload, error = self._page_fetch(url, timeout=30)
        if error or not isinstance(payload, dict):
            self.log.detail(f"гео через браузер не дал dest: {error}")
            return ""

        xinfo = payload.get("xinfo")
        if isinstance(xinfo, str) and xinfo:
            dest = parse_qs(xinfo).get("dest", [""])[0].strip()
            if dest:
                return dest

        destinations = payload.get("destinations")
        if isinstance(destinations, list):
            positive = [d for d in destinations if isinstance(d, int) and d > 0]
            if positive:
                return str(max(positive))
        return ""

    def fetch(self, nms: Sequence[int], attempts: int = 0) -> tuple[list[dict], str]:
        """Пачка артикулов одним запросом от имени страницы."""
        if self.delay:
            time.sleep(self.delay)

        url = CARD_URL + "?" + urlencode({
            "appType": "1",
            "curr": "rub",
            "dest": self.dest,
            "spp": "30",
            "hide_dtype": "10",
            "ab_testing": "false",
            "lang": "ru",
            "nm": ";".join(str(nm) for nm in nms),
        })

        payload, error = self._page_fetch(url)
        if error:
            self.dead_streak += 1
            return [], error

        self.dead_streak = 0
        data = payload.get("data") if isinstance(payload, dict) else None
        if isinstance(data, dict) and isinstance(data.get("products"), list):
            return [p for p in data["products"] if isinstance(p, dict)], ""
        if isinstance(payload, dict) and isinstance(payload.get("products"), list):
            return [p for p in payload["products"] if isinstance(p, dict)], ""
        return [], ""

    def shrink(self, failed_size: int) -> None:
        if failed_size >= self.batch_size and self.batch_size > 5:
            self.batch_size = max(5, self.batch_size // 2)
            self.log(f"  уменьшаю пачку до {self.batch_size} шт")

    # -- завершение --------------------------------------------------------

    def close(self) -> None:
        if self.cdp is not None:
            self.cdp.close()
            self.cdp = None

        if self.process is not None:
            try:
                self.process.terminate()
                self.process.wait(timeout=10)
            except Exception:
                try:
                    self.process.kill()
                except Exception:
                    pass
            self.process = None

        if self.profile_dir:
            shutil.rmtree(self.profile_dir, ignore_errors=True)
            self.profile_dir = ""
