#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Браузерный режим: чтение публичных страниц товаров.

Как это работает. Программа запускает установленный у вас Edge или Chrome
и открывает обычную страницу товара — ту же самую, что видит любой
посетитель, — после чего читает с неё название, бренд, цену и наличие.
Ничего, кроме открытия публичной страницы, здесь не делается.

Про остаток. Точное число Wildberries показывает не всегда: обычно только
когда товара мало («Осталось 3 шт»). Если числа на странице нет, а товар
продаётся, в отчёт идёт «в наличии» без количества и пометка в комментарии.
Это свойство сайта, а не программы.

Про скорость. Одна страница — примерно две-три секунды, поэтому список
на восемьсот с лишним артикулов занимает около получаса. Картинки и шрифты
не загружаются, это заметно ускоряет обход.

Управление браузером идёт по CDP — протоколу отладки, встроенному в любой
Chromium. Playwright не взят намеренно: он тянет Node-драйвер и плохо
пакуется в exe, а здесь хватает одного websocket-клиента, и сам браузер
не скачивается — используется уже установленный.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import socket
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Any, Sequence
from urllib.parse import urlparse
from urllib.request import urlopen

WB_MAIN = "https://www.wildberries.ru/"
PRODUCT_URL = "https://www.wildberries.ru/catalog/{nm}/detail.aspx"

# Пачка нужна только для отображения хода работы: страницы всё равно
# открываются по одной.
PAGE_BATCH = 10

# Картинки и шрифты для чтения текста не нужны, а грузятся долго.
BLOCKED_RESOURCES = [
    "*.jpg", "*.jpeg", "*.png", "*.webp", "*.gif", "*.svg",
    "*.woff", "*.woff2", "*.ttf", "*.mp4",
]

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

    for path in (WINDOWS_BROWSERS if os.name == "nt" else LINUX_BROWSERS):
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
        result = self.call("Runtime.evaluate", {
            "expression": expression,
            "awaitPromise": True,
            "returnByValue": True,
        }, timeout=timeout)

        if result.get("exceptionDetails"):
            raise CDPError(result["exceptionDetails"].get("text", "ошибка JS"))
        return (result.get("result") or {}).get("value")

    def close(self) -> None:
        try:
            self.ws.close()
        except Exception:
            pass


# Сбор данных со страницы. Вёрстка у WB меняется, поэтому собираем несколько
# признаков сразу, а разбираем их уже на стороне программы: так замена
# одного класса в разметке не ломает всё чтение.
EXTRACT_JS = r"""
(() => {
    const text = document.body ? document.body.innerText : '';
    const pick = (sel) => {
        const el = document.querySelector(sel);
        return el ? (el.innerText || '').trim() : '';
    };

    return {
        title: document.title || '',
        name: pick('h1'),
        brand: pick('.product-page__header-brand')
            || pick('[data-link*="brandName"]')
            || pick('.same-part-kt__header-link'),
        price: pick('.price-block__final-price')
            || pick('ins.price-block__final-price')
            || pick('.price-block__price'),
        seller: pick('[data-link*="supplierName"]')
             || pick('.seller-info__name'),
        text: text.slice(0, 4000),
        length: text.length,
    };
})()
"""

LEFT_RE = re.compile(r"[Оо]сталось\s+(\d+)\s*шт", re.IGNORECASE)
SOLD_OUT_RE = re.compile(
    r"нет в наличии|товар закончил|распродан|нет в продаже|товара нет",
    re.IGNORECASE,
)
NOT_FOUND_RE = re.compile(
    r"страница не найдена|такой страницы|ничего не найдено|товар не найден",
    re.IGNORECASE,
)
PRICE_RE = re.compile(r"(\d[\d\s\u00a0]*)\s*₽")


class BrowserClient:
    """Совместим по интерфейсу с WBClient, поэтому основной цикл не меняется."""

    def __init__(self, log, browser_path: str = "", headless: bool = True,
                 delay: float = 0.2) -> None:
        self.log = log
        self.browser_path = browser_path
        self.headless = headless
        self.delay = delay

        self.batch_size = PAGE_BATCH
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
            self.log("укажите путь в wb_config.json, ключ browser_path.")
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

        self._speed_up()
        return self._open_storefront()

    def _speed_up(self) -> None:
        """Отключаем картинки и шрифты: нам нужен только текст страницы."""
        if self.cdp is None:
            return
        try:
            self.cdp.call("Network.enable")
            self.cdp.call("Network.setBlockedURLs", {"urls": BLOCKED_RESOURCES})
            self.log.detail("загрузка картинок и шрифтов отключена")
        except CDPError as exc:
            self.log.detail(f"не удалось отключить лишние ресурсы: {exc}")

    def _wait_for_page(self, port: int, timeout: int = 40) -> str:
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
        """Дожидаемся витрины: заодно это проверка, что сайт вообще открыт."""
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

    # -- чтение страницы ---------------------------------------------------

    def _navigate(self, url: str, timeout: int = 40) -> bool:
        if self.cdp is None:
            return False
        try:
            self.cdp.call("Page.navigate", {"url": url}, timeout=timeout)
        except CDPError as exc:
            self.log.detail(f"навигация на {url}: {exc}")
            return False

        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                state = self.cdp.evaluate("document.readyState", timeout=10)
            except CDPError:
                time.sleep(0.4)
                continue
            if state == "complete":
                return True
            time.sleep(0.3)
        return False

    def _wait_for_card(self, timeout: int = 12) -> None:
        """Цена и наличие дорисовываются скриптом уже после загрузки.

        Ждём именно заголовок товара: объём текста для этого не годится —
        у коротких карточек его мало, и ожидание упирается в таймаут,
        превращая пару секунд на страницу в полтора десятка.
        """
        if self.cdp is None:
            return

        deadline = time.time() + timeout
        seen_title = False
        while time.time() < deadline:
            try:
                state = self.cdp.evaluate(
                    "(() => ({"
                    " h1: !!document.querySelector('h1'),"
                    " price: !!document.querySelector("
                    "'.price-block__final-price, .price-block__price'),"
                    " len: document.body ? document.body.innerText.length : 0"
                    "}))()", timeout=10)
            except CDPError:
                return
            if not isinstance(state, dict):
                return

            if state.get("h1") or state.get("len", 0) > 200:
                # Заголовок есть; цене даём короткую фору дорисоваться.
                if state.get("price") or seen_title:
                    return
                seen_title = True
                time.sleep(0.4)
                continue
            time.sleep(0.3)

    def read_product(self, nm: int) -> dict | None:
        """Открывает публичную страницу товара и снимает с неё данные."""
        if not self._navigate(PRODUCT_URL.format(nm=nm)):
            return None
        self._wait_for_card()

        try:
            raw = self.cdp.evaluate(EXTRACT_JS, timeout=20) if self.cdp else None
        except CDPError as exc:
            self.log.detail(f"артикул {nm}: страница не прочитана: {exc}")
            return None
        if not isinstance(raw, dict):
            return None

        text = str(raw.get("text") or "")
        if NOT_FOUND_RE.search(text) or int(raw.get("length") or 0) < 120:
            self.log.detail(f"артикул {nm}: карточки нет")
            return {"id": nm, "_absent": True}

        left = LEFT_RE.search(text)
        if left:
            quantity, known = int(left.group(1)), True
        elif SOLD_OUT_RE.search(text):
            quantity, known = 0, True
        else:
            # Товар продаётся, но точного числа WB на странице не показывает.
            quantity, known = 0, False

        price = None
        match = PRICE_RE.search(str(raw.get("price") or "")) or PRICE_RE.search(text)
        if match:
            digits = re.sub(r"\D", "", match.group(1))
            if digits:
                price = float(digits)

        return {
            "id": nm,
            "name": str(raw.get("name") or raw.get("title") or "").strip()[:200],
            "brand": str(raw.get("brand") or "").strip()[:100],
            "supplier": str(raw.get("seller") or "").strip()[:100],
            "totalQuantity": quantity,
            "_stock_known": known,
            "_page_price": price,
        }

    def fetch(self, nms: Sequence[int], attempts: int = 0) -> tuple[list[dict], str]:
        """Страницы открываются по одной; пачка нужна для отображения хода."""
        products: list[dict] = []
        failures = 0

        for nm in nms:
            if self.delay:
                time.sleep(self.delay)
            product = self.read_product(nm)
            if product is None:
                failures += 1
                continue
            if product.get("_absent"):
                continue           # карточки нет — в отчёте «нет на сайте»
            products.append(product)

        if failures and not products:
            self.dead_streak += 1
            return [], "страница не открылась"

        self.dead_streak = 0
        return products, ""

    def shrink(self, failed_size: int) -> None:
        """Страницы читаются по одной, дробить нечего."""
        return

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
