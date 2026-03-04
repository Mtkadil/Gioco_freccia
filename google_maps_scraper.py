#!/usr/bin/env python3
"""Scraper Google Maps con Playwright.

Funzionalità:
- Cerca "società edilizie" in una località scelta dall'utente.
- Esegue lo scrolling della barra laterale dei risultati fino a fine elenco.
- Estrae Nome, Indirizzo e Telefono dei risultati visibili.
- Tenta di escludere risultati sponsorizzati (annunci/sponsored).
- Salva i dati in CSV o XLSX.
"""

import argparse
import csv
import random
import time
from pathlib import Path
from typing import Dict, List

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


def random_sleep(min_s: float = 0.8, max_s: float = 2.0) -> None:
    """Pausa casuale per ridurre pattern ripetitivi da bot."""
    time.sleep(random.uniform(min_s, max_s))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Estrae aziende da Google Maps con Playwright"
    )
    parser.add_argument(
        "--location",
        required=True,
        help="Località di ricerca, es. 'Milano'",
    )
    parser.add_argument(
        "--query",
        default="società edilizie",
        help="Query da cercare (default: società edilizie)",
    )
    parser.add_argument(
        "--output",
        default="risultati_maps.csv",
        help="File di output (.csv o .xlsx)",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Esegue il browser in modalità headless",
    )
    return parser.parse_args()


def is_sponsored(article_text: str) -> bool:
    """Heuristica semplice per filtrare risultati sponsorizzati."""
    text = article_text.lower()
    sponsored_keywords = ["sponsored", "annuncio", "ads", "ad ·", "sponsorizzato"]
    return any(keyword in text for keyword in sponsored_keywords)


def scroll_results_panel(page) -> None:
    """Scorre il pannello risultati finché non sembra arrivato in fondo."""
    feed = page.locator('div[role="feed"]').first
    feed.wait_for(timeout=15000)

    unchanged_rounds = 0
    last_count = 0

    for _ in range(60):
        cards = page.locator('div[role="article"]')
        current_count = cards.count()

        if current_count == last_count:
            unchanged_rounds += 1
        else:
            unchanged_rounds = 0
            last_count = current_count

        if unchanged_rounds >= 6:
            break

        page.evaluate(
            """(el) => {
                el.scrollBy(0, el.clientHeight * 0.85);
            }""",
            feed.element_handle(),
        )
        random_sleep(0.8, 1.8)


def extract_details_for_card(page, index: int) -> Dict[str, str]:
    """Apre una card e prova ad estrarre nome, indirizzo, telefono."""
    cards = page.locator('div[role="article"]')
    card = cards.nth(index)

    try:
        preview_text = card.inner_text(timeout=3000)
    except PlaywrightTimeoutError:
        return {}

    if is_sponsored(preview_text):
        return {}

    try:
        card.click(timeout=5000)
        random_sleep(1.2, 2.6)
    except PlaywrightTimeoutError:
        return {}

    # Nome
    name = ""
    try:
        name = page.locator("h1.DUwDvf").first.inner_text(timeout=5000).strip()
    except PlaywrightTimeoutError:
        return {}

    # Indirizzo
    address = ""
    try:
        address_btn = page.locator('button[data-item-id^="address"]').first
        address = address_btn.inner_text(timeout=3000).replace("Indirizzo: ", "").strip()
    except PlaywrightTimeoutError:
        address = ""

    # Telefono
    phone = ""
    try:
        phone_btn = page.locator('button[data-item-id^="phone:tel"]').first
        phone = phone_btn.inner_text(timeout=3000).replace("Telefono: ", "").strip()
    except PlaywrightTimeoutError:
        phone = ""

    if not name:
        return {}

    return {
        "Nome": name,
        "Indirizzo": address,
        "Telefono": phone,
    }


def save_results(rows: List[Dict[str, str]], output_path: str) -> None:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if output.suffix.lower() == ".xlsx":
        try:
            from openpyxl import Workbook

            wb = Workbook()
            ws = wb.active
            ws.title = "Risultati"
            ws.append(["Nome", "Indirizzo", "Telefono"])
            for row in rows:
                ws.append([row.get("Nome", ""), row.get("Indirizzo", ""), row.get("Telefono", "")])
            wb.save(output)
            return
        except Exception as exc:
            fallback_csv = output.with_suffix(".csv")
            print(
                f"[WARN] Salvataggio XLSX non riuscito ({exc}). Fallback su CSV: {fallback_csv}"
            )
            output = fallback_csv

    with output.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["Nome", "Indirizzo", "Telefono"])
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    args = parse_args()
    search_term = f"{args.query} {args.location}".strip()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless)
        context = browser.new_context(locale="it-IT")
        page = context.new_page()

        try:
            page.goto("https://www.google.com/maps", wait_until="domcontentloaded", timeout=60000)
            random_sleep(1.0, 2.5)

            # Accetta cookie se compare il pop-up
            cookie_btn = page.locator('button:has-text("Accetta tutto"), button:has-text("Accept all")').first
            if cookie_btn.is_visible(timeout=3000):
                cookie_btn.click(timeout=3000)
                random_sleep(1.0, 2.0)
        except Exception:
            # Proseguiamo comunque, il popup potrebbe non comparire.
            pass

        try:
            search_box = page.locator('input#searchboxinput').first
            search_box.wait_for(timeout=15000)
            search_box.fill(search_term)
            random_sleep(0.5, 1.1)
            page.keyboard.press("Enter")
            random_sleep(2.0, 3.5)
        except PlaywrightTimeoutError:
            browser.close()
            raise RuntimeError("Casella di ricerca non trovata su Google Maps.")

        try:
            scroll_results_panel(page)
        except PlaywrightTimeoutError:
            print("[WARN] Pannello risultati non trovato o non caricabile completamente.")

        cards = page.locator('div[role="article"]')
        total = cards.count()
        print(f"Card rilevate: {total}")

        results: List[Dict[str, str]] = []
        seen = set()

        for i in range(total):
            try:
                item = extract_details_for_card(page, i)
                if not item:
                    continue

                key = (item["Nome"], item["Indirizzo"], item["Telefono"])
                if key in seen:
                    continue
                seen.add(key)
                results.append(item)

                # Torna alla lista risultati (se necessario)
                back_btn = page.locator('button[aria-label="Indietro"], button[aria-label="Back"]').first
                if back_btn.is_visible(timeout=1500):
                    back_btn.click(timeout=3000)
                    random_sleep(0.9, 1.8)
            except Exception as exc:
                print(f"[WARN] Errore sulla card {i}: {exc}")
                continue

        save_results(results, args.output)
        print(f"Estratti {len(results)} record. File salvato in: {args.output}")

        context.close()
        browser.close()


if __name__ == "__main__":
    main()
