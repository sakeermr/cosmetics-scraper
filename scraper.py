#!/usr/bin/env python3
"""
Cosmetics.lk product scraper
----------------------------
Opens the "Women" collection on cosmetics.lk, visits each product page one by one,
and extracts: image, product name, price, and description (plus a few useful extras).
Results are saved to an Excel file with the product image embedded in each row.

Usage:
    python scraper.py                       # 10 products (demo)
    python scraper.py --max-products 50     # more products
    python scraper.py --collection men      # a different collection

cosmetics.lk runs on Shopify, so every product page also has a clean JSON version
(/products/<handle>.json). The scraper uses that first because it is far more reliable
than parsing the page layout, and falls back to reading the HTML page if JSON fails.
"""

import argparse
import html
import io
import logging
import os
import re
import sys
import time
from datetime import datetime
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

BASE_URL = "https://cosmetics.lk"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}
REQUEST_DELAY = 1.0      # seconds between requests (be polite to the site)
MAX_RETRIES = 3
TIMEOUT = 30
THUMB_SIZE = 120         # pixel size of the image embedded in Excel

log = logging.getLogger("scraper")


# --------------------------------------------------------------------------- #
# HTTP helpers
# --------------------------------------------------------------------------- #
session = requests.Session()
session.headers.update(HEADERS)


def fetch(url, expect_json=False):
    """GET a URL with retries. Returns parsed JSON, text, or None on failure."""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.get(url, timeout=TIMEOUT)
            if resp.status_code == 429:  # rate limited -> back off
                wait = 5 * attempt
                log.warning("Rate limited on %s, waiting %ss", url, wait)
                time.sleep(wait)
                continue
            resp.raise_for_status()
            time.sleep(REQUEST_DELAY)
            return resp.json() if expect_json else resp.text
        except (requests.RequestException, ValueError) as exc:
            log.warning("Attempt %d/%d failed for %s: %s", attempt, MAX_RETRIES, url, exc)
            time.sleep(2 * attempt)
    return None


def fetch_bytes(url):
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.get(url, timeout=TIMEOUT)
            resp.raise_for_status()
            return resp.content
        except requests.RequestException as exc:
            log.warning("Image attempt %d/%d failed: %s", attempt, MAX_RETRIES, exc)
            time.sleep(2 * attempt)
    return None


def normalise_url(src):
    if not src:
        return ""
    if src.startswith("//"):
        return "https:" + src
    return urljoin(BASE_URL, src)


# --------------------------------------------------------------------------- #
# Step 1: open the collection page ("click Women") and list product links
# --------------------------------------------------------------------------- #
def get_product_handles(collection, max_products):
    """Return product handles in the order shown on the collection page."""
    handles = []
    page = 1
    while len(handles) < max_products:
        url = f"{BASE_URL}/collections/{collection}?sort_by=best-selling&page={page}"
        log.info("Opening collection page: %s", url)
        page_html = fetch(url)
        if not page_html:
            break
        soup = BeautifulSoup(page_html, "html.parser")
        found_on_page = 0
        for a in soup.select('a[href*="/products/"]'):
            path = urlparse(a.get("href", "")).path
            m = re.search(r"/products/([^/?#]+)", path)
            if not m:
                continue
            handle = m.group(1)
            if handle not in handles:
                handles.append(handle)
                found_on_page += 1
            if len(handles) >= max_products:
                break
        if found_on_page == 0:  # no more products / last page
            break
        page += 1

    # Fallback: Shopify's collection JSON feed
    if not handles:
        log.info("HTML listing gave no products, trying collection JSON feed")
        data = fetch(
            f"{BASE_URL}/collections/{collection}/products.json?limit={min(max_products, 250)}",
            expect_json=True,
        )
        if data and data.get("products"):
            handles = [p["handle"] for p in data["products"]]

    return handles[:max_products]


# --------------------------------------------------------------------------- #
# Step 2: open each product ("click one by one") and extract details
# --------------------------------------------------------------------------- #
def html_to_text(raw_html):
    """Turn product description HTML into readable plain text with line breaks."""
    if not raw_html:
        return ""
    soup = BeautifulSoup(raw_html, "html.parser")
    for br in soup.find_all("br"):
        br.replace_with("\n")
    blocks = []
    for el in soup.find_all(["p", "li", "h1", "h2", "h3", "h4", "h5", "h6", "div"]):
        if el.find(["p", "li", "div"]):  # skip wrappers, keep leaf blocks
            continue
        text = el.get_text(" ", strip=True)
        if text:
            blocks.append(("• " + text) if el.name == "li" else text)
    text = "\n".join(blocks) if blocks else soup.get_text("\n", strip=True)
    text = html.unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def format_price(value):
    if value in (None, "", "0", "0.00"):
        return None
    try:
        return round(float(str(value).replace(",", "")), 2)
    except ValueError:
        return None


def product_from_json(handle):
    data = fetch(f"{BASE_URL}/products/{handle}.json", expect_json=True)
    if not data or "product" not in data:
        return None
    p = data["product"]
    variant = (p.get("variants") or [{}])[0]
    images = [normalise_url(img.get("src")) for img in p.get("images") or []]
    return {
        "name": p.get("title", "").strip(),
        "brand": p.get("vendor", ""),
        "type": p.get("product_type", ""),
        "price": format_price(variant.get("price")),
        "original_price": format_price(variant.get("compare_at_price")),
        "sku": variant.get("sku", ""),
        "description": html_to_text(p.get("body_html", "")),
        "image_url": images[0] if images else "",
        "all_images": " | ".join(images),
        "url": f"{BASE_URL}/products/{handle}",
    }


def product_from_html(handle):
    """Fallback: read the visible product page."""
    url = f"{BASE_URL}/products/{handle}"
    page_html = fetch(url)
    if not page_html:
        return None
    soup = BeautifulSoup(page_html, "html.parser")

    def meta(prop):
        tag = soup.find("meta", attrs={"property": prop}) or soup.find("meta", attrs={"name": prop})
        return tag.get("content", "").strip() if tag else ""

    name = (soup.select_one("h1") or {}).get_text(strip=True) if soup.select_one("h1") else meta("og:title")
    price_text = ""
    price_el = soup.select_one(".price-item--sale, .price-item--regular, .price, [class*=price]")
    if price_el:
        price_text = price_el.get_text(" ", strip=True)
    m = re.search(r"([\d,]+\.\d{2})", price_text) or re.search(r"([\d,]+\.\d{2})", meta("og:price:amount") or "")
    desc_el = soup.select_one(".product__description, .product-description, [class*=description]")
    return {
        "name": name,
        "brand": "",
        "type": "",
        "price": format_price(m.group(1)) if m else format_price(meta("og:price:amount")),
        "original_price": None,
        "sku": "",
        "description": html_to_text(str(desc_el)) if desc_el else meta("og:description"),
        "image_url": normalise_url(meta("og:image")),
        "all_images": normalise_url(meta("og:image")),
        "url": url,
    }


def scrape_product(handle):
    log.info("Opening product: %s", handle)
    product = product_from_json(handle)
    if not product or not product.get("name"):
        log.info("  JSON unavailable, reading HTML page instead")
        product = product_from_html(handle)
    return product


# --------------------------------------------------------------------------- #
# Step 3: download images and build the Excel file
# --------------------------------------------------------------------------- #
def make_thumbnail(image_bytes, out_path):
    img = PILImage.open(io.BytesIO(image_bytes))
    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGBA")
    img.thumbnail((THUMB_SIZE, THUMB_SIZE))
    canvas = PILImage.new("RGB", (THUMB_SIZE, THUMB_SIZE), "white")
    offset = ((THUMB_SIZE - img.width) // 2, (THUMB_SIZE - img.height) // 2)
    canvas.paste(img, offset, img if img.mode == "RGBA" else None)
    canvas.save(out_path, "PNG")
    return out_path


def build_excel(products, out_file, images_dir):
    os.makedirs(images_dir, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"

    headers = [
        ("#", 5), ("Image", 18), ("Product Name", 38), ("Brand", 16), ("Type", 14),
        ("Price (Rs)", 13), ("Original Price (Rs)", 15), ("On Sale", 9), ("SKU", 14),
        ("Description", 70), ("Product URL", 40), ("Image URL", 40),
    ]
    header_fill = PatternFill("solid", fgColor="1F3864")
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, (title, width) in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.font = Font(bold=True, color="FFFFFF", name="Arial")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "C2"

    for i, p in enumerate(products, start=1):
        row = i + 1
        on_sale = bool(p["original_price"] and p["price"] and p["original_price"] > p["price"])
        values = [
            i, None, p["name"], p["brand"], p["type"], p["price"],
            p["original_price"] if on_sale else None, "Yes" if on_sale else "No",
            p["sku"], p["description"], p["url"], p["image_url"],
        ]
        for col, val in enumerate(values, start=1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(vertical="top", wrap_text=col in (3, 10))
            cell.border = border
            if col in (6, 7):
                cell.number_format = "#,##0.00"
        for col in (11, 12):
            cell = ws.cell(row=row, column=col)
            if cell.value:
                cell.hyperlink = cell.value
                cell.font = Font(name="Arial", size=10, color="0563C1", underline="single")
        ws.row_dimensions[row].height = 95

        # Embed product image
        if p["image_url"]:
            img_bytes = fetch_bytes(p["image_url"])
            if img_bytes:
                try:
                    thumb = make_thumbnail(img_bytes, os.path.join(images_dir, f"product_{i:03d}.png"))
                    xl_img = XLImage(thumb)
                    xl_img.width = xl_img.height = THUMB_SIZE
                    ws.add_image(xl_img, f"B{row}")
                except Exception as exc:  # corrupt / unsupported image
                    log.warning("Could not embed image for %s: %s", p["name"], exc)

    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(products) + 1}"

    # Small info sheet
    info = wb.create_sheet("Run Info")
    info["A1"], info["B1"] = "Source", f"{BASE_URL}"
    info["A2"], info["B2"] = "Scraped at (UTC)", datetime.utcnow().strftime("%Y-%m-%d %H:%M")
    info["A3"], info["B3"] = "Products", len(products)
    info.column_dimensions["A"].width = 20
    info.column_dimensions["B"].width = 40
    for r in range(1, 4):
        info[f"A{r}"].font = Font(bold=True, name="Arial")

    wb.save(out_file)
    log.info("Saved %d products to %s", len(products), out_file)


# --------------------------------------------------------------------------- #
def main():
    parser = argparse.ArgumentParser(description="Scrape cosmetics.lk products to Excel")
    parser.add_argument("--collection", default=os.getenv("COLLECTION", "women"))
    parser.add_argument("--max-products", type=int, default=int(os.getenv("MAX_PRODUCTS", "10")))
    parser.add_argument("--output", default=os.getenv("OUTPUT_FILE", "output/cosmetics_products.xlsx"))
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    os.makedirs(os.path.dirname(args.output) or ".", exist_ok=True)

    handles = get_product_handles(args.collection, args.max_products)
    if not handles:
        log.error("No products found - the site may be blocking requests or the layout changed.")
        sys.exit(1)
    log.info("Found %d products to scrape", len(handles))

    products = []
    for h in handles:
        p = scrape_product(h)
        if p and p.get("name"):
            products.append(p)
        else:
            log.warning("Skipped %s (no data)", h)

    if not products:
        log.error("No product details could be extracted.")
        sys.exit(1)

    build_excel(products, args.output, os.path.join(os.path.dirname(args.output) or ".", "images"))


if __name__ == "__main__":
    main()
