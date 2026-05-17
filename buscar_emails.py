#!/usr/bin/env python3
"""
Buscador de emails para clínicas.

Lee el Excel generado por buscar_clinicas.py, identifica las clínicas
sin email, y busca el email rastreando su web a fondo (subpáginas de
contacto, aviso legal, política de privacidad, etc.).

Uso:
    python buscar_emails.py
    python buscar_emails.py --input clinicas_barcelona.xlsx
    python buscar_emails.py --workers 10
"""

import argparse
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font

HEADERS_BROWSER = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
}

CONTACT_PAGE_PATTERNS = [
    "contacto", "contact", "contacta", "contacte",
    "sobre-nosotros", "about", "quienes-somos", "quien-somos",
    "aviso-legal", "legal", "aviso_legal",
    "politica-de-privacidad", "privacidad", "privacy",
    "equipo", "team", "staff",
    "informacion", "info",
    "cita", "appointment", "pedir-cita", "reservar",
    "donde-estamos", "ubicacion", "como-llegar",
    "impressum", "imprint",
]

EMAIL_PATTERN = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")

JUNK_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".css", ".js", ".woff", ".woff2", ".ttf", ".ico")
JUNK_DOMAINS = ("sentry.io", "w3.org", "schema.org", "example.com", "wordpress.org", "gravatar.com", "googleapis.com")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Busca emails de clínicas que no los tienen en el Excel."
    )
    parser.add_argument(
        "--input", "-i",
        default="clinicas_prospecto.xlsx",
        help="Archivo Excel de entrada (default: clinicas_prospecto.xlsx).",
    )
    parser.add_argument(
        "--workers", "-w",
        type=int,
        default=5,
        help="Hilos paralelos (default: 5).",
    )
    return parser.parse_args()


def extract_emails(html: str) -> list[str]:
    """Extrae emails limpios del HTML."""
    raw = EMAIL_PATTERN.findall(html)
    seen = set()
    cleaned = []
    for email in raw:
        lower = email.lower()
        if lower in seen:
            continue
        if lower.endswith(JUNK_EXTENSIONS):
            continue
        if any(domain in lower for domain in JUNK_DOMAINS):
            continue
        seen.add(lower)
        cleaned.append(email)
    return cleaned


def fetch_page(url: str) -> str | None:
    """Descarga una página y devuelve el HTML, o None si falla."""
    try:
        resp = requests.get(url, headers=HEADERS_BROWSER, timeout=10, allow_redirects=True, verify=True)
        resp.raise_for_status()
        return resp.text
    except requests.exceptions.SSLError:
        try:
            resp = requests.get(url, headers=HEADERS_BROWSER, timeout=10, allow_redirects=True, verify=False)
            resp.raise_for_status()
            return resp.text
        except Exception:
            return None
    except Exception:
        return None


def find_contact_links(html: str, base_url: str) -> list[str]:
    """Encuentra enlaces a páginas de contacto, legal, etc."""
    soup = BeautifulSoup(html, "html.parser")
    parsed_base = urlparse(base_url)
    base_domain = parsed_base.netloc

    contact_urls = []
    seen = set()

    for a_tag in soup.find_all("a", href=True):
        href = a_tag["href"].strip()
        if not href or href.startswith(("#", "javascript:", "mailto:", "tel:")):
            continue

        full_url = urljoin(base_url, href)
        parsed = urlparse(full_url)

        if parsed.netloc and parsed.netloc != base_domain:
            continue

        path_lower = parsed.path.lower().rstrip("/")
        link_text = a_tag.get_text(strip=True).lower()

        is_contact = any(p in path_lower for p in CONTACT_PAGE_PATTERNS)
        is_contact = is_contact or any(p in link_text for p in CONTACT_PAGE_PATTERNS)

        if is_contact and full_url not in seen:
            seen.add(full_url)
            contact_urls.append(full_url)

    return contact_urls


def find_all_internal_links(html: str, base_url: str, max_links: int = 30) -> list[str]:
    """Encuentra todos los enlaces internos (para rastreo amplio)."""
    soup = BeautifulSoup(html, "html.parser")
    parsed_base = urlparse(base_url)
    base_domain = parsed_base.netloc

    urls = []
    seen = {base_url}

    for a_tag in soup.find_all("a", href=True):
        if len(urls) >= max_links:
            break

        href = a_tag["href"].strip()
        if not href or href.startswith(("#", "javascript:", "mailto:", "tel:")):
            continue

        full_url = urljoin(base_url, href)
        parsed = urlparse(full_url)

        if parsed.netloc and parsed.netloc != base_domain:
            continue

        if any(full_url.lower().endswith(ext) for ext in JUNK_EXTENSIONS):
            continue

        clean_url = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"
        if clean_url not in seen:
            seen.add(clean_url)
            urls.append(clean_url)

    return urls


def search_emails_for_clinic(name: str, web: str) -> tuple[str, list[str], list[str]]:
    """
    Busca emails para una clínica rastreando su web.
    Retorna (nombre, emails_encontrados, paginas_rastreadas).
    """
    if not web:
        return name, [], []

    all_emails = []
    pages_visited = []

    main_html = fetch_page(web)
    if not main_html:
        return name, [], [web]

    pages_visited.append(web)
    emails = extract_emails(main_html)
    all_emails.extend(emails)

    if all_emails:
        return name, all_emails, pages_visited

    contact_links = find_contact_links(main_html, web)

    for url in contact_links:
        if url in pages_visited:
            continue
        html = fetch_page(url)
        if not html:
            continue
        pages_visited.append(url)
        emails = extract_emails(html)
        all_emails.extend(emails)
        if all_emails:
            return name, list(dict.fromkeys(all_emails)), pages_visited

    internal_links = find_all_internal_links(main_html, web, max_links=20)
    for url in internal_links:
        if url in pages_visited:
            continue
        html = fetch_page(url)
        if not html:
            continue
        pages_visited.append(url)
        emails = extract_emails(html)
        all_emails.extend(emails)
        if all_emails:
            break

    unique_emails = list(dict.fromkeys(all_emails))
    return name, unique_emails, pages_visited


def main():
    args = parse_args()

    print("=" * 60)
    print("  BUSCADOR DE EMAILS - Completar datos de clínicas")
    print("=" * 60)

    try:
        wb = load_workbook(args.input)
    except FileNotFoundError:
        print(f"\n  [ERROR] No se encontró el archivo '{args.input}'")
        print("  Ejecuta primero buscar_clinicas.py para generar el Excel.")
        sys.exit(1)
    except Exception as e:
        print(f"\n  [ERROR] No se pudo abrir '{args.input}': {e}")
        sys.exit(1)

    ws = wb["Clínicas - Prospectos"]

    clinics_without_email = []
    total_clinics = 0

    for row_idx in range(2, ws.max_row + 1):
        name = ws.cell(row=row_idx, column=1).value
        if not name:
            continue
        total_clinics += 1
        email = ws.cell(row=row_idx, column=3).value
        web_cell = ws.cell(row=row_idx, column=5)
        web = web_cell.hyperlink.target if web_cell.hyperlink else (web_cell.value or "")
        if web == "Ver web":
            web = web_cell.hyperlink.target if web_cell.hyperlink else ""

        if not email or not str(email).strip():
            clinics_without_email.append({
                "row": row_idx,
                "name": name,
                "web": web,
            })

    print(f"\n  Total clínicas en Excel:  {total_clinics}")
    print(f"  Sin email:                {len(clinics_without_email)}")
    sin_web = sum(1 for c in clinics_without_email if not c["web"])
    print(f"  Sin email NI web:         {sin_web} (no se puede buscar)")
    buscables = len(clinics_without_email) - sin_web
    print(f"  Buscables (tienen web):   {buscables}")

    if buscables == 0:
        print("\n  No hay clínicas con web a las que buscar email.")
        wb.close()
        sys.exit(0)

    print(f"\n  Rastreando webs para encontrar emails ({args.workers} hilos)...")
    print()

    found_count = 0
    not_found_count = 0
    results = {}

    searchable = [c for c in clinics_without_email if c["web"]]

    with ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {
            executor.submit(search_emails_for_clinic, c["name"], c["web"]): c
            for c in searchable
        }
        for i, future in enumerate(as_completed(futures), 1):
            clinic = futures[future]
            try:
                name, emails, pages = future.result()
                results[clinic["row"]] = emails
                if emails:
                    found_count += 1
                    print(f"  [{i}/{buscables}] ENCONTRADO  {name[:40]}")
                    for em in emails:
                        print(f"               → {em}")
                else:
                    not_found_count += 1
                    print(f"  [{i}/{buscables}] sin resultado  {name[:40]}  ({len(pages)} págs rastreadas)")
            except Exception as e:
                not_found_count += 1
                print(f"  [{i}/{buscables}] ERROR  {clinic['name'][:40]}: {e}")

    if found_count > 0:
        print(f"\n  Actualizando Excel...")
        for row_idx, emails in results.items():
            if emails:
                ws.cell(row=row_idx, column=3, value=", ".join(emails))

        wb.save(args.input)
        print(f"  Excel guardado: {args.input}")

    print("\n" + "=" * 60)
    print("  RESUMEN")
    print("=" * 60)
    print(f"  Clínicas buscadas:    {buscables}")
    print(f"  Emails encontrados:   {found_count}")
    print(f"  Sin resultado:        {not_found_count}")
    if sin_web > 0:
        print(f"  Sin web (no buscable):{sin_web}")
    print("=" * 60)

    wb.close()


if __name__ == "__main__":
    main()
