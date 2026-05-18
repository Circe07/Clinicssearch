#!/usr/bin/env python3
"""
Buscador avanzado de emails para clínicas.

Lee el Excel generado por buscar_clinicas.py, identifica las clínicas
sin email, y usa múltiples estrategias para encontrar sus correos:

  1. Rastreo profundo de toda la web (todas las páginas internas)
  2. Extracción de mailto: links y datos estructurados (JSON-LD / schema.org)
  3. Búsqueda en Google: "nombre clínica" + "email" / "@"
  4. Búsqueda en directorios españoles (Doctoralia, Páginas Amarillas, etc.)
  5. Adivinación de patrones (info@, contacto@, clinica@) + verificación SMTP/MX
  6. Redes sociales (Facebook about page)

Uso:
    python buscar_emails.py
    python buscar_emails.py --input clinicas_barcelona.xlsx
    python buscar_emails.py --workers 8
"""

import argparse
import json
import re
import smtplib
import socket
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse

import dns.resolver
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

HEADERS_BROWSER = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")

JUNK_EXTENSIONS = (
    ".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp",
    ".css", ".js", ".woff", ".woff2", ".ttf", ".ico",
    ".pdf", ".zip", ".mp4", ".mp3",
)
JUNK_DOMAINS = (
    "sentry.io", "w3.org", "schema.org", "example.com",
    "wordpress.org", "gravatar.com", "googleapis.com",
    "googletagmanager.com", "facebook.com", "twitter.com",
    "instagram.com", "youtube.com", "linkedin.com",
    "cookiebot.com", "cookielaw.com", "onetrust.com",
    "cloudflare.com", "gstatic.com", "google.com",
    "wixpress.com", "squarespace.com", "sentry-next.wixpress.com",
)

CONTACT_KEYWORDS = [
    "contacto", "contact", "contacta", "contacte", "kontakt",
    "sobre-nosotros", "about", "quienes-somos", "quien-somos",
    "aviso-legal", "legal", "aviso_legal", "legal-notice",
    "politica-de-privacidad", "privacidad", "privacy",
    "equipo", "team", "staff", "profesionales",
    "informacion", "info",
    "cita", "appointment", "pedir-cita", "reservar", "reserva",
    "donde-estamos", "ubicacion", "como-llegar", "localizacion",
    "impressum", "imprint",
    "pie-de-pagina", "footer",
]

EMAIL_PREFIXES = [
    "info", "contacto", "contact", "clinica", "recepcion",
    "administracion", "admin", "hola", "consulta", "consultas",
    "citas", "atencion", "secretaria", "gerencia", "direccion",
]

DIRECTORY_SEARCH_TEMPLATES = [
    "site:doctoralia.es {name}",
    "site:topdoctors.es {name}",
    "site:paginasamarillas.es {name}",
    "site:yelp.es {name}",
    '"{name}" "{city}" email OR correo OR "@"',
    '"{name}" "{city}" arroba',
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def parse_args():
    parser = argparse.ArgumentParser(
        description="Busca emails de clínicas con múltiples estrategias."
    )
    parser.add_argument(
        "--input", "-i",
        default="clinicas_prospecto.xlsx",
        help="Archivo Excel (default: clinicas_prospecto.xlsx).",
    )
    parser.add_argument(
        "--workers", "-w",
        type=int,
        default=5,
        help="Hilos paralelos (default: 5).",
    )
    return parser.parse_args()


def clean_emails(raw_emails: list[str]) -> list[str]:
    """Filtra emails basura y devuelve lista única."""
    seen = set()
    cleaned = []
    for email in raw_emails:
        lower = email.lower().strip()
        if lower in seen:
            continue
        if lower.endswith(JUNK_EXTENSIONS):
            continue
        if any(domain in lower for domain in JUNK_DOMAINS):
            continue
        if len(lower) > 60:
            continue
        if lower.startswith("noreply") or lower.startswith("no-reply"):
            continue
        if lower.count("@") != 1:
            continue
        seen.add(lower)
        cleaned.append(email)
    return cleaned


def fetch_page(url: str, timeout: int = 10) -> str | None:
    """Descarga HTML de una URL."""
    try:
        resp = requests.get(
            url, headers=HEADERS_BROWSER,
            timeout=timeout, allow_redirects=True, verify=True,
        )
        resp.raise_for_status()
        content_type = resp.headers.get("content-type", "")
        if "text/html" not in content_type and "text/plain" not in content_type:
            return None
        return resp.text
    except requests.exceptions.SSLError:
        try:
            resp = requests.get(
                url, headers=HEADERS_BROWSER,
                timeout=timeout, allow_redirects=True, verify=False,
            )
            resp.raise_for_status()
            return resp.text
        except Exception:
            return None
    except Exception:
        return None


def google_search(query: str, num_results: int = 5) -> list[str]:
    """Busca en Google y devuelve las URLs de los resultados."""
    url = "https://www.google.com/search"
    params = {"q": query, "num": num_results, "hl": "es"}
    headers = {**HEADERS_BROWSER, "Referer": "https://www.google.com/"}
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=10)
        if resp.status_code != 200:
            return []
        soup = BeautifulSoup(resp.text, "html.parser")
        urls = []
        for a_tag in soup.find_all("a", href=True):
            href = a_tag["href"]
            if href.startswith("/url?q="):
                real_url = href.split("/url?q=")[1].split("&")[0]
                if real_url.startswith("http") and "google." not in real_url:
                    urls.append(real_url)
        return urls[:num_results]
    except Exception:
        return []


# ---------------------------------------------------------------------------
# Strategy 1: Deep website crawl
# ---------------------------------------------------------------------------


def crawl_website(base_url: str, max_pages: int = 25) -> tuple[list[str], list[str]]:
    """
    Rastrea todo el sitio web buscando emails.
    Prioriza páginas de contacto, luego el resto.
    Retorna (emails, pages_visited).
    """
    if not base_url:
        return [], []

    parsed_base = urlparse(base_url)
    base_domain = parsed_base.netloc
    visited = set()
    all_emails = []
    pages_visited = []
    priority_queue = []
    normal_queue = []

    main_html = fetch_page(base_url)
    if not main_html:
        return [], [base_url]

    visited.add(base_url)
    pages_visited.append(base_url)
    all_emails.extend(extract_emails_from_html(main_html))

    soup = BeautifulSoup(main_html, "html.parser")
    for a_tag in soup.find_all("a", href=True):
        href = a_tag["href"].strip()
        if not href or href.startswith(("#", "javascript:", "tel:")):
            continue

        if href.startswith("mailto:"):
            email = href.replace("mailto:", "").split("?")[0].strip()
            if EMAIL_REGEX.match(email):
                all_emails.append(email)
            continue

        full_url = urljoin(base_url, href)
        parsed = urlparse(full_url)
        if parsed.netloc and parsed.netloc != base_domain:
            continue
        if any(full_url.lower().endswith(ext) for ext in JUNK_EXTENSIONS):
            continue

        clean_url = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"
        if clean_url in visited:
            continue

        path_lower = parsed.path.lower()
        link_text = a_tag.get_text(strip=True).lower()
        is_priority = any(kw in path_lower or kw in link_text for kw in CONTACT_KEYWORDS)

        if is_priority:
            priority_queue.append(clean_url)
        else:
            normal_queue.append(clean_url)
        visited.add(clean_url)

    for url in priority_queue + normal_queue:
        if len(pages_visited) >= max_pages:
            break
        html = fetch_page(url, timeout=8)
        if not html:
            continue
        pages_visited.append(url)
        all_emails.extend(extract_emails_from_html(html))

    return clean_emails(all_emails), pages_visited


def extract_emails_from_html(html: str) -> list[str]:
    """Extrae emails del HTML: regex + mailto + JSON-LD + schema.org."""
    emails = []

    emails.extend(EMAIL_REGEX.findall(html))

    soup = BeautifulSoup(html, "html.parser")

    for a_tag in soup.find_all("a", href=True):
        href = a_tag["href"]
        if href.startswith("mailto:"):
            email = href.replace("mailto:", "").split("?")[0].strip()
            if EMAIL_REGEX.match(email):
                emails.append(email)

    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "")
            emails.extend(extract_emails_from_jsonld(data))
        except (json.JSONDecodeError, TypeError):
            pass

    for tag in soup.find_all(attrs={"itemprop": "email"}):
        content = tag.get("content", "") or tag.get_text(strip=True)
        if EMAIL_REGEX.match(content):
            emails.append(content)

    return emails


def extract_emails_from_jsonld(data) -> list[str]:
    """Recorre recursivamente JSON-LD buscando campos de email."""
    emails = []
    if isinstance(data, dict):
        for key, value in data.items():
            if key.lower() in ("email", "contactpoint", "contactpoints"):
                if isinstance(value, str) and EMAIL_REGEX.match(value):
                    emails.append(value)
                elif isinstance(value, (list, dict)):
                    emails.extend(extract_emails_from_jsonld(value))
            else:
                emails.extend(extract_emails_from_jsonld(value))
    elif isinstance(data, list):
        for item in data:
            emails.extend(extract_emails_from_jsonld(item))
    return emails


# ---------------------------------------------------------------------------
# Strategy 2: Google Search
# ---------------------------------------------------------------------------


def search_email_google(name: str, city: str) -> list[str]:
    """Busca el email de una clínica en Google."""
    emails = []

    queries = [
        f'"{name}" email OR correo OR "@"',
        f'"{name}" "{city}" email',
    ]

    for query in queries:
        urls = google_search(query, num_results=5)
        for url in urls[:3]:
            html = fetch_page(url, timeout=8)
            if not html:
                continue
            found = EMAIL_REGEX.findall(html)
            emails.extend(found)
        if clean_emails(emails):
            break

    return clean_emails(emails)


# ---------------------------------------------------------------------------
# Strategy 3: Spanish directories
# ---------------------------------------------------------------------------


def search_directories(name: str, city: str) -> list[str]:
    """Busca emails en directorios españoles."""
    emails = []

    direct_urls = []

    slug = name.lower().replace(" ", "-").replace(".", "")
    slug = re.sub(r"[^a-z0-9\-]", "", slug)

    doctoralia_queries = [f"site:doctoralia.es \"{name}\""]
    for q in doctoralia_queries:
        urls = google_search(q, num_results=3)
        direct_urls.extend(urls)

    for url in direct_urls[:5]:
        html = fetch_page(url, timeout=8)
        if not html:
            continue
        found = EMAIL_REGEX.findall(html)
        emails.extend(found)

    return clean_emails(emails)


# ---------------------------------------------------------------------------
# Strategy 4: Email pattern guessing + MX/SMTP verify
# ---------------------------------------------------------------------------


def get_domain_from_web(web_url: str) -> str | None:
    """Extrae el dominio raíz de una URL de web."""
    if not web_url:
        return None
    parsed = urlparse(web_url)
    domain = parsed.netloc
    if domain.startswith("www."):
        domain = domain[4:]
    return domain if domain else None


def check_mx_exists(domain: str) -> bool:
    """Verifica si un dominio tiene registros MX (acepta correo)."""
    try:
        dns.resolver.resolve(domain, "MX")
        return True
    except Exception:
        return False


def verify_email_smtp(email: str) -> bool | None:
    """
    Verifica si un email existe via SMTP (sin enviar correo).
    Retorna True (existe), False (no existe), None (no se pudo verificar).
    """
    domain = email.split("@")[1]
    try:
        records = dns.resolver.resolve(domain, "MX")
        mx_host = str(records[0].exchange).rstrip(".")
    except Exception:
        return None

    try:
        smtp = smtplib.SMTP(timeout=5)
        smtp.connect(mx_host, 25)
        smtp.helo("verificador.local")
        smtp.mail("test@verificador.local")
        code, _ = smtp.rcpt(email)
        smtp.quit()
        return code == 250
    except Exception:
        return None


def guess_and_verify_emails(web_url: str) -> list[str]:
    """Genera emails probables y verifica cuáles existen."""
    domain = get_domain_from_web(web_url)
    if not domain:
        return []

    if not check_mx_exists(domain):
        return []

    verified = []
    for prefix in EMAIL_PREFIXES:
        candidate = f"{prefix}@{domain}"
        result = verify_email_smtp(candidate)
        if result is True:
            verified.append(candidate)
            break
        elif result is None:
            verified.append(candidate)
            break

    return verified


# ---------------------------------------------------------------------------
# Strategy 5: Social media (Facebook)
# ---------------------------------------------------------------------------


def find_facebook_email(html: str, base_url: str) -> list[str]:
    """Busca la página de Facebook y extrae el email si está visible."""
    soup = BeautifulSoup(html, "html.parser")
    fb_urls = []

    for a_tag in soup.find_all("a", href=True):
        href = a_tag["href"]
        if "facebook.com" in href and "/posts/" not in href and "/photos/" not in href:
            fb_urls.append(href)

    emails = []
    for fb_url in fb_urls[:2]:
        html = fetch_page(fb_url, timeout=8)
        if not html:
            continue
        found = EMAIL_REGEX.findall(html)
        emails.extend(found)

    return clean_emails(emails)


# ---------------------------------------------------------------------------
# Orchestrator: run all strategies for one clinic
# ---------------------------------------------------------------------------


def find_email_for_clinic(name: str, web: str, address: str) -> dict:
    """
    Ejecuta todas las estrategias para encontrar el email de una clínica.
    Retorna dict con resultado y detalles de cada estrategia.
    """
    result = {
        "name": name,
        "emails": [],
        "strategy_used": None,
        "pages_crawled": 0,
        "strategies_tried": [],
    }

    city = ""
    if address:
        parts = [p.strip() for p in address.split(",")]
        for part in parts:
            cleaned = re.sub(r"\d{5}", "", part).strip()
            if cleaned and len(cleaned) > 2 and not cleaned.isdigit():
                city = cleaned
                break

    # --- Strategy 1: Deep crawl ---
    result["strategies_tried"].append("crawl")
    if web:
        emails, pages = crawl_website(web, max_pages=25)
        result["pages_crawled"] = len(pages)
        if emails:
            result["emails"] = emails
            result["strategy_used"] = "Rastreo web profundo"
            return result

    # --- Strategy 2: Facebook from website ---
    if web:
        result["strategies_tried"].append("facebook")
        main_html = fetch_page(web)
        if main_html:
            fb_emails = find_facebook_email(main_html, web)
            if fb_emails:
                result["emails"] = fb_emails
                result["strategy_used"] = "Facebook"
                return result

    # --- Strategy 3: Email pattern guessing + SMTP ---
    if web:
        result["strategies_tried"].append("smtp_guess")
        guessed = guess_and_verify_emails(web)
        if guessed:
            result["emails"] = guessed
            result["strategy_used"] = "Patrón email verificado (SMTP)"
            return result

    # --- Strategy 4: Google search ---
    result["strategies_tried"].append("google")
    google_emails = search_email_google(name, city)
    if google_emails:
        result["emails"] = google_emails
        result["strategy_used"] = "Búsqueda en Google"
        return result

    # --- Strategy 5: Directories ---
    result["strategies_tried"].append("directories")
    dir_emails = search_directories(name, city)
    if dir_emails:
        result["emails"] = dir_emails
        result["strategy_used"] = "Directorios (Doctoralia, etc.)"
        return result

    result["strategy_used"] = "No encontrado"
    return result


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    args = parse_args()

    print("=" * 60)
    print("  BUSCADOR AVANZADO DE EMAILS")
    print("  Sistema multi-estrategia")
    print("=" * 60)

    try:
        wb = load_workbook(args.input)
    except FileNotFoundError:
        print(f"\n  [ERROR] No se encontró '{args.input}'")
        print("  Ejecuta primero buscar_clinicas.py para generar el Excel.")
        sys.exit(1)
    except Exception as e:
        print(f"\n  [ERROR] No se pudo abrir '{args.input}': {e}")
        sys.exit(1)

    ws = wb["Clínicas - Prospectos"]

    clinics_to_search = []
    total_clinics = 0
    already_have_email = 0

    for row_idx in range(2, ws.max_row + 1):
        name = ws.cell(row=row_idx, column=1).value
        if not name:
            continue
        total_clinics += 1
        email = ws.cell(row=row_idx, column=3).value
        address = ws.cell(row=row_idx, column=4).value or ""

        web_cell = ws.cell(row=row_idx, column=5)
        web = ""
        if web_cell.hyperlink:
            web = web_cell.hyperlink.target
        elif web_cell.value and web_cell.value != "Ver web":
            web = web_cell.value

        if email and str(email).strip():
            already_have_email += 1
        else:
            clinics_to_search.append({
                "row": row_idx,
                "name": str(name),
                "web": web,
                "address": str(address),
            })

    print(f"\n  Total clínicas en Excel:    {total_clinics}")
    print(f"  Ya tienen email:            {already_have_email}")
    print(f"  Sin email (a buscar):       {len(clinics_to_search)}")
    sin_web = sum(1 for c in clinics_to_search if not c["web"])
    con_web = len(clinics_to_search) - sin_web
    print(f"    - Con web (todas las estrategias): {con_web}")
    print(f"    - Sin web (solo Google/directorios): {sin_web}")

    if not clinics_to_search:
        print("\n  Todas las clínicas ya tienen email.")
        wb.close()
        sys.exit(0)

    print(f"\n  Estrategias disponibles:")
    print(f"    1. Rastreo web profundo (hasta 25 páginas por clínica)")
    print(f"    2. Extracción mailto: + datos estructurados (JSON-LD)")
    print(f"    3. Redes sociales (Facebook)")
    print(f"    4. Adivinación de patrones + verificación SMTP")
    print(f"    5. Búsqueda en Google")
    print(f"    6. Directorios (Doctoralia, etc.)")
    print(f"\n  Buscando emails ({args.workers} hilos)...\n")

    found_count = 0
    not_found_count = 0
    strategy_stats = {}
    results = {}

    with ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {
            executor.submit(
                find_email_for_clinic,
                c["name"], c["web"], c["address"]
            ): c
            for c in clinics_to_search
        }
        for i, future in enumerate(as_completed(futures), 1):
            clinic = futures[future]
            try:
                res = future.result()
                short_name = res["name"][:40]
                if res["emails"]:
                    found_count += 1
                    results[clinic["row"]] = res["emails"]
                    strategy = res["strategy_used"]
                    strategy_stats[strategy] = strategy_stats.get(strategy, 0) + 1
                    emails_str = ", ".join(res["emails"][:3])
                    print(f"  [{i}/{len(clinics_to_search)}] ENCONTRADO  {short_name}")
                    print(f"       Emails: {emails_str}")
                    print(f"       Via:    {strategy}")
                else:
                    not_found_count += 1
                    tried = ", ".join(res["strategies_tried"])
                    print(f"  [{i}/{len(clinics_to_search)}] sin resultado  {short_name}")
                    print(f"       Estrategias usadas: {tried} ({res['pages_crawled']} págs)")
            except Exception as e:
                not_found_count += 1
                print(f"  [{i}/{len(clinics_to_search)}] ERROR  {clinic['name'][:40]}: {e}")

    if found_count > 0:
        print(f"\n  Actualizando Excel con {found_count} emails encontrados...")
        for row_idx, emails in results.items():
            ws.cell(row=row_idx, column=3, value=", ".join(emails))
        wb.save(args.input)
        print(f"  Excel guardado: {args.input}")
    else:
        print(f"\n  No se encontraron emails nuevos.")

    print("\n" + "=" * 60)
    print("  RESUMEN")
    print("=" * 60)
    print(f"  Clínicas buscadas:      {len(clinics_to_search)}")
    print(f"  Emails encontrados:     {found_count}")
    print(f"  Sin resultado:          {not_found_count}")
    if strategy_stats:
        print(f"\n  Emails por estrategia:")
        for strategy, count in sorted(strategy_stats.items(), key=lambda x: -x[1]):
            print(f"    {strategy}: {count}")
    print("=" * 60)

    wb.close()


if __name__ == "__main__":
    main()
