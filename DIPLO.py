#!/usr/bin/env python3

# ============================================================
# Interactive Nitter Search + LLM Translator
# + Location Inference + Interactive Map
# ============================================================
#
# What this does:
# 1. Prompts for any search terms and optional query-translation languages.
# 2. Searches Nitter using the original and translated terms.
# 3. Parses visible Nitter HTML.
# 4. Translates posts using the selected LLM provider.
# 5. Infers broad public locations using the selected LLM provider.
# 6. Geocodes inferred locations with Nominatim / OpenStreetMap.
# 7. Exports clean CSV and formatted Excel files.
# 8. Creates an interactive HTML map with popups.
#
# This does NOT bypass login walls, CAPTCHAs, private accounts,
# anti-bot systems, access controls, or rate limits. 
#
# HOW TO RUN (macOS / Linux Terminal)
# -----------------------------------
# First-time setup, from your project directory:
#
#     python3 -m venv .venv
#     source .venv/bin/activate
#     python3 -m pip install --upgrade pip
#     python3 DIPLO.py
#
# The script installs missing Python packages into the active virtual
# environment. On later runs, from your project directory:
#
#     source .venv/bin/activate
#     python3 DIPLO.py
#
# At startup, choose OpenAI or Virginia Tech ARC, select a model, and enter
# that provider's API key using hidden input. Keys are kept only in memory for
# the current run and are not written to project files.
#
# To leave the virtual environment afterward:
#
#     deactivate
#
# ============================================================


# ============================================================
# USER SETTINGS
# ============================================================

# Use live mode to fetch from Nitter.
# Use saved_html mode to parse saved Nitter pages from your browser.
MODE = "live"
# MODE = "saved_html"

NITTER_BASE_URLS = [
    "https://nitter.net",
]

# For saved_html mode, save Nitter pages from your browser and list them here.
SAVED_HTML_FILES = [
    # "nitter_results.html",
]

LLM_MODEL = "gpt-5.6-luna"
LLM_PROVIDER = "openai"
LLM_BASE_URL = ""
LLM_API_KEY = ""
TARGET_LANGUAGE = "English"

# Start broad. Once results work, add dates.
SINCE_DATE = None
UNTIL_DATE = None

# Example:
# SINCE_DATE = "2026-01-01"
# UNTIL_DATE = "2026-06-29"

MAX_POSTS_PER_QUERY = 10
MAX_PAGES_PER_QUERY = 1

INCLUDE_RETWEETS = False
TRANSLATE_POSTS = True

REQUEST_DELAY_SECONDS = 3.0
TRANSLATION_DELAY_SECONDS = 0.5

OUTPUT_FILE = "nitter_search_posts.csv"
RUN_INSTANCE_DIAGNOSTIC_FIRST = True

# Leave empty to search keywords.
# Add handles for cleaner results from known accounts.
SEARCH_HANDLES = [
    # "@ConfuciusINST",
    # "@SomeUniversityCI",
]

DEFAULT_SEARCH_TERMS = [
    "Confucius Institute",
]

# The original terms are always searched. The user can add translations in any
# of these languages at runtime by entering the corresponding number(s).
SEARCH_LANGUAGE_OPTIONS = {
    "1": "English",
    "2": "Simplified Chinese",
    "3": "Traditional Chinese",
    "4": "Spanish",
    "5": "French",
    "6": "Portuguese",
    "7": "Italian",
    "8": "German",
    "9": "Russian",
    "10": "Arabic",
    "11": "Japanese",
    "12": "Korean",
    "13": "Hindi",
    "14": "Turkish",
    "15": "Indonesian",
}


# ============================================================
# MAP / LOCATION SETTINGS
# ============================================================

CREATE_MAP = True
INFER_LOCATIONS = True

MAP_FILE = "nitter_search_map.html"
GEOCODE_CACHE_FILE = "geocode_cache.json"

# Map tile settings.
# The old Folium default, tiles="OpenStreetMap", can request tiles from
# tile.openstreetmap.org and may produce browser errors such as:
# 503r: Access Blocked / osm.wiki/blocked.
# To avoid that, this script does NOT use the direct OpenStreetMap tile server.
# It uses CARTO's public basemap CDN instead, with proper attribution.
MAP_BASE_TILE_NAME = "CartoDB Positron"
MAP_BASE_TILE_URL = "https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png"
MAP_BASE_TILE_ATTRIBUTION = (
    '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> '
    'contributors &copy; <a href="https://carto.com/attributions">CARTO</a>'
)

# Optional fallback/alternate basemaps. These also avoid the direct OSM tile server.
ADD_ALTERNATE_BASEMAPS = True

# Inferred locations are broad, public, city/region/country-level locations.
# They are not verified geotags.
MIN_LOCATION_CONFIDENCE = 0.2

# Nominatim requires a real identifying User-Agent.
NOMINATIM_USER_AGENT = "interactive-nitter-search-map-research/1.0"

# Nominatim public service requires no more than 1 request per second.
GEOCODE_DELAY_SECONDS = 1.2


# ============================================================
# INSTALL PACKAGES
# ============================================================

import sys
import subprocess
import importlib.util
import os
import time
import csv
import re
import json
import html as html_lib
from getpass import getpass
from dataclasses import asdict, dataclass
from typing import List, Optional, Dict, Any
from urllib.parse import urlencode, urljoin, urlparse



def install_if_missing(import_name, pip_name=None):
    if pip_name is None:
        pip_name = import_name

    if importlib.util.find_spec(import_name) is None:
        print(f"Installing {pip_name}...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "-q", pip_name]
        )
    else:
        print(f"{pip_name} already installed.")


install_if_missing("openai")
install_if_missing("pandas")
install_if_missing("bs4", "beautifulsoup4")
install_if_missing("requests")
install_if_missing("certifi")
install_if_missing("langdetect")
install_if_missing("tqdm")
install_if_missing("dateutil", "python-dateutil")
install_if_missing("openpyxl")
install_if_missing("folium")
install_if_missing("geopy")
install_if_missing("python-dotenv")


# ============================================================
# FIX SSL CERTIFICATE ISSUES
# ============================================================

print("Upgrading certificate-related packages...")

subprocess.check_call([
    sys.executable, "-m", "pip", "install", "-q", "--upgrade",
    "certifi", "requests", "urllib3"
])

import certifi

os.environ["SSL_CERT_FILE"] = certifi.where()
os.environ["REQUESTS_CA_BUNDLE"] = certifi.where()
os.environ["CURL_CA_BUNDLE"] = certifi.where()

print("Using certifi CA bundle:")
print(certifi.where())


# ============================================================
# IMPORT PACKAGES
# ============================================================

import pandas as pd
import requests
from bs4 import BeautifulSoup
from langdetect import detect, LangDetectException
from openai import OpenAI
from dateutil import parser as dateparser
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import folium
from folium.plugins import MarkerCluster
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderUnavailable

try:
    from tqdm import tqdm
except Exception:
    tqdm = None


# ============================================================
# LLM PROVIDER, MODEL, AND API KEY
# ============================================================

ARC_BASE_URL = "https://llm-api.arc.vt.edu/api/v1"


def prompt_required_secret(label: str) -> str:
    """Read a required secret without echoing or persisting it."""
    while True:
        value = getpass(label).strip()
        if value:
            return value
        print("An API key is required.")


def prompt_llm_settings():
    """Choose OpenAI or Virginia Tech ARC, a model, and a runtime API key."""
    print("\nLLM provider:")
    print("  1. OpenAI API (default)")
    print("  2. Virginia Tech ARC LLM API")

    while True:
        provider_choice = input("Choose provider [1]: ").strip() or "1"
        if provider_choice in {"1", "2"}:
            break
        print("Please enter 1 or 2.")

    if provider_choice == "1":
        models = {
            "1": "gpt-5.6-luna",
            "2": "gpt-5.6-terra",
            "3": "gpt-5.6-sol",
        }
        print("\nOpenAI model:")
        print("  1. gpt-5.6-luna — cost-sensitive/high-volume (default)")
        print("  2. gpt-5.6-terra — balance of capability and cost")
        print("  3. gpt-5.6-sol — flagship capability")
        print("  4. Enter another OpenAI model ID")

        while True:
            model_choice = input("Choose OpenAI model [1]: ").strip() or "1"
            if model_choice in models:
                model = models[model_choice]
                break
            if model_choice == "4":
                model = input("OpenAI model ID: ").strip()
                if model:
                    break
            print("Please enter a number from 1 to 4.")

        print("Enter an OpenAI Platform API key; input is hidden and not saved.")
        api_key = prompt_required_secret("OpenAI API key: ")
        return "openai", model, "", api_key

    models = {
        "1": "gpt-oss-120b",
        "2": "DeepSeek-V4-Flash",
        "3": "GLM-5.2",
        "4": "Kimi-K3",
    }
    print("\nVirginia Tech ARC model:")
    print("  1. gpt-oss-120b (default)")
    print("  2. DeepSeek-V4-Flash")
    print("  3. GLM-5.2")
    print("  4. Kimi-K3")
    print("  5. Enter another ARC model ID")

    while True:
        model_choice = input("Choose ARC model [1]: ").strip() or "1"
        if model_choice in models:
            model = models[model_choice]
            break
        if model_choice == "5":
            model = input("ARC model ID: ").strip()
            if model:
                break
        print("Please enter a number from 1 to 5.")

    print(
        "Create a personal key at https://llm.arc.vt.edu under "
        "User profile > Settings > Account > API keys."
    )
    print("Input is hidden and the key is not saved.")
    api_key = prompt_required_secret("ARC personal API key: ")
    return "arc", model, ARC_BASE_URL, api_key


def create_llm_client() -> OpenAI:
    """Create a client for the provider selected at startup."""
    options = {"api_key": LLM_API_KEY}
    if LLM_BASE_URL:
        options["base_url"] = LLM_BASE_URL
    return OpenAI(**options)


def generate_llm_text(client: OpenAI, model: str, prompt: str) -> str:
    """Generate text through OpenAI Responses or ARC Chat Completions."""
    if LLM_PROVIDER == "arc":
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=8_000,
        )
        return (response.choices[0].message.content or "").strip()

    response = client.responses.create(model=model, input=prompt)
    return response.output_text.strip()


def ensure_llm_configured():
    """Prompt once, including when only search-term translation needs an LLM."""
    global LLM_PROVIDER, LLM_MODEL, LLM_BASE_URL, LLM_API_KEY
    if LLM_API_KEY:
        return
    LLM_PROVIDER, LLM_MODEL, LLM_BASE_URL, LLM_API_KEY = prompt_llm_settings()
    print(f"\nUsing {LLM_PROVIDER.upper()} model {LLM_MODEL}.")
    print("API key accepted for this run (not saved).")


if TRANSLATE_POSTS or INFER_LOCATIONS:
    ensure_llm_configured()
else:
    print(
        "Post translation/location inference are off. An LLM key will be "
        "requested only if you choose translated search terms."
    )


# ============================================================
# DATA STRUCTURE
# ============================================================

@dataclass
class PostRecord:
    query: str
    source_mode: str
    nitter_instance: str
    scraped_from: str
    nitter_url: str
    x_url: str
    tweet_id: str
    date_raw: str
    date_iso: str
    username: str
    display_name: str
    detected_language: str
    original_text: str
    translated_en: str
    raw_stats: str
    is_retweet: bool
    inferred_location: str
    location_confidence: float
    location_source: str
    location_reason: str
    latitude: str
    longitude: str
    geocode_display_name: str


# ============================================================
# BASIC HELPERS
# ============================================================

def clean_handle(handle: str) -> str:
    handle = handle.strip()
    if handle.startswith("@"):
        handle = handle[1:]
    return handle


def normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def clean_cell(value):
    """
    Make social media text safer and cleaner for CSV/Excel output.
    Removes internal line breaks, normalizes spacing, and prevents
    spreadsheet formula injection.
    """

    if value is None:
        return ""

    value = str(value)

    value = value.replace("\r", " ")
    value = value.replace("\n", " ")
    value = value.replace("\t", " ")

    value = normalize_whitespace(value)

    # Prevent spreadsheet formula injection
    if value.startswith(("=", "+", "-", "@")):
        value = "'" + value

    return value


def detect_language_safe(text: str) -> str:
    try:
        return detect(text)
    except LangDetectException:
        return "unknown"
    except Exception:
        return "unknown"


def parse_nitter_date(date_text: str) -> str:
    if not date_text:
        return ""

    cleaned = date_text.replace("·", " ").strip()

    try:
        dt = dateparser.parse(cleaned, fuzzy=True)
        if dt:
            return dt.isoformat()
    except Exception:
        return ""

    return ""


def build_queries(
    terms: List[str],
    handles: Optional[List[str]] = None,
) -> List[str]:
    if handles:
        queries = []

        for handle in handles:
            h = clean_handle(handle)
            queries.append(f"from:{h}")

        return queries

    return [term.strip() for term in terms if term.strip()]


def make_nitter_search_url(
    base_url: str,
    query: str,
    since: Optional[str] = None,
    until: Optional[str] = None,
) -> str:
    base_url = base_url.rstrip("/")

    q = query.strip()

    q = re.sub(r"\bsince:\d{4}-\d{2}-\d{2}\b", "", q)
    q = re.sub(r"\buntil:\d{4}-\d{2}-\d{2}\b", "", q)
    q = q.replace("-filter:retweets", "")
    q = normalize_whitespace(q)

    if q.startswith('"') and q.endswith('"'):
        q = q[1:-1]

    params = {
        "f": "tweets",
        "q": q,
    }

    if since:
        params["since"] = since

    if until:
        params["until"] = until

    return f"{base_url}/search?{urlencode(params)}"


def make_x_url_from_nitter_url(nitter_url: str) -> str:
    try:
        parsed = urlparse(nitter_url)
        return "https://x.com" + parsed.path
    except Exception:
        return ""


def extract_tweet_id_from_url(url: str) -> str:
    match = re.search(r"/status/(\d+)", url)
    if match:
        return match.group(1)
    return ""


def get_text_or_empty(element) -> str:
    if element is None:
        return ""
    return element.get_text(" ", strip=True)


def prompt_for_search_terms(default_terms: List[str]) -> List[str]:
    """Collect arbitrary search terms, one per line, from the user."""

    print("\n" + "=" * 70)
    print("SEARCH SETUP")
    print("=" * 70)
    print("Enter any topics, names, phrases, or hashtags to search for.")
    print("Enter one term per line, then press Enter on a blank line to finish.")
    print(f"Press Enter immediately to use the default: {', '.join(default_terms)}")

    terms = []

    while True:
        try:
            value = input("Search term: ").strip()
        except EOFError:
            print("No interactive input available; using the default search terms.")
            return list(default_terms)

        if not value:
            break

        if value not in terms:
            terms.append(value)

    return terms or list(default_terms)


def prompt_for_search_languages() -> List[str]:
    """Let the user select languages for translated search-query variants."""

    print("\nThe original search terms will always be used.")
    print("Optionally translate them into additional search languages:")

    for number, language in SEARCH_LANGUAGE_OPTIONS.items():
        print(f"  {number:>2}. {language}")

    print("Enter numbers separated by commas (for example: 2,4,6).")
    print("Press Enter to search only the original terms.")

    try:
        raw_choices = input("Language choices: ").strip()
    except EOFError:
        return []

    if not raw_choices:
        return []

    selected = []
    invalid = []

    for choice in re.split(r"[,\s]+", raw_choices):
        if not choice:
            continue

        language = SEARCH_LANGUAGE_OPTIONS.get(choice)
        if language and language not in selected:
            selected.append(language)
        elif not language:
            invalid.append(choice)

    if invalid:
        print(f"Ignoring unrecognized choices: {', '.join(invalid)}")

    return selected


def translate_search_term(
    client: OpenAI,
    term: str,
    target_language: str,
    model: str,
    max_retries: int = 3,
) -> str:
    """Translate one query while preserving search syntax where practical."""

    prompt = f"""
Translate this social-media search term into {target_language}.

Return only the translated search term, with no quotation marks, label, or explanation.
Preserve hashtags, @handles, URLs, Boolean operators, and date/filter syntax exactly.
If the term is already in {target_language}, return it unchanged.

Search term:
{term}
""".strip()

    for attempt in range(1, max_retries + 1):
        try:
            translated = normalize_whitespace(
                generate_llm_text(client=client, model=model, prompt=prompt)
            )
            return translated.strip('"').strip("'")
        except Exception as e:
            if attempt == max_retries:
                print(
                    f"Could not translate {term!r} into {target_language}: {e}. "
                    "Skipping that query variant."
                )
                return ""

            time.sleep(2 * attempt)

    return ""


def build_interactive_search_terms() -> List[str]:
    """Prompt for terms/languages and return unique original + translated queries."""

    original_terms = prompt_for_search_terms(DEFAULT_SEARCH_TERMS)
    target_languages = prompt_for_search_languages()
    all_terms = list(original_terms)

    if target_languages:
        ensure_llm_configured()
        client = create_llm_client()
        print("\nTranslating search terms...")

        for term in original_terms:
            for language in target_languages:
                translated = translate_search_term(
                    client=client,
                    term=term,
                    target_language=language,
                    model=LLM_MODEL,
                )

                if translated and translated.casefold() not in {
                    existing.casefold() for existing in all_terms
                }:
                    all_terms.append(translated)
                    print(f"  {language}: {term!r} -> {translated!r}")

    print("\nSearch queries that will be used:")
    for term in all_terms:
        print(f"  - {term}")

    return all_terms


# ============================================================
# HTTP / NITTER FETCHING
# ============================================================

def create_requests_session() -> requests.Session:
    session = requests.Session()

    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 "
            "(Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "keep-alive",
    })

    return session


def fetch_html(
    session: requests.Session,
    url: str,
    timeout: int = 30,
    max_retries: int = 2,
) -> Optional[str]:
    for attempt in range(1, max_retries + 1):
        try:
            print(f"Fetching page: {url}")

            response = session.get(
                url,
                timeout=timeout,
                verify=certifi.where(),
            )

            status = response.status_code
            text = response.text or ""

            print(f"Final URL: {response.url}")
            print(f"Status code: {status}")
            print(f"Response length: {len(text)} characters")
            print(f"Content-Type: {response.headers.get('content-type', '')}")

            if status != 200:
                preview = normalize_whitespace(text[:3000])
                print("\n--- NON-200 RESPONSE PREVIEW ---")
                print(preview)
                print("--- END PREVIEW ---\n")

            if status == 200:
                if len(text.strip()) == 0:
                    print("Instance returned a blank page.")
                    return None

                return text

            if status in [403, 404, 429, 503]:
                print(f"Instance returned {status}. Not bypassing; stopping this request.")
                return None

            print(f"Unexpected status code {status}. Attempt {attempt}/{max_retries}.")

        except requests.exceptions.RequestException as e:
            print(f"Request error on attempt {attempt}/{max_retries}: {repr(e)}")

        if attempt < max_retries:
            sleep_for = REQUEST_DELAY_SECONDS * attempt
            print(f"Sleeping {sleep_for} seconds before retry...")
            time.sleep(sleep_for)

    return None


def debug_nitter_page(html: str, max_chars: int = 3000):
    soup = BeautifulSoup(html, "html.parser")

    title = soup.find("title")
    title_text = title.get_text(" ", strip=True) if title else "[no title]"

    body_text = soup.get_text(" ", strip=True)
    body_text = normalize_whitespace(body_text)

    print("\n--- NITTER PAGE DEBUG ---")
    print("Title:", title_text)
    print("First body text:")
    print(body_text[:max_chars])
    print("\nDetected timeline item count:", len(soup.select(".timeline-item")))
    print("Detected tweet-content count:", len(soup.select(".tweet-content")))
    print("Detected timeline count:", len(soup.select(".timeline")))
    print("Detected show-more count:", len(soup.select(".show-more")))
    print("--- END DEBUG ---\n")


def find_next_nitter_page_url(
    html: str,
    current_url: str,
) -> Optional[str]:
    soup = BeautifulSoup(html, "html.parser")

    candidate_links = []

    candidate_links.extend(soup.select("div.show-more a"))
    candidate_links.extend(soup.select("a[href*='cursor=']"))

    for link in candidate_links:
        href = link.get("href")

        if not href:
            continue

        if "cursor=" in href or "search?" in href:
            return urljoin(current_url, href)

    return None


def parse_tweets_from_nitter_html(
    html: str,
    page_url: str,
    query: str,
) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "html.parser")
    items = soup.select(".timeline-item")

    parsed = []

    for item in items:
        content_el = item.select_one(".tweet-content")

        if content_el is None:
            continue

        text = content_el.get_text("\n", strip=True)
        text = text.strip()

        if not text:
            continue

        username_el = item.select_one(".username")
        fullname_el = item.select_one(".fullname")
        date_link = item.select_one(".tweet-date a")
        tweet_link = item.select_one("a.tweet-link")

        username = get_text_or_empty(username_el).replace("@", "").strip()
        display_name = get_text_or_empty(fullname_el)

        date_raw = ""
        if date_link is not None:
            date_raw = date_link.get("title") or get_text_or_empty(date_link)

        date_iso = parse_nitter_date(date_raw)

        tweet_href = ""
        if tweet_link is not None:
            tweet_href = tweet_link.get("href") or ""

        nitter_url = urljoin(page_url, tweet_href) if tweet_href else ""
        x_url = make_x_url_from_nitter_url(nitter_url) if nitter_url else ""
        tweet_id = extract_tweet_id_from_url(nitter_url)

        stats_el = item.select_one(".tweet-stats")
        raw_stats = normalize_whitespace(get_text_or_empty(stats_el))

        retweet_header = item.select_one(".retweet-header")
        is_retweet = retweet_header is not None

        parsed.append({
            "query": query,
            "scraped_from": page_url,
            "nitter_url": nitter_url,
            "x_url": x_url,
            "tweet_id": tweet_id,
            "date_raw": date_raw,
            "date_iso": date_iso,
            "username": username,
            "display_name": display_name,
            "original_text": text,
            "raw_stats": raw_stats,
            "is_retweet": is_retweet,
        })

    return parsed


# ============================================================
# LLM TRANSLATION
# ============================================================

def translate_with_openai(
    client: OpenAI,
    text: str,
    source_lang: str,
    model: str = "gpt-5.6-luna",
    target_language: str = "English",
    max_retries: int = 3,
) -> str:
    if not isinstance(text, str) or not text.strip():
        return ""

    if source_lang == "en" and target_language.lower() == "english":
        return text

    prompt = f"""
Translate the following public social media post into {target_language}.

Rules:
- Preserve proper names, institution names, places, hashtags, @handles, emojis, and URLs.
- Preserve dates and numbers exactly.
- Do not summarize.
- Do not add commentary.
- If the post is already in {target_language}, return it unchanged.
- If the post contains mixed languages, translate only the non-{target_language} parts.
- Keep the tone close to the original.

Detected source language: {source_lang}

Post:
{text}
""".strip()

    for attempt in range(1, max_retries + 1):
        try:
            return generate_llm_text(client=client, model=model, prompt=prompt)

        except Exception as e:
            if attempt == max_retries:
                return f"[TRANSLATION ERROR: {e}] {text}"

            wait_time = 2 * attempt
            print(f"LLM translation error. Retrying in {wait_time} seconds...")
            time.sleep(wait_time)

    return text


# ============================================================
# LOCATION INFERENCE + GEOCODING + MAP HELPERS
# ============================================================

def load_json_cache(path: str) -> dict:
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    return {}


def save_json_cache(path: str, data: dict):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_language_geo_hint(lang_code: str) -> str:
    """
    Give the model a broad geographic prior based on detected language.
    These are intentionally broad and should not be treated as proof.
    """

    lang_code = (lang_code or "").lower().strip()

    hints = {
        "en": "English is global and should be treated as a weak location signal only.",
        "es": "Spanish suggests Spain or Latin America; use other evidence to distinguish.",
        "fr": "French suggests France, Belgium, Switzerland, Canada/Quebec, parts of West Africa, North Africa, or other Francophone regions.",
        "pt": "Portuguese suggests Portugal, Brazil, Angola, Mozambique, or other Lusophone regions.",
        "it": "Italian suggests Italy or Italian-language institutional contexts.",
        "de": "German suggests Germany, Austria, Switzerland, or German-language institutional contexts.",
        "ru": "Russian suggests Russia or Russian-speaking post-Soviet regions.",
        "ar": "Arabic suggests Arabic-speaking countries in the Middle East or North Africa.",
        "tr": "Turkish suggests Turkey or Turkish-speaking communities.",
        "ja": "Japanese suggests Japan or Japanese-language institutional contexts.",
        "ko": "Korean suggests South Korea or Korean-language institutional contexts.",
        "zh-cn": "Simplified Chinese suggests Mainland China, Singapore, Malaysia, or simplified-Chinese institutional contexts.",
        "zh-tw": "Traditional Chinese suggests Taiwan, Hong Kong, Macau, or traditional-Chinese institutional contexts.",
        "zh": "Chinese suggests a Chinese-language context; use script, institution names, and other evidence to distinguish Mainland China, Taiwan, Hong Kong, Macau, Singapore, Malaysia, or diaspora contexts.",
    }

    return hints.get(lang_code, "No strong language-based geographic prior available.")


def infer_location_with_openai(
    client: OpenAI,
    post: Dict[str, Any],
    model: str = "gpt-5.6-luna",
    max_retries: int = 3,
) -> Dict[str, Any]:
    """
    Infer a broad public location from tweet/account metadata.

    This intentionally avoids exact street-level locations.
    It should return city/region/country-level places only.
    """

    username = post.get("username", "")
    display_name = post.get("display_name", "")
    query = post.get("query", "")
    original_text = post.get("original_text", "")
    translated_text = post.get("translated_en", "")
    detected_language = post.get("detected_language", "unknown")
    language_geo_hint = get_language_geo_hint(detected_language)

    prompt = f"""
You are helping map public institutional social media posts.

Infer the most likely broad public location associated with this post.
Use only city, region, university, or country-level locations.
Do NOT infer private addresses, homes, street addresses, or precise personal locations.

Return ONLY valid JSON with these keys:
- location_name: string, a geocodable place such as "Madrid, Spain" or "University of Nairobi, Kenya"; use "" if unknown
- confidence: number between 0 and 1
- source: string, one of ["language", "username", "display_name", "tweet_text", "query", "combined", "unknown"]
- reason: short string explaining the evidence

Location inference rules:
1. Treat the language of the original post as a major geographic prior.
2. Use the detected language and script to narrow the likely region before considering weaker clues.
3. If the original language strongly aligns with a country, region, institution name, or search query, increase confidence.
4. If the language is widely used across many countries, such as English, French, Spanish, Portuguese, Arabic, or Chinese, do NOT choose a specific country based on language alone.
5. Prefer explicit place names in the original post text when present.
6. Next prefer institution names, university names, account names, or handles.
7. Use the translated text only to understand meaning; prioritize the original post for language, script, and named entities.
8. If language is the only evidence, return a broad country/region-level location with low confidence, usually 0.15 to 0.35.
9. If language plus account/institution/text evidence all point to the same place, confidence may be higher.
10. If no responsible location can be inferred, return location_name "" and confidence 0.

Detected language:
{detected_language}

Language-based geographic prior:
{language_geo_hint}

Metadata:
username: {username}
display_name: {display_name}
search_query: {query}

Original post:
{original_text}

Translated post:
{translated_text}
""".strip()

    for attempt in range(1, max_retries + 1):
        try:
            text = generate_llm_text(client=client, model=model, prompt=prompt)
            text = text.replace("```json", "").replace("```", "").strip()

            data = json.loads(text)

            return {
                "location_name": str(data.get("location_name", "")).strip(),
                "confidence": float(data.get("confidence", 0) or 0),
                "source": str(data.get("source", "unknown")).strip(),
                "reason": str(data.get("reason", "")).strip(),
            }

        except Exception as e:
            if attempt == max_retries:
                return {
                    "location_name": "",
                    "confidence": 0.0,
                    "source": "unknown",
                    "reason": f"Location inference failed: {e}",
                }

            time.sleep(2 * attempt)

    return {
        "location_name": "",
        "confidence": 0.0,
        "source": "unknown",
        "reason": "Location inference failed.",
    }


def geocode_location(
    location_name: str,
    cache: dict,
    geolocator: Nominatim,
    delay_seconds: float = 1.2,
) -> Dict[str, str]:
    """
    Geocode a location with Nominatim using local caching.
    """

    location_name = normalize_whitespace(location_name)

    if not location_name:
        return {
            "latitude": "",
            "longitude": "",
            "geocode_display_name": "",
        }

    if location_name in cache:
        return cache[location_name]

    try:
        print(f"Geocoding: {location_name}")

        result = geolocator.geocode(
            location_name,
            addressdetails=False,
            exactly_one=True,
            timeout=20,
        )

        time.sleep(delay_seconds)

        if result:
            data = {
                "latitude": str(result.latitude),
                "longitude": str(result.longitude),
                "geocode_display_name": result.address or location_name,
            }
        else:
            data = {
                "latitude": "",
                "longitude": "",
                "geocode_display_name": "",
            }

        cache[location_name] = data
        save_json_cache(GEOCODE_CACHE_FILE, cache)

        return data

    except (GeocoderTimedOut, GeocoderUnavailable) as e:
        print(f"Geocoding failed for {location_name}: {e}")

        return {
            "latitude": "",
            "longitude": "",
            "geocode_display_name": "",
        }

    except Exception as e:
        print(f"Unexpected geocoding error for {location_name}: {e}")

        return {
            "latitude": "",
            "longitude": "",
            "geocode_display_name": "",
        }


def enrich_records_with_locations(
    records: List[PostRecord],
    openai_model: str = "gpt-5.6-luna",
) -> List[PostRecord]:
    """
    Infer broad locations and geocode them.
    """

    if not records:
        return records

    if not INFER_LOCATIONS:
        return records

    print("\nInferring and geocoding locations...")

    client = create_llm_client()
    geolocator = Nominatim(user_agent=NOMINATIM_USER_AGENT)
    geocode_cache = load_json_cache(GEOCODE_CACHE_FILE)

    enriched_records = []

    iterator = records
    if tqdm is not None:
        iterator = tqdm(records)

    for record in iterator:
        post_dict = asdict(record)

        inference = infer_location_with_openai(
            client=client,
            post=post_dict,
            model=openai_model,
        )

        inferred_location = inference.get("location_name", "")
        confidence = float(inference.get("confidence", 0) or 0)
        source = inference.get("source", "unknown")
        reason = inference.get("reason", "")

        latitude = ""
        longitude = ""
        geocode_display_name = ""

        if inferred_location and confidence >= MIN_LOCATION_CONFIDENCE:
            geocode = geocode_location(
                location_name=inferred_location,
                cache=geocode_cache,
                geolocator=geolocator,
                delay_seconds=GEOCODE_DELAY_SECONDS,
            )

            latitude = geocode.get("latitude", "")
            longitude = geocode.get("longitude", "")
            geocode_display_name = geocode.get("geocode_display_name", "")

        record.inferred_location = inferred_location
        record.location_confidence = confidence
        record.location_source = source
        record.location_reason = reason
        record.latitude = latitude
        record.longitude = longitude
        record.geocode_display_name = geocode_display_name

        enriched_records.append(record)

    return enriched_records


def create_tweet_map(
    df: pd.DataFrame,
    map_file: str = "nitter_search_map.html",
):
    """
    Create an interactive Folium map with popups showing original and translated tweets.

    Tile fix:
    - Do not use Folium's direct OpenStreetMap default tile layer.
    - That default can trigger 503r / Access Blocked errors from the OSM tile server.
    - Use CARTO basemap tiles instead, with proper attribution.
    """

    if df is None or df.empty:
        print("No dataframe available for mapping.")
        return None

    if "latitude" not in df.columns or "longitude" not in df.columns:
        print("No latitude/longitude columns found. Skipping map.")
        return None

    map_df = df.copy()

    map_df["latitude_num"] = pd.to_numeric(map_df["latitude"], errors="coerce")
    map_df["longitude_num"] = pd.to_numeric(map_df["longitude"], errors="coerce")

    map_df = map_df.dropna(subset=["latitude_num", "longitude_num"])

    if map_df.empty:
        print("No geocoded rows available for mapping.")
        return None

    center_lat = map_df["latitude_num"].mean()
    center_lon = map_df["longitude_num"].mean()

    # IMPORTANT:
    # tiles=None prevents Folium from adding its default OpenStreetMap layer.
    # The default OSM tile endpoint is the source of the 503r access-blocked errors.
    m = folium.Map(
        location=[center_lat, center_lon],
        zoom_start=2,
        tiles=None,
        control_scale=True,
    )

    # Primary basemap: CARTO Positron. This avoids direct requests to
    # tile.openstreetmap.org while still providing a clean global basemap.
    folium.TileLayer(
        tiles=MAP_BASE_TILE_URL,
        attr=MAP_BASE_TILE_ATTRIBUTION,
        name=MAP_BASE_TILE_NAME,
        overlay=False,
        control=True,
    ).add_to(m)

    # Optional alternate basemaps. These are useful if one tile provider is slow
    # or temporarily unavailable. None of these uses the direct OSM tile server.
    if ADD_ALTERNATE_BASEMAPS:
        folium.TileLayer(
            tiles="https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png",
            attr=MAP_BASE_TILE_ATTRIBUTION,
            name="CartoDB Dark Matter",
            overlay=False,
            control=True,
        ).add_to(m)

        folium.TileLayer(
            tiles=(
                "https://server.arcgisonline.com/ArcGIS/rest/services/"
                "World_Imagery/MapServer/tile/{z}/{y}/{x}"
            ),
            attr=(
                "Tiles &copy; Esri &mdash; Source: Esri, i-cubed, USDA, USGS, "
                "AEX, GeoEye, Getmapping, Aerogrid, IGN, IGP, UPR-EGP, and the GIS User Community"
            ),
            name="Esri World Imagery",
            overlay=False,
            control=True,
        ).add_to(m)

    cluster = MarkerCluster(name="Inferred tweet locations").add_to(m)

    for _, row in map_df.iterrows():
        original_text = html_lib.escape(str(row.get("original_text", "")))
        translated_en = html_lib.escape(str(row.get("translated_en", "")))
        username = html_lib.escape(str(row.get("username", "")))
        display_name = html_lib.escape(str(row.get("display_name", "")))
        date_raw = html_lib.escape(str(row.get("date_raw", "")))
        inferred_location = html_lib.escape(str(row.get("inferred_location", "")))
        confidence = html_lib.escape(str(row.get("location_confidence", "")))
        source = html_lib.escape(str(row.get("location_source", "")))
        reason = html_lib.escape(str(row.get("location_reason", "")))
        x_url = html_lib.escape(str(row.get("x_url", "")))
        nitter_url = html_lib.escape(str(row.get("nitter_url", "")))

        popup_html = f"""
        <div style="width: 380px; font-family: Arial, sans-serif;">
            <h4 style="margin-bottom: 4px;">@{username}</h4>
            <div><strong>Name:</strong> {display_name}</div>
            <div><strong>Date:</strong> {date_raw}</div>
            <div><strong>Inferred location:</strong> {inferred_location}</div>
            <div><strong>Confidence:</strong> {confidence}</div>
            <div><strong>Source:</strong> {source}</div>
            <div><strong>Reason:</strong> {reason}</div>
            <hr>
            <div><strong>Original:</strong></div>
            <div style="white-space: pre-wrap; margin-bottom: 8px;">{original_text}</div>
            <div><strong>Translation:</strong></div>
            <div style="white-space: pre-wrap; margin-bottom: 8px;">{translated_en}</div>
            <hr>
            <div><a href="{x_url}" target="_blank">Open on X</a></div>
            <div><a href="{nitter_url}" target="_blank">Open on Nitter</a></div>
        </div>
        """

        tooltip = f"@{username} — {inferred_location}"

        folium.Marker(
            location=[row["latitude_num"], row["longitude_num"]],
            popup=folium.Popup(popup_html, max_width=440),
            tooltip=tooltip,
        ).add_to(cluster)

    folium.LayerControl().add_to(m)

    m.save(map_file)

    print(f"Map saved to: {map_file}")
    print("Map tile layer: CARTO basemap CDN, not the direct OpenStreetMap tile server.")

    return m


# ============================================================
# DIAGNOSTIC FUNCTION
# ============================================================

def run_instance_diagnostic(test_query: str):
    print("\n" + "=" * 70)
    print("Running Nitter instance diagnostic...")
    print("=" * 70)

    session = create_requests_session()

    for base_url in NITTER_BASE_URLS:
        test_url = make_nitter_search_url(
            base_url=base_url,
            query=test_query,
            since=None,
            until=None,
        )

        print("\n" + "-" * 70)
        print(f"Testing instance: {base_url}")
        print(f"Test URL: {test_url}")

        html = fetch_html(session, test_url)

        if html:
            print("Usable HTML returned.")
            debug_nitter_page(html, max_chars=3000)
        else:
            print("No usable HTML returned.")


# ============================================================
# RECORD CREATION
# ============================================================

def make_record_from_post(
    post: Dict[str, Any],
    source_mode: str,
    nitter_instance: str,
    client: Optional[OpenAI],
    openai_model: str,
    target_language: str,
    translate: bool,
) -> PostRecord:
    original_text = post.get("original_text", "")
    detected_language = detect_language_safe(original_text)

    if translate and client is not None:
        translated_en = translate_with_openai(
            client=client,
            text=original_text,
            source_lang=detected_language,
            model=openai_model,
            target_language=target_language,
        )
        time.sleep(TRANSLATION_DELAY_SECONDS)
    else:
        translated_en = ""

    return PostRecord(
        query=post.get("query", ""),
        source_mode=source_mode,
        nitter_instance=nitter_instance,
        scraped_from=post.get("scraped_from", ""),
        nitter_url=post.get("nitter_url", ""),
        x_url=post.get("x_url", ""),
        tweet_id=post.get("tweet_id", ""),
        date_raw=post.get("date_raw", ""),
        date_iso=post.get("date_iso", ""),
        username=post.get("username", ""),
        display_name=post.get("display_name", ""),
        detected_language=detected_language,
        original_text=original_text,
        translated_en=translated_en,
        raw_stats=post.get("raw_stats", ""),
        is_retweet=post.get("is_retweet", False),
        inferred_location="",
        location_confidence=0.0,
        location_source="",
        location_reason="",
        latitude="",
        longitude="",
        geocode_display_name="",
    )


# ============================================================
# LIVE NITTER COLLECTION
# ============================================================

def run_live_nitter_collection(
    nitter_base_urls: List[str],
    since: Optional[str] = None,
    until: Optional[str] = None,
    max_posts_per_query: int = 5,
    max_pages_per_query: int = 1,
    output_file: str = "nitter_search_posts.csv",
    handles: Optional[List[str]] = None,
    search_terms: Optional[List[str]] = None,
    include_retweets: bool = False,
    target_language: str = "English",
    openai_model: str = "gpt-5.6-luna",
    translate: bool = False,
):
    print("Live Nitter collection started.")

    session = create_requests_session()

    if translate:
        print("Translation is ON.")
        client = create_llm_client()
    else:
        print("Translation is OFF.")
        client = None

    if search_terms is None:
        search_terms = DEFAULT_SEARCH_TERMS

    queries = build_queries(
        terms=search_terms,
        handles=handles,
    )

    print(f"\nBuilt {len(queries)} queries:")
    for q in queries:
        print("  ", q)

    records: List[PostRecord] = []
    seen_ids = set()
    seen_text_fallbacks = set()

    for query in queries:
        print("\n" + "=" * 70)
        print(f"Searching Nitter for: {query}")

        posts_collected_for_query = 0

        for base_url in nitter_base_urls:
            print("\n" + "-" * 70)
            print(f"Trying Nitter instance: {base_url}")

            page_url = make_nitter_search_url(
                base_url=base_url,
                query=query,
                since=since,
                until=until,
            )

            for page_num in range(1, max_pages_per_query + 1):
                print(f"\nPage {page_num}/{max_pages_per_query}")

                html = fetch_html(session, page_url)

                if html is None:
                    print("No usable HTML returned from this page.")
                    break

                parsed_posts = parse_tweets_from_nitter_html(
                    html=html,
                    page_url=page_url,
                    query=query,
                )

                print(f"Parsed {len(parsed_posts)} posts from page.")

                if not parsed_posts:
                    print("No posts found on this page. Showing page debug:")
                    debug_nitter_page(html)
                    break

                for post in parsed_posts:
                    if not include_retweets and post.get("is_retweet"):
                        print("Skipping retweet.")
                        continue

                    tweet_id = post.get("tweet_id", "")

                    fallback_key = (
                        post.get("username", ""),
                        post.get("date_raw", ""),
                        post.get("original_text", "")[:120],
                    )

                    if tweet_id and tweet_id in seen_ids:
                        continue

                    if not tweet_id and fallback_key in seen_text_fallbacks:
                        continue

                    if tweet_id:
                        seen_ids.add(tweet_id)
                    else:
                        seen_text_fallbacks.add(fallback_key)

                    record = make_record_from_post(
                        post=post,
                        source_mode="live",
                        nitter_instance=base_url,
                        client=client,
                        openai_model=openai_model,
                        target_language=target_language,
                        translate=translate,
                    )

                    records.append(record)
                    posts_collected_for_query += 1

                    print(
                        f"Collected {posts_collected_for_query}/{max_posts_per_query} "
                        f"for this query: @{record.username} {record.date_raw}"
                    )

                    if posts_collected_for_query >= max_posts_per_query:
                        break

                if posts_collected_for_query >= max_posts_per_query:
                    print("Reached max posts for this query.")
                    break

                next_url = find_next_nitter_page_url(html, page_url)

                if not next_url:
                    print("No next page found.")
                    break

                page_url = next_url

                print(f"Sleeping {REQUEST_DELAY_SECONDS} seconds before next page...")
                time.sleep(REQUEST_DELAY_SECONDS)

            if posts_collected_for_query >= max_posts_per_query:
                break

            print(f"Sleeping {REQUEST_DELAY_SECONDS} seconds before trying next instance...")
            time.sleep(REQUEST_DELAY_SECONDS)

    return save_records(records, output_file)


# ============================================================
# SAVED HTML COLLECTION
# ============================================================

def run_saved_html_collection(
    html_files: List[str],
    output_file: str,
    include_retweets: bool = False,
    target_language: str = "English",
    openai_model: str = "gpt-5.6-luna",
    translate: bool = False,
):
    print("Saved HTML collection started.")

    if translate:
        print("Translation is ON.")
        client = create_llm_client()
    else:
        print("Translation is OFF.")
        client = None

    records: List[PostRecord] = []
    seen_ids = set()
    seen_text_fallbacks = set()

    for html_file in html_files:
        print("\n" + "=" * 70)
        print(f"Parsing saved HTML file: {html_file}")

        if not os.path.exists(html_file):
            print(f"File not found: {html_file}")
            continue

        with open(html_file, "r", encoding="utf-8", errors="replace") as f:
            html = f.read()

        debug_nitter_page(html, max_chars=1500)

        parsed_posts = parse_tweets_from_nitter_html(
            html=html,
            page_url="https://nitter.net/search",
            query=f"saved_html:{html_file}",
        )

        print(f"Parsed {len(parsed_posts)} posts from saved HTML.")

        for post in parsed_posts:
            if not include_retweets and post.get("is_retweet"):
                print("Skipping retweet.")
                continue

            tweet_id = post.get("tweet_id", "")

            fallback_key = (
                post.get("username", ""),
                post.get("date_raw", ""),
                post.get("original_text", "")[:120],
            )

            if tweet_id and tweet_id in seen_ids:
                continue

            if not tweet_id and fallback_key in seen_text_fallbacks:
                continue

            if tweet_id:
                seen_ids.add(tweet_id)
            else:
                seen_text_fallbacks.add(fallback_key)

            record = make_record_from_post(
                post=post,
                source_mode="saved_html",
                nitter_instance="saved_html",
                client=client,
                openai_model=openai_model,
                target_language=target_language,
                translate=translate,
            )

            records.append(record)

            print(f"Collected saved HTML post: @{record.username} {record.date_raw}")

    return save_records(records, output_file)


# ============================================================
# SAVE OUTPUT
# ============================================================

def save_records(records: List[PostRecord], output_file: str):
    if not records:
        print("\nNo records collected.")
        print("Likely causes:")
        print("1. The Nitter instance returned blank, blocked, or no-result pages.")
        print("2. The query is too narrow.")
        print("3. Date filters are too restrictive.")
        print("4. Nitter search syntax changed.")
        print("5. Saved HTML file does not contain visible timeline items.")
        return None

    if CREATE_MAP or INFER_LOCATIONS:
        records = enrich_records_with_locations(
            records=records,
            openai_model=LLM_MODEL,
        )

    df = pd.DataFrame([asdict(r) for r in records])

    # Clean all cells for CSV and Excel readability
    for col in df.columns:
        df[col] = df[col].apply(clean_cell)

    # Put the most useful columns first
    preferred_order = [
        "date_iso",
        "date_raw",
        "username",
        "display_name",
        "detected_language",
        "inferred_location",
        "location_confidence",
        "location_source",
        "location_reason",
        "latitude",
        "longitude",
        "geocode_display_name",
        "original_text",
        "translated_en",
        "x_url",
        "nitter_url",
        "tweet_id",
        "raw_stats",
        "is_retweet",
        "query",
        "source_mode",
        "nitter_instance",
        "scraped_from",
    ]

    existing_order = [col for col in preferred_order if col in df.columns]
    remaining_cols = [col for col in df.columns if col not in existing_order]
    df = df[existing_order + remaining_cols]

    # Make filenames
    if output_file.lower().endswith(".csv"):
        csv_file = output_file
        xlsx_file = output_file[:-4] + ".xlsx"
    else:
        csv_file = output_file + ".csv"
        xlsx_file = output_file + ".xlsx"

    # Clean CSV export
    df.to_csv(
        csv_file,
        index=False,
        encoding="utf-8-sig",
        quoting=csv.QUOTE_ALL,
        lineterminator="\n",
    )

    # Formatted Excel export
    with pd.ExcelWriter(xlsx_file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="posts")

        worksheet = writer.sheets["posts"]

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        # Header styling
        header_fill = PatternFill("solid", fgColor="D9EAF7")
        header_font = Font(bold=True)

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        # Set column widths by column name
        width_by_column_name = {
            "date_iso": 24,
            "date_raw": 28,
            "username": 20,
            "display_name": 30,
            "detected_language": 18,
            "inferred_location": 34,
            "location_confidence": 18,
            "location_source": 18,
            "location_reason": 60,
            "latitude": 18,
            "longitude": 18,
            "geocode_display_name": 60,
            "original_text": 90,
            "translated_en": 90,
            "x_url": 60,
            "nitter_url": 60,
            "tweet_id": 24,
            "raw_stats": 30,
            "is_retweet": 14,
            "query": 30,
            "source_mode": 16,
            "nitter_instance": 24,
            "scraped_from": 60,
        }

        for idx, column_name in enumerate(df.columns, start=1):
            col_letter = get_column_letter(idx)
            worksheet.column_dimensions[col_letter].width = width_by_column_name.get(
                column_name,
                24,
            )

        # Wrap and top-align all cells
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True,
                )

        # Make rows taller for readability
        for row_idx in range(2, worksheet.max_row + 1):
            worksheet.row_dimensions[row_idx].height = 60

    print(f"\nDone. Wrote {len(df)} records.")
    print(f"CSV file:   {csv_file}")
    print(f"Excel file: {xlsx_file}")

    if CREATE_MAP:
        create_tweet_map(
            df=df,
            map_file=MAP_FILE,
        )

    print("\nPreview:")
    print(df.head())

    return df


# ============================================================
# RUN THE SCRIPT
# ============================================================

if MODE == "live":
    active_search_terms = build_interactive_search_terms()

    if RUN_INSTANCE_DIAGNOSTIC_FIRST:
        run_instance_diagnostic(active_search_terms[0])

    df = run_live_nitter_collection(
        nitter_base_urls=NITTER_BASE_URLS,
        since=SINCE_DATE,
        until=UNTIL_DATE,
        max_posts_per_query=MAX_POSTS_PER_QUERY,
        max_pages_per_query=MAX_PAGES_PER_QUERY,
        output_file=OUTPUT_FILE,
        handles=SEARCH_HANDLES if SEARCH_HANDLES else None,
        search_terms=active_search_terms,
        include_retweets=INCLUDE_RETWEETS,
        target_language=TARGET_LANGUAGE,
        openai_model=LLM_MODEL,
        translate=TRANSLATE_POSTS,
    )

elif MODE == "saved_html":
    df = run_saved_html_collection(
        html_files=SAVED_HTML_FILES,
        output_file=OUTPUT_FILE,
        include_retweets=INCLUDE_RETWEETS,
        target_language=TARGET_LANGUAGE,
        openai_model=LLM_MODEL,
        translate=TRANSLATE_POSTS,
    )

else:
    raise ValueError("MODE must be either 'live' or 'saved_html'.")
