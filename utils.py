import json
import difflib
import requests
import streamlit as st
import numpy as np
from urllib.parse import urlparse
from bs4 import BeautifulSoup
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from sentence_transformers import SentenceTransformer, util
from dotenv import load_dotenv
from docx.oxml.ns import qn
import time
from requests_html import HTMLSession
import asyncio

# Load environment variables from .env if present
load_dotenv()

# Constants
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/98.0.4758.102 Safari/537.36"
    )
}

SCHEMA_CHECKLIST = [
        ("Breadcrumbs", "BreadcrumbList"),
        ("FAQ", "FAQPage"),
        ("Article", "articleBody"),
        ("Video", "VideoObject"),
        ("Organization", "Organization"),
        ("How-to", "HowTo"),
    ]

COMPETITOR_DIRECTORY = {
    "efax": ["Fax.Plus", "CocoFax", "eFax", "iFax", "FaxBurner"],
    "splashtop": ["TeamViewer", "AnyDesk", "ManageEngine", "BeyondTrust", "GoTo", "Splashtop"]
}

YOUTUBE_CHANNEL = {
    "efax": "@eFax", "splashtop": "@SplashtopInc", "fortinet": "@fortinet"
}

CHROMEDRIVER_PATH = "/usr/bin/chromedriver"

# -------------------------------
# Utility and Analysis Functions
# -------------------------------

def is_similar(a, b, threshold=0.8):
    return difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio() >= threshold

def trim_url(url):
    return url.split('#')[0] if url else url

def extract_domain(url):
    parsed_url = urlparse(url)
    domain_parts = parsed_url.netloc.split(".")
    return domain_parts[-2] if len(domain_parts) > 2 else domain_parts[0]

def fetch_page_content(url):
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        return response.text
    except requests.RequestException as e:
        print(f"Error fetching URL: {e}")
        return None

def extract_ai_overview_headers(serp_data):
    headers = []
    if "ai_overview" in serp_data and "text_blocks" in serp_data["ai_overview"]:
        for block in serp_data["ai_overview"]["text_blocks"]:
            if block.get("type") == "paragraph":
                snippet = block.get("snippet", "").strip()
                if snippet.endswith(":") or snippet.istitle():
                    headers.append(snippet.rstrip(":").strip())
            elif block.get("type") == "list":
                for item in block.get("list", []):
                    title = item.get("title", "").strip()
                    if title:
                        headers.append(title.rstrip(":").strip())
    return headers

def compare_headers(page_headers, ai_overview_headers):
    missing = []
    page_header_texts = [h["text"] for h in page_headers]
    for ai_header in ai_overview_headers:
        if not any(is_similar(ai_header, ph) for ph in page_header_texts):
            missing.append(ai_header)
    return missing



def check_domain_in_ai_overview(serp_data, domain, url):
    domain_found = False
    if "ai_overview" in serp_data:
        for block in serp_data["ai_overview"].get("text_blocks", []):
            snippet = block.get("snippet", "").lower()
            if domain in snippet:
                domain_found = True
        for ref in serp_data["ai_overview"].get("references", []):
            if domain in ref.get("link", "").lower():
                domain_found = True
    return domain_found

def find_domain_position_in_organic(serp_data, domain):
    if "organic_results" in serp_data:
        for i, result in enumerate(serp_data["organic_results"]):
            link = result.get("link", "")
            if domain in link.lower():
                return i + 1
    return None

def find_domain_position_in_ai(serp_data, domain):
    if "ai_overview" in serp_data and "references" in serp_data["ai_overview"]:
        for i, ref in enumerate(serp_data["ai_overview"]["references"]):
            link = ref.get("link", "")
            if domain in link.lower():
                return i + 1
    return None

def get_serp_results(keyword, serp_api_key):
    params = {
        "engine": "google",
        "hl": "en",
        "gl": "us",
        "q": keyword,
        "api_key": serp_api_key,
    }
    response = requests.get("https://serpapi.com/search", params=params, headers=HEADERS)
    return response.json()

def extract_competitor_urls(serp_data):
    competitor_urls = []
    if "organic_results" in serp_data:
        for result in serp_data["organic_results"]:
            if "link" in result:
                competitor_urls.append(result["link"])
    return competitor_urls

def get_ai_overview_competitors(serp_data, competitor_key):
    competitor_directory = COMPETITOR_DIRECTORY.get(competitor_key, [])
    prioritized = []
    others = []
    
    if "ai_overview" in serp_data:
        ai_overview = serp_data["ai_overview"]
        if "references" in ai_overview:
            for ref in ai_overview["references"]:
                if "link" in ref:
                    trimmed_link = trim_url(ref["link"])
                    if any(comp.lower() in trimmed_link.lower() for comp in competitor_directory):
                        prioritized.append(trimmed_link)
                    else:
                        others.append(trimmed_link)
    
    return (prioritized + others)

def get_ai_overview_othersites(serp_data, site):
    sited = []
    
    if "ai_overview" in serp_data:
        ai_overview = serp_data["ai_overview"]
        if "references" in ai_overview:
            for ref in ai_overview["references"]:
                if "link" in ref:
                    trimmed_link = trim_url(ref["link"])
                    
                    # Check if site name appears in the domain of the link
                    if site.lower() in trimmed_link.lower():
                        sited.append(trimmed_link)
    
    return sited[:5]

def get_ai_overview_content(serp_data):
    content_lines = []
    
    if "ai_overview" in serp_data and "text_blocks" in serp_data["ai_overview"]:
        for block in serp_data["ai_overview"]["text_blocks"]:
            
            if block.get("type") == "paragraph":
                snippet = block.get("snippet", "").strip()
                if snippet:
                    content_lines.append(snippet)
            
            elif block.get("type") == "list":
                list_items = block.get("list", [])
                for item in list_items:
                    title = item.get("title", "").strip()
                    snippet = item.get("snippet", "").strip()
                    combined = f"{title} {snippet}" if title and snippet else title or snippet
                    if combined:
                        content_lines.append(combined)
                    
                    # Handle nested lists
                    if "list" in item:
                        for sub_item in item["list"]:
                            sub_snippet = sub_item.get("snippet", "").strip()
                            if sub_snippet:
                                content_lines.append(f"- {sub_snippet}")
    
    return "\n\n".join(content_lines)

def get_ai_overview_competitors_content(serp_response, domain):
    ai_overview = serp_response.get("ai_overview", {})
    competitors = []
    domain_competitors = COMPETITOR_DIRECTORY.get(domain, [])  # Fetch competitors for the domain
    
    if "references" in ai_overview:
        reference_map = {ref.get("source"): ref for ref in ai_overview["references"]}
        
        for competitor in domain_competitors:
            if competitor in reference_map:
                competitor_ref = reference_map[competitor]
                competitor_index = competitor_ref.get("index")
                competitor_link = competitor_ref.get("link", "")
                
                if "text_blocks" in ai_overview:
                    for block in ai_overview["text_blocks"]:
                        if "reference_indexes" in block and competitor_index in block["reference_indexes"]:
                            competitors.append({
                                "name": competitor,
                                "content": block.get("snippet", ""),
                                "source": competitor_link
                            })
    return competitors

def get_ai_overview_questions(serp_data):
    related_questions = serp_data.get("related_questions", [])
    ai_questions = []

    for question_data in related_questions:
        if question_data.get("title") == "AI Overview":
            ai_questions.append({
                "question": question_data.get("question"),
                "content": question_data.get("list", []),  # Some AI Overviews use a list format
                "link": question_data.get("link"),
            })

    return ai_questions

def flatten_schema(schema_item):
    if isinstance(schema_item, list):
        for item in schema_item:
            yield from flatten_schema(item)
    elif isinstance(schema_item, dict):
        if '@graph' in schema_item:
            yield from flatten_schema(schema_item['@graph'])
        yield schema_item

def schema_implemented(schema_data, schema_type):
    for item in flatten_schema(schema_data):
        atype = item.get("@type", "")
        if isinstance(atype, str) and atype.lower() == schema_type.lower():
            return True
        elif isinstance(atype, list):
            for t in atype:
                if t.lower() == schema_type.lower():
                    return True
    return False



def build_schema_table(schema_data, url):
    
    content = fetch_page_content(url)
    if content is None:
        content = ""
        
    results = []
    for label, stype in SCHEMA_CHECKLIST:
        if schema_implemented(schema_data, stype):
            results.append({"schema": label, "implemented": "Yes", "remarks": "-"})
        elif stype in content:
            results.append({"schema": label, "implemented": "Yes", "remarks": "-"})
        else:
            results.append({"schema": label, "implemented": "No", "remarks": "Need to be Implemented"})
    return results

def get_headers_and_images_in_range(soup):
    found_first_h1 = False
    reached_faq = False
    headers = []
    images = []
    for el in soup.find_all(["h1", "h2", "h3", "img"]):
        if el.name in ["h1", "h2", "h3"]:
            headers.append({"tag": el.name.upper(), "text": el.get_text(strip=True)})
            if el.name == "h1" and not found_first_h1:
                found_first_h1 = True
            if "faq" in el.get_text(strip=True).lower():
                reached_faq = True
        elif el.name == "img":
            src = el.get("src", "").lower()
            alt = el.get("alt", "").lower()
            if not src.startswith("http") or "icon" in src or "favicon" in src or "icon" in alt:
                continue
            if found_first_h1 and not reached_faq:
                images.append({"src": el.get("src", ""), "alt": el.get("alt", "")})
    return headers, images

# Function to check for embedded videos
def get_embedded_videos(url):
    session = HTMLSession()
    r = session.get(url)

    # Run .render() safely for Streamlit
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(r.html.arender(timeout=20, sleep=2))
    except Exception as e:
        print(f"Render failed: {e}")
        return []

    # Use BeautifulSoup for consistency with your existing code
    soup = BeautifulSoup(r.html.html, "html.parser")

    # Extract <iframe>, <embed>, and <video> elements
    videos = []
    for el in soup.find_all(["iframe", "embed", "video"]):
        src = el.get("src") or el.get("data-src") or el.get("poster")
        if src and any(domain in src for domain in ["youtube", "vimeo", "wistia"]):
            videos.append({
                "tag": el.name,
                "src": src
            })

    return videos

def search_youtube_video(keyword, domain, serp_api_key=None):
    yt_channel = YOUTUBE_CHANNEL.get(domain, domain)  # Use the mapped channel or fallback to domain
    query = f"https://www.youtube.com/{yt_channel} {keyword}"
    
    search_url = f"https://serpapi.com/search.json?engine=youtube&search_query={query}&api_key={serp_api_key}"
    
    try:
        response = requests.get(search_url)
        response.raise_for_status()
        data = response.json()

        # Extract the first video result
        if "video_results" in data and data["video_results"]:
            top_video = data["video_results"][0]
            return f"{top_video['title']}: {top_video['link']}"
        else:
            return "No relevant video found."
    
    except requests.RequestException as e:
        return f"Error fetching YouTube video: {e}"

def analyze_target_content(target_url, serp_data):
    response = requests.get(target_url, headers=HEADERS)
    if response.status_code != 200:
        st.error(f"Warning: Received status code {response.status_code} from {target_url}.")
        return {"headers": [{"tag": "", "text": f"{response.status_code} Forbidden"}],
                "missing_headers": [], "images": [], "schema_table": []}
    soup = BeautifulSoup(response.text, "html.parser")
    page_headers, images_in_range = get_headers_and_images_in_range(soup)
    videos_in_range = get_embedded_videos(target_url)
    schema_scripts = soup.find_all("script", type="application/ld+json")
    schema_data = []
    for script in schema_scripts:
        try:
            data = json.loads(script.string)
            schema_data.append(data)
        except Exception as e:
            st.error("Error parsing schema: " + str(e))
    schema_table = build_schema_table(schema_data, target_url)
    ai_overview_headers = extract_ai_overview_headers(serp_data)
    missing_headers = compare_headers(page_headers, ai_overview_headers)
    return {"headers": page_headers, "missing_headers": missing_headers,
            "images": images_in_range, "schema_table": schema_table, "videos": videos_in_range}

def get_social_results(keyword, site, limit_max=5, serp_api_key=None):
    query = f"site:{site} {keyword}"
    params = {"engine": "google", "q": query, "hl": "en", "gl": "us", "api_key": serp_api_key}
    response = requests.get("https://serpapi.com/search", params=params, headers=HEADERS)
    data = response.json()
    results = []
    if "organic_results" in data:
        for result in data["organic_results"]:
            results.append({"title": result.get("title", "No Title"), "link": result.get("link", "")})
            if len(results) >= limit_max:
                break
    return results

def rank_titles_by_semantic_similarity(primary_keyword, titles, threshold=0.75):
    model = SentenceTransformer('all-MiniLM-L6-v2')
    query_embedding = model.encode(primary_keyword, convert_to_tensor=True)
    title_embeddings = model.encode(titles, convert_to_tensor=True)
    cosine_scores = util.pytorch_cos_sim(query_embedding, title_embeddings)
    cosine_scores = cosine_scores.cpu().numpy().flatten()
    ranked_titles = [(titles[i], float(cosine_scores[i])) for i in np.argsort(cosine_scores)[::-1]]
    return [item for item in ranked_titles if item[1] > threshold]

def get_youtube_results(keyword, limit_max=5, serp_api_key=None):
    query = f"site:youtube.com {keyword}"
    params = {"engine": "google", "q": query, "hl": "en", "gl": "us", "api_key": serp_api_key}
    response = requests.get("https://serpapi.com/search", params=params, headers=HEADERS)
    data = response.json()
    results = []
    if "organic_results" in data:
        for result in data["organic_results"]:
            key_moments_raw = result.get("key_moments", None)
            if key_moments_raw and isinstance(key_moments_raw, list):
                key_moments = "\n".join([f"• {km.get('time', '')} - {km.get('title', '')}" for km in key_moments_raw])
            else:
                key_moments = "Key Moments not found for video."
            source_raw = result.get("source", "")
            source_processed = source_raw.split("·")[-1].strip() if "·" in source_raw else source_raw
            results.append({"title": result.get("title", "No Title"),
                            "link": result.get("link", ""),
                            "displayed_link": result.get("displayed_link", ""),
                            "source": source_processed,
                            "snippet": result.get("snippet", ""),
                            "key_moments": key_moments})
            if len(results) >= limit_max:
                break
    return results

def add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement("w:rPr")
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")  # Properly namespaced
    rPr.append(u)
    new_run.append(rPr)
    new_run_text = OxmlElement('w:t')
    new_run_text.text = text
    new_run.append(new_run_text)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink