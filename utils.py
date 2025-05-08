"""utils.py"""
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
import aiohttp
import os
from urllib.parse import urlparse, parse_qs
# Load environment variables from .env if present
# Extract JSON from the response
import re
import openai
from serpapi import GoogleSearch
            
load_dotenv()

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/123.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Connection": "keep-alive"
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
    "splashtop": ["TeamViewer", "AnyDesk", "ManageEngine", "BeyondTrust", "GoTo", "Splashtop"],
    "fortinet": ["paloaltonetworks", "cisco", "forcepoint", "hpe", "checkpoint"]
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
        response = requests.get(url, headers=HEADERS, timeout=10)
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

def is_domain_match(url, target_domain):
        st.info("TARGET::::")
        st.info(target_domain)
        if not url:
            return False
            
        parsed = urlparse(url.lower())
        url_domain = parsed.netloc.replace('www.', '')
        st.info("URL::::")
        st.info(url_domain)
        # Direct match
        if url_domain == target_domain:
            return True
            
        # Subdomain match
        if url_domain.endswith('.' + target_domain):
            return True
            
        # Check if domain is a significant part
        domain_parts = url_domain.split('.')
        if target_domain in domain_parts:
            return True
            
        return False

def find_domain_position_in_organic(serp_data, domain):
     # Normalize the domain
    domain = domain.lower().replace('www.', '')

    # Check the current page of results
    if "organic_results" in serp_data:
        for i, result in enumerate(serp_data["organic_results"]):
            link = result.get("link", "")
            if is_domain_match(link, domain):
                return i + 1  # Found the domain at position i+1
        # If not found in current results, return "> 50"
        return "> 50"
    # No organic results found
    return "Not Ranking"

def find_domain_position_in_ai(serp_data, domain):
    # Normalize the domain
    domain = domain.lower().replace('www.', '')

    if "ai_overview" in serp_data and "references" in serp_data["ai_overview"]:
        for i, ref in enumerate(serp_data["ai_overview"]["references"]):
            link = ref.get("link", "")
            if is_domain_match(link, domain):
                return i + 1
        return "Not Ranking"  # Return "Not Ranking" if domain not found in AI references
    return "AI Not Appearing"  # Return "AI Not Appearing" if no AI overview exists

def get_serp_results(keyword, serp_api_key):
    try:
        search = GoogleSearch({
            "q": keyword,
            "api_key": serp_api_key,
            "engine": "google",
            "hl": "en",
            "gl": "us"
        })
        
        result = search.get_dict()
        return result
    except Exception as e:
        st.error(f"SerpAPI search failed: {str(e)}")
        return {"organic_results": [], "pagination": {}}
    
def get_50serp_results(keyword, serp_api_key):
    try:
        search = GoogleSearch({
            "engine": "google",
            "hl": "en",
            "gl": "us",
            "q": keyword,
            "start": 0,
            "num": 50,
            "api_key": serp_api_key
        })
        
        result = search.get_dict()
        return result
    except Exception as e:
        st.error(f"SerpAPI search failed: {str(e)}")
        return {"organic_results": [], "pagination": {}}
    

def extract_competitor_urls(serp_data):
    competitor_urls = []
    if "organic_results" in serp_data:
        for result in serp_data["organic_results"]:
            if "link" in result:
                competitor_urls.append(result["link"])
    return competitor_urls

def get_ai_overview_competitors(serp_data, serp_data50, competitor_key):
    competitor_directory = COMPETITOR_DIRECTORY.get(competitor_key, [])
    prioritized = []
    others = []
    
    if "ai_overview" in serp_data:
        ai_overview = serp_data["ai_overview"]
        if "references" in ai_overview:
            ref_indexes = {ref.get("index"): i + 1 for i, ref in enumerate(ai_overview["references"])}
            
            for ref in ai_overview["references"]:
                if "link" in ref:
                    trimmed_link = trim_url(ref["link"])
                    position = find_domain_position_in_organic(serp_data50, trimmed_link.lower())
                    citation = ref_indexes.get(ref.get("index"))
                    entry = {
                        "url": trimmed_link,
                        "position": position,
                        "citation": citation
                    }
                    if any(comp.lower() in trimmed_link.lower() for comp in competitor_directory):
                        prioritized.append(entry)
                    else:
                        others.append(entry)
    
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

# Function to check for embedded videos - improved version that doesn't use requests_html
def get_embedded_videos(url):
    try:
        response = requests.get(url, headers=HEADERS, timeout=15)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        
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
    except requests.RequestException as e:
        print(f"Error fetching URL: {e}")
        return []

def search_youtube_video(keyword, domain, serp_api_key=None):
    """
    Search for YouTube videos from a specific channel related to a keyword
    using SerpAPI's official client.
    
    Args:
        keyword (str): The search term
        domain (str): The domain or channel identifier
        serp_api_key (str): SerpAPI key for authentication
        
    Returns:
        str: Title and link of the top result or error message
    """
    try:
        # Get the YouTube channel ID from the mapping or use domain as fallback
        yt_channel = YOUTUBE_CHANNEL.get(domain, domain)
        
        # Prepare the search query
        query = f"https://www.youtube.com/{yt_channel} {keyword}"
        
        # Set up the search parameters for YouTube search
        search_params = {
            "engine": "youtube",
            "search_query": query,
            "api_key": serp_api_key
        }
        
        # Create and execute the search
        search = GoogleSearch(search_params)
        data = search.get_dict()
        
        # Extract the first video result
        if "video_results" in data and data["video_results"]:
            top_video = data["video_results"][0]
            return f"{top_video['title']}: {top_video['link']}"
        else:
            return "No relevant video found."
            
    except Exception as e:
        return f"Error fetching YouTube video: {str(e)}"

   
def extract_full_page_content(soup):
    """
    Extract the full content of the page from H1 to FAQ (if present).
    
    Args:
        soup (BeautifulSoup): BeautifulSoup object of the page.
        
    Returns:
        str: The extracted page content as text.
    """
    # Find the first H1 as starting point
    h1 = soup.find('h1')
    if not h1:
        return "No H1 found on page."
    
    # Find FAQ section if present (often marked by h2/h3 with 'faq' in text)
    faq_section = None
    for header in soup.find_all(['h2', 'h3', 'h4']):
        if 'faq' in header.get_text().lower():
            faq_section = header
            break
    
    content_elements = []
    
    # If we found both H1 and FAQ section
    if h1 and faq_section:
        # Get all elements between H1 and FAQ
        current = h1
        while current and current != faq_section:
            if current.name:  # Skip NavigableString objects
                content_elements.append(current)
            current = current.next_element
        
        # Include the FAQ section and its content (common pattern of dt/dd)
        content_elements.append(faq_section)
        
        # Look for FAQ content - usually in a dl, ul, or div after the FAQ header
        faq_content = None
        current = faq_section.next_sibling
        
        # Look through next siblings to find likely FAQ content
        while current and not faq_content:
            if current.name in ['dl', 'ul', 'div', 'section']:
                # Check if this might be a FAQ container
                if current.find(['dt', 'li', 'h3', 'h4', 'strong']):
                    faq_content = current
                    break
            current = current.next_sibling

        if faq_content:
            content_elements.append(faq_content)
    
    # If no FAQ found, get all content after H1
    elif h1:
        # Start with H1
        content_elements.append(h1)
        
        # Get siblings of H1 (typically the main content follows H1)
        for sibling in h1.next_siblings:
            if sibling.name:  # Skip NavigableString objects
                content_elements.append(sibling)
    
    # Extract text from all elements
    page_content = ""
    
    for element in content_elements:
        # For headers, format them appropriately
        if element.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            page_content += f"\n{element.name.upper()}: {element.get_text(strip=True)}\n"
        
        # For paragraphs and other text elements
        elif element.name in ['p', 'div', 'span', 'li', 'dt', 'dd', 'section', 'article']:
            text = element.get_text(strip=True)
            if text:  # Only add non-empty text
                page_content += f"{text}\n\n"
    
    return page_content.strip()

def perform_content_gap_analysis(ai_overview_content, page_content, openai_api_key):
    # Create the prompt for OpenAI API
    prompt = f"""
    * **Objective:** Identify content gaps between the AIO summary and the content on the page, and categorize critical issues under "Definition," "Content Updation," and "Content Addition." Provide actionable suggestions for improvement and cite relevant references to align with the AIO summary.
    * **Categories to Analyze:**
    * **Definition:**
        * **Current Status:** Is the definition related to the topic available on the page? If yes, is it contextually aligned with the AIO summary?
        * **Suggested Status:**
            * If the definition is missing, suggest adding a concise, relevant definition.
            * If the definition is not contextually aligned, suggest revising it and provide citation from the AIO summary.
            * If the definition is fine, indicate it with "-".
    * **Content Updation:**
        * **Current Status:**
            * Identify which content is partially discussed compared to the AIO summary.
            * Identify if any part of the page is text-heavy or lacks readability (e.g., long paragraphs or dense information).
        * **Suggested Status:**
            * Suggest which existing sections need expanded content or updates.
            * Propose formatting changes such as breaking content into **numbered lists**, **bullet points**, etc., for improved readability.
            * If no content updates are needed, leave it as "-".
    * **Content Addition:**
        * **Current Status:** Identify critical headers or topics that are missing on the page.
        * **Suggested Status:**
            * Suggest adding new headers, sub-headers, or FAQ questions to cover missing topics.
            * Mention where exactly these additions should be placed on the page.
            * Provide citation references from AIO summary for content suggestions.
            * If no content updates are needed, leave it as "-".
    * **Expected Format:**
    * Provide a **table format** with three columns:
        * **Category** (Definition, Content Updation, Content Addition)
        * **Current Status** (Current state of the content on the page)
        * **Suggestions** (Improvement suggestions, including specific references for AIO summary alignment)

    Page Content: {page_content}
    AI overview summary for keywords: {ai_overview_content}

    Based on the above content, provide a content gap analysis in the format of a table with three rows (Definition, Content Updation, Content Addition) and three columns (Category, Current Status, Suggestions). Present it as a JSON object with the following structure:
    {{"results": [
    {{"category": "Definition", "current_status": "...", "suggestions": "..."}},
    {{"category": "Content Updation", "current_status": "...", "suggestions": "..."}},
    {{"category": "Content Addition", "current_status": "...", "suggestions": "..."}}
    ]}}
    """

    # Configure OpenAI client
    client = openai.OpenAI(api_key=openai_api_key)
    
    # Call OpenAI API
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are an SEO specialist analyzing content gaps."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=800
        )
        
        # Extract and parse the response
        response_text = response.choices[0].message.content.strip()

        # Try to find JSON pattern in the response
        json_pattern = r'({.*})'
        json_match = re.search(json_pattern, response_text, re.DOTALL)
        
        if json_match:
            try:
                # Parse the JSON
                analysis_data = json.loads(json_match.group(0))
                return analysis_data
            except json.JSONDecodeError:
                # If JSON parsing fails, return a formatted error
                return {
                    "results": [
                        {"category": "Definition", "current_status": "Error parsing API response", "suggestions": "Please try again."},
                        {"category": "Content Updation", "current_status": "Error parsing API response", "suggestions": "Please try again."},
                        {"category": "Content Addition", "current_status": "Error parsing API response", "suggestions": "Please try again."}
                    ]
                }
        else:
            # If no JSON pattern found, structure response anyway
            return {
                "results": [
                    {"category": "Definition", "current_status": "API response format error", "suggestions": "Please try again."},
                    {"category": "Content Updation", "current_status": "API response format error", "suggestions": "Please try again."},
                    {"category": "Content Addition", "current_status": "API response format error", "suggestions": "Please try again."}
                ]
            }
            
    except Exception as e:
        # Return error information
        return {
            "results": [
                {"category": "Definition", "current_status": f"API Error: {str(e)[:50]}...", "suggestions": "Please check your API key and try again."},
                {"category": "Content Updation", "current_status": f"API Error: {str(e)[:50]}...", "suggestions": "Please check your API key and try again."},
                {"category": "Content Addition", "current_status": f"API Error: {str(e)[:50]}...", "suggestions": "Please check your API key and try again."}
            ]
        }

def get_alt_text_suggestion(current_img, page_url):
    """Get alt text suggestion from OpenAI API using the updated API (1.0.0+)."""
    try:
        # Check if OPENAI_API_KEY is in secrets
        openai_api_key = st.secrets.get("OPENAI_API_KEY")
        if not openai_api_key:
            return "OpenAI API key not configured"
        
        import openai
        client = openai.OpenAI(api_key=openai_api_key)
        
        prompt = f"""create an alt text for this image in max 10 to 15 words that's in one line,
describing image,
tone should be objective,
don't use words like [person, man or woman],
alingn with the content and intent of the page.

page url: {page_url}
Image Source: {current_img}

Suggested alt text:"""

        # Using the new client-based API format
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are an SEO specialist helping optimize alt text for images."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=100,
            temperature=0.7
        )
        
        # New way to access the response content
        suggestion = response.choices[0].message.content.strip()
        return suggestion
    except Exception as e:
        print(f"Error getting alt text suggestion: {e}")
        return f"Error generating suggestion: {str(e)[:50]}..."

def analyze_target_content(target_url, serp_data):
    try:
        response = requests.get(target_url, headers=HEADERS, timeout=15)
        if response.status_code != 200:
            st.error(f"Warning: Received status code {response.status_code} from {target_url}.")
            return {"headers": [{"tag": "", "text": f"{response.status_code} Forbidden"}],
                    "missing_headers": [], "images": [], "schema_table": [], "videos": [], 
                    "full_content": f"Error: Status code {response.status_code}"}
        
        soup = BeautifulSoup(response.text, "html.parser")
        page_headers, images_in_range = get_headers_and_images_in_range(soup)
        
        # Extract full page content
        full_page_content = extract_full_page_content(soup)

        # The rest of the function remains the same...
        if images_in_range:
            for image in images_in_range:
                suggested_alt = get_alt_text_suggestion(
                    image.get("src", ""),
                    target_url
                )
                image["suggested_alt"] = suggested_alt
                
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
                "images": images_in_range, "schema_table": schema_table, "videos": videos_in_range,
                "full_content": full_page_content}  # Add full content to return value
    except Exception as e:
        st.error(f"Error analyzing target content: {str(e)}")
        return {
            "headers": [], 
            "missing_headers": [],
            "images": [], 
            "schema_table": [], 
            "videos": [],
            "full_content": f"Error extracting content: {str(e)}"
        }

def analyze_secondary_content(page_headers, serp_data):
    ai_overview_headers = extract_ai_overview_headers(serp_data)
    missing_headers = compare_headers(page_headers, ai_overview_headers)
    return {"missing_headers": missing_headers}

def get_competitors_content(competitors):
    competitor_content = {}

    for competitor in competitors:
        url = competitor.get("url")
        name = extract_domain(url).lower()

        try:
            videos = get_embedded_videos(url)
            response = requests.get(url, headers=HEADERS, timeout=15)

            if response.status_code != 200:
                st.error(f"Warning: Received status code {response.status_code} from {url}.")
                competitor_content[name] = {
                    "headers": [{"tag": "", "text": f"{response.status_code} Forbidden"}],
                    "missing_headers": [],
                    "images": [],
                    "videos": [],
                    "schema_table": []
                }
                continue  # Skip to the next competitor

            soup = BeautifulSoup(response.text, "html.parser")
            headers, images = get_headers_and_images_in_range(soup)

            schema_scripts = soup.find_all("script", type="application/ld+json")
            schema_data = []
            for script in schema_scripts:
                try:
                    data = json.loads(script.string)
                    schema_data.append(data)
                except Exception as e:
                    st.error("Error parsing schema: " + str(e))
            schema_table = build_schema_table(schema_data, url)

            competitor_content[name] = {
                "headers": headers,
                "images": images,
                "videos": videos,
                "schema_table": schema_table
            }
        except Exception as e:
            st.error(f"Error processing competitor {name}: {str(e)}")
            competitor_content[name] = {
                "headers": [],
                "images": [],
                "videos": [],
                "schema_table": []
            }

    return competitor_content

def get_social_results(keyword, site, limit_max=5, serp_api_key=None):
    """
    Get search results from a specific site using SerpAPI's GoogleSearch client.
    
    Args:
        keyword (str): The search term
        site (str): The site to search within (e.g., "twitter.com")
        limit_max (int): Maximum number of results to return
        serp_api_key (str): SerpAPI key for authentication
        
    Returns:
        list: List of dictionaries with title and link of search results
    """
    # Construct the search query with site: operator
    query = f"site:{site} {keyword}"
    
    # Set up the search parameters
    search_params = {
        "engine": "google",
        "q": query,
        "hl": "en",
        "gl": "us",
        "api_key": serp_api_key
    }
    
    try:
        # Create and execute the search
        search = GoogleSearch(search_params)
        data = search.get_dict()
        
        # Extract and format the results
        results = []
        if "organic_results" in data:
            for result in data["organic_results"]:
                results.append({
                    "title": result.get("title", "No Title"),
                    "link": result.get("link", "")
                })
                if len(results) >= limit_max:
                    break
        
        return results
    except Exception as e:
        # Handle any exceptions that might occur
        print(f"Error with SerpAPI search: {str(e)}")
        return []

def rank_titles_by_semantic_similarity(primary_keyword, titles, threshold=0.75):
    """
    Rank titles by their semantic similarity to a primary keyword.
    
    Args:
        primary_keyword (str): The primary keyword to compare against.
        titles (list): List of titles to rank.
        threshold (float, optional): Minimum similarity score to include in results. Defaults to 0.75.
        
    Returns:
        list: List of tuples (title, similarity_score) sorted by similarity in descending order.
    """
    if not titles:
        return []
    
    try:
        # Load a smaller, faster model that doesn't require a Hugging Face token
        model = SentenceTransformer('paraphrase-MiniLM-L3-v2')
        
        # Encode the primary keyword and titles
        query_embedding = model.encode(primary_keyword, convert_to_tensor=True)
        title_embeddings = model.encode(titles, convert_to_tensor=True)
        
        # Calculate cosine similarities
        cosine_scores = util.pytorch_cos_sim(query_embedding, title_embeddings)
        cosine_scores = cosine_scores.cpu().numpy().flatten()
        
        # Rank titles by similarity score
        ranked_titles = [(titles[i], float(cosine_scores[i])) for i in np.argsort(cosine_scores)[::-1]]
        
        # Filter by threshold
        return [item for item in ranked_titles if item[1] > threshold]
    
    except Exception as e:
        st.error(f"Error calculating semantic similarity: {str(e)}")
        # Fall back to returning titles without ranking if something goes wrong
        return [(title, 0.8) for title in titles]

def get_youtube_results(keyword, limit_max=5, serp_api_key=None):
    """
    Get YouTube search results using SerpAPI's GoogleSearch client.
    
    Args:
        keyword (str): The search term
        limit_max (int): Maximum number of results to return
        serp_api_key (str): SerpAPI key for authentication
        
    Returns:
        list: List of dictionaries with YouTube video details
    """
    try:
        # Create the query with site: operator for YouTube
        query = f"site:youtube.com {keyword}"
        
        # Set up search parameters
        search_params = {
            "engine": "google",
            "q": query,
            "hl": "en",
            "gl": "us",
            "api_key": serp_api_key
        }
        
        # Execute the search using the official client
        search = GoogleSearch(search_params)
        data = search.get_dict()
        
        results = []
        if "organic_results" in data:
            for result in data["organic_results"]:
                # Process key moments if available
                key_moments_raw = result.get("key_moments", None)
                if key_moments_raw and isinstance(key_moments_raw, list):
                    key_moments = "\n".join([f"• {km.get('time', '')} - {km.get('title', '')}" for km in key_moments_raw])
                else:
                    key_moments = "Key Moments not found for video."
                
                # Process source information
                source_raw = result.get("source", "")
                source_processed = source_raw.split("·")[-1].strip() if "·" in source_raw else source_raw
                
                # Compile the result
                results.append({
                    "title": result.get("title", "No Title"),
                    "link": result.get("link", ""),
                    "displayed_link": result.get("displayed_link", ""),
                    "source": source_processed,
                    "snippet": result.get("snippet", ""),
                    "key_moments": key_moments
                })
                
                # Limit the number of results
                if len(results) >= limit_max:
                    break
                    
        return results
        
    except Exception as e:
        print(f"Error fetching YouTube search results: {str(e)}")
        return []

def add_hyperlink(paragraph, url, text):
    """
    Add a hyperlink to a paragraph.
    
    Args:
        paragraph: The paragraph object to add the hyperlink to.
        url (str): The URL for the hyperlink.
        text (str): The display text for the hyperlink.
        
    Returns:
        The hyperlink object.
    """
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

def process_links_for_template(text):
    """
    Process hyperlinks in text for template rendering.
    This extracts all HTML links from a string and returns a list of (url, text) tuples.
    
    Args:
        text (str): Text containing HTML links.
        
    Returns:
        list: List of (url, text) tuples for each link in the text.
    """
    if not text or "<a href=" not in text:
        return []
        
    import re
    links = re.findall(r"<a href='(.*?)' target='_blank'>(.*?)</a>", text)
    return links