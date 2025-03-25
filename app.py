import os
import json
import difflib
import tempfile
import requests
import streamlit as st
import pdfkit
import numpy as np
from urllib.parse import urlparse
from bs4 import BeautifulSoup
from jinja2 import Environment, FileSystemLoader, Template
from docx import Document
from docx.shared import Inches, Pt
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from sentence_transformers import SentenceTransformer, util
from dotenv import load_dotenv
import shutil

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

# -------------------------------
# Utility and Analysis Functions
# -------------------------------

def is_similar(a, b, threshold=0.8):
    return difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio() >= threshold

def trim_url(url):
    return url.split('#')[0] if url else url

def extract_ai_overview_headers(serp_data):
    headers = []
    if "ai_overview" in serp_data and "text_blocks" in serp_data["ai_overview"]:
        for block in serp_data["ai_overview"]["text_blocks"]:
            if block.get("type") == "paragraph":
                snippet = block.get("snippet", "").strip()
                if snippet.endswith(":"):
                    headers.append(snippet.rstrip(":").strip())
    return headers

def compare_headers(page_headers, ai_overview_headers):
    missing = []
    page_header_texts = [h["text"] for h in page_headers]
    for ai_header in ai_overview_headers:
        if not any(is_similar(ai_header, ph) for ph in page_header_texts):
            missing.append(ai_header)
    return missing

def extract_domain(url):
    parsed_url = urlparse(url)
    domain_parts = parsed_url.netloc.split(".")
    return domain_parts[-2] if len(domain_parts) > 2 else domain_parts[0]

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

def get_ai_overview_competitors(serp_data):
    priority_sources = {
        "faxplus.com", "ifaxapp.com", "faxburner.com", "faxauthority.com",
        "pandadoc.com", "cocofax.com", "wisefax.com"
    }
    prioritized = []
    others = []
    if "ai_overview" in serp_data:
        ai_overview = serp_data["ai_overview"]
        if "references" in ai_overview:
            for ref in ai_overview["references"]:
                if "link" in ref:
                    trimmed_link = trim_url(ref["link"])
                    if any(source in trimmed_link for source in priority_sources):
                        prioritized.append(trimmed_link)
                    else:
                        others.append(trimmed_link)
    return (prioritized + others)[:5]

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
    return "\n\n".join(content_lines)

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

def build_schema_table(schema_data):
    SCHEMA_CHECKLIST = [
        ("Breadcrumbs", "BreadcrumbList"),
        ("FAQ", "FAQPage"),
        ("Article", "Article"),
        ("Video", "VideoObject"),
        ("Organization", "Organization"),
        ("How-to", "HowTo"),
    ]
    results = []
    for label, stype in SCHEMA_CHECKLIST:
        if schema_implemented(schema_data, stype):
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
            if "icon" in src or "favicon" in src or "icon" in alt:
                continue
            if found_first_h1 and not reached_faq:
                images.append({"src": el.get("src", ""), "alt": el.get("alt", "")})
    return headers, images

def analyze_target_content(target_url, serp_data):
    response = requests.get(target_url, headers=HEADERS)
    if response.status_code != 200:
        st.error(f"Warning: Received status code {response.status_code} from {target_url}.")
        return {"headers": [{"tag": "", "text": f"{response.status_code} Forbidden"}],
                "missing_headers": [], "images": [], "schema_table": []}
    soup = BeautifulSoup(response.text, "html.parser")
    page_headers, images_in_range = get_headers_and_images_in_range(soup)
    schema_scripts = soup.find_all("script", type="application/ld+json")
    schema_data = []
    for script in schema_scripts:
        try:
            data = json.loads(script.string)
            schema_data.append(data)
        except Exception as e:
            st.error("Error parsing schema: " + str(e))
    schema_table = build_schema_table(schema_data)
    ai_overview_headers = extract_ai_overview_headers(serp_data)
    missing_headers = compare_headers(page_headers, ai_overview_headers)
    return {"headers": page_headers, "missing_headers": missing_headers,
            "images": images_in_range, "schema_table": schema_table}

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
    new_run_text = OxmlElement('w:t')
    new_run_text.text = text
    new_run.append(new_run_text)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def generate_docx_report(data, domain, output_file="seo_report.docx"):
    document = Document()
    document.add_heading("SEO Analysis Report", level=1)
    p = document.add_paragraph()
    p.add_run("Keyword: ").bold = True
    p.add_run(str(data.get("keyword", "")))
    p = document.add_paragraph()
    p.add_run("Target URL: ").bold = True
    p_link = document.add_paragraph()
    add_hyperlink(p_link, data.get("target_url", ""), data.get("target_url", ""))
    p = document.add_paragraph()
    p.add_run(domain) 
    p.add_run(" Found in AI Overview Sources: ").bold = True
    p.add_run(str(data.get("domain_found", "")))
    document.add_heading("Domain Ranking", level=2)
    table = document.add_table(rows=2, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Keyword"
    hdr_cells[1].text = "Google Search"
    hdr_cells[2].text = "Google - AI Overview"
    row_cells = table.rows[1].cells
    row_cells[0].text = data.get("keyword", "")
    row_cells[1].text = str(data.get("domain_organic_position", "Not Ranking"))
    row_cells[2].text = str(data.get("domain_ai_position", "Not Ranking"))
    document.add_heading("AI Overview Content", level=2)
    ai_content = data.get("ai_overview_content", "")
    for line in ai_content.split("\n"):
        document.add_paragraph(line)
    document.add_heading("Top 4 Pages from AI Overview Sources", level=2)
    if data.get("ai_overview_competitors"):
        for url in data.get("ai_overview_competitors"):
            p = document.add_paragraph(style="List Bullet")
            add_hyperlink(p, url, url)
    else:
        document.add_paragraph("No AI Overview Competitors found.")
    p = document.add_paragraph()
    p.add_run("Number of AI Sources in Organic Search (first 20): ").bold = True
    p.add_run(str(data.get("ai_sources_in_organic_count", "")))
    document.add_heading("Content Analysis", level=2)
    document.add_heading("Headers", level=3)
    if data.get("content_analysis", {}).get("headers"):
        tbl = document.add_table(rows=1, cols=2)
        tbl.style = 'Table Grid'
        hdr_cells = tbl.rows[0].cells
        hdr_cells[0].text = "Header Tag"
        hdr_cells[1].text = "Header Text"
        for header in data["content_analysis"]["headers"]:
            row_cells = tbl.add_row().cells
            row_cells[0].text = header.get("tag", "")
            row_cells[1].text = header.get("text", "")
    else:
        document.add_paragraph("No headers found.")
    document.add_heading("Missing Headers (compared to AI Overview)", level=3)
    if data.get("content_analysis", {}).get("missing_headers"):
        for mh in data["content_analysis"]["missing_headers"]:
            document.add_paragraph(mh, style="List Bullet")
    else:
        document.add_paragraph("No missing headers compared to AI Overview.")
    document.add_heading("Images (After H1 and Before FAQ)", level=3)
    if data.get("content_analysis", {}).get("images"):
        tbl = document.add_table(rows=1, cols=2)
        tbl.style = 'Table Grid'
        hdr_cells = tbl.rows[0].cells
        hdr_cells[0].text = "Image Source"
        hdr_cells[1].text = "Alt Text"
        for image in data["content_analysis"]["images"]:
            row_cells = tbl.add_row().cells
            row_cells[0].text = image.get("src", "")
            row_cells[1].text = image.get("alt", "")
    else:
        document.add_paragraph("No images found.")
    document.add_heading("Schema Markup", level=3)
    if data.get("content_analysis", {}).get("schema_table"):
        tbl = document.add_table(rows=1, cols=3)
        tbl.style = 'Table Grid'
        hdr_cells = tbl.rows[0].cells
        hdr_cells[0].text = "Schema"
        hdr_cells[1].text = "Implemented"
        hdr_cells[2].text = "Remarks"
        for row in data["content_analysis"]["schema_table"]:
            row_cells = tbl.add_row().cells
            row_cells[0].text = row.get("schema", "")
            row_cells[1].text = row.get("implemented", "")
            row_cells[2].text = row.get("remarks", "")
    else:
        document.add_paragraph("No schema markup data found.")
    document.add_heading("Brand Mentionings", level=2)
    document.add_heading("YouTube", level=3)
    if data.get("youtube_results"):
        tbl = document.add_table(rows=1, cols=5)
        tbl.style = 'Table Grid'
        hdr_cells = tbl.rows[0].cells
        hdr_cells[0].text = "Title"
        hdr_cells[1].text = "Displayed Link"
        hdr_cells[2].text = "Source"
        hdr_cells[3].text = "Snippet"
        hdr_cells[4].text = "Key Moments"
        for yt in data["youtube_results"]:
            row_cells = tbl.add_row().cells
            p = row_cells[0].paragraphs[0]
            add_hyperlink(p, yt.get("link", ""), yt.get("title", ""))
            row_cells[1].text = yt.get("displayed_link", "")
            row_cells[2].text = yt.get("source", "")
            row_cells[3].text = yt.get("snippet", "")
            row_cells[4].text = yt.get("key_moments", "")
    else:
        document.add_paragraph("No YouTube results found.")
    document.add_heading("Social Channels", level=3)
    if data.get("social_channels"):
        tbl = document.add_table(rows=1, cols=3)
        tbl.style = 'Table Grid'
        hdr_cells = tbl.rows[0].cells
        hdr_cells[0].text = "Social Channel"
        hdr_cells[1].text = "Relevant Articles / Questions"
        hdr_cells[2].text = "Suggestions"
        for channel in data["social_channels"]:
            row_cells = tbl.add_row().cells
            row_cells[0].text = channel.get("channel", "")
            row_cells[1].text = channel.get("relevant", "")
            row_cells[2].text = channel.get("suggestions", "")
    else:
        document.add_paragraph("No social channels data found.")
    document.add_heading("Top SERP URLs", level=2)
    if data.get("competitor_urls"):
        for url in data["competitor_urls"]:
            p = document.add_paragraph(style="List Bullet")
            add_hyperlink(p, url, url)
    else:
        document.add_paragraph("No competitor URLs found.")
    document.save(output_file)
    st.success("DOCX report generated: " + output_file)

def generate_pdf_report(data):
    """
    Generate an HTML report from a template and convert it into a PDF.
    """
    HTML_TEMPLATE = """
    <!DOCTYPE html>
    <html lang=\"en\">
    <head>
        <meta charset=\"utf-8\">
        <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
        <title>AIO Analysis Report</title>
        <style>
            body {
                font-family: \"Segoe UI\", Tahoma, Geneva, Verdana, sans-serif;
                margin: 20px;
                color: #444;
                line-height: 1.6;
            }
            h1, h2, h3 {
                color: #333;
                margin-bottom: 10px;
            }
            h1 {
                border-bottom: 2px solid #333;
                padding-bottom: 5px;
            }
            p {
                margin: 10px 0;
            }
            a {
                color: #1a73e8;
                text-decoration: none;
            }
            a:hover {
                text-decoration: underline;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                margin: 20px 0;
            }
            th, td {
                border: 1px solid #ddd;
                padding: 10px;
                text-align: left;
            }
            th {
                background-color: #f5f5f5;
            }
            ul {
                margin: 10px 0 20px 20px;
            }
            .small-heading {
                margin-top: 40px;
            }
        </style>
    </head>
    <body>
        <h1>AI Overview Analysis Report</h1>
        <p><strong>Keyword:</strong> {{ keyword }}</p>
        <p><strong>Target URL:</strong> <a href=\"{{ target_url }}\">{{ target_url }}</a></p>
        <p><strong>{{ domain }} Found in AI Overview Sources:</strong> {{ domain_found }}</p>
        <h2>{{ domain }} Ranking</h2>
        <table>
            <thead>
                <tr>
                    <th>Keyword</th>
                    <th>Google Search</th>
                    <th>Google - AI Overview</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>{{ keyword }}</td>
                    <td>{{ organic_position if organic_position else 'Not Ranking within 5 Pages' }}</td>
                    <td>{{ ai_position if ai_position else 'Not Ranking' }}</td>
                </tr>
            </tbody>
        </table>
    </body>
    </html>
    """
    template = Template(HTML_TEMPLATE)
    html_report = template.render(**data)

    # Auto-detect wkhtmltopdf
    wkhtmltopdf_path = shutil.which("wkhtmltopdf") or r"C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe"

    # Verify the path
    if not shutil.which("wkhtmltopdf") and not os.path.exists(wkhtmltopdf_path):
        raise FileNotFoundError("wkhtmltopdf not found. Install from https://wkhtmltopdf.org/downloads.html")

    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as pdf_file:
        pdfkit.from_string(html_report, pdf_file.name, configuration=config)
        return pdf_file.name
    
# def generate_docx_report(data):
#     """
#     Generate a DOCX report.
#     """
#     document = Document()
#     document.add_heading("AI Overview Analysis Report", level=1)
#     document.add_paragraph(f"Keyword: {data['keyword']}")
#     document.add_paragraph(f"Target URL: {data['target_url']}")
#     document.add_paragraph(f"{data['domain']} Found in AI Overview Sources: {data['domain_found']}")
    
#     document.add_heading("Ranking", level=2)
#     table = document.add_table(rows=2, cols=3)
#     hdr_cells = table.rows[0].cells
#     hdr_cells[0].text = "Keyword"
#     hdr_cells[1].text = "Google Search"
#     hdr_cells[2].text = "Google - AI Overview"
#     row_cells = table.rows[1].cells
#     row_cells[0].text = data['keyword']
#     row_cells[1].text = str(data.get("organic_position", "Not Ranking"))
#     row_cells[2].text = str(data.get("ai_position", "Not Ranking"))
    
#     with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as docx_file:
#         document.save(docx_file.name)
#         return docx_file.name

# -------------------------------
# Streamlit UI Integration
# -------------------------------

st.set_page_config(page_title="SEO AI Analysis Report Generator", layout="wide")

st.title("SEO AI Analysis Report Generator")

# User inputs for analysis
with st.form("analysis_form"):
    st.markdown("### Enter the following details to run the full SEO analysis:")
    keyword = st.text_input("Keyword")
    target_url = st.text_input("Target URL")
    submitted = st.form_submit_button("Run Analysis")
    
if submitted:
    # Access the API key from Streamlit's secrets
    SERPAPI_KEY = st.secrets["SERPAPI_KEY"]
    if not SERPAPI_KEY:
        st.error("SERP API Key is required!")
    else:
        st.info("Fetching SERP data for keyword: " + keyword)
        serp_data = get_serp_results(keyword, SERPAPI_KEY)
        domain = extract_domain(target_url).lower()
        competitor_urls = extract_competitor_urls(serp_data)
        ai_overview_competitors = get_ai_overview_competitors(serp_data)
        domain_present = check_domain_in_ai_overview(serp_data, domain, target_url)
        domain_organic_position = find_domain_position_in_organic(serp_data, domain)
        domain_ai_position = find_domain_position_in_ai(serp_data, domain)
        competitor_urls_first20 = [trim_url(url) for url in competitor_urls[:20]]
        ai_sources_in_organic_count = sum(1 for source in ai_overview_competitors if source in competitor_urls_first20)
        ai_overview_content = get_ai_overview_content(serp_data)
        st.info("Analyzing target URL content...")
        content_data = analyze_target_content(target_url, serp_data)
        
        st.info("Fetching social results from LinkedIn and Reddit...")
        linkedin_results = get_social_results(keyword, "linkedin.com", limit_max=5, serp_api_key=SERPAPI_KEY)
        reddit_results = get_social_results(keyword, "reddit.com", limit_max=5, serp_api_key=SERPAPI_KEY)
        linkedin_titles = [r["title"] for r in linkedin_results]
        reddit_titles = [r["title"] for r in reddit_results]
        ranked_linkedin_titles = rank_titles_by_semantic_similarity(keyword, linkedin_titles, threshold=0.75)
        ranked_reddit_titles = rank_titles_by_semantic_similarity(keyword, reddit_titles, threshold=0.75)
        social_channels = [
            {
                "channel": "LinkedIn",
                "relevant": "<br><br>".join(
                    [f"<a href='{linkedin_results[i]['link']}' target='_blank'>{title}</a><br><small>{linkedin_results[i]['link']}</small>"
                     for i, (title, _) in enumerate(ranked_linkedin_titles)]
                ) if ranked_linkedin_titles else "No relevant LinkedIn discussions found.",
                "suggestions": "Create an official LinkedIn presence and engage in relevant discussions."
            },
            {
                "channel": "Reddit",
                "relevant": "<br><br>".join(
                    [f"<a href='{reddit_results[i]['link']}' target='_blank'>{title}</a><br><small>{reddit_results[i]['link']}</small>"
                     for i, (title, _) in enumerate(ranked_reddit_titles)]
                ) if ranked_reddit_titles else "No relevant Reddit discussions found.",
                "suggestions": "Participate in Reddit discussions to boost engagement."
            }
        ]
        youtube_results = get_youtube_results(keyword, limit_max=5, serp_api_key=SERPAPI_KEY)
        
        if domain == "efax":
            domain = "eFax"
        else:
            domain = domain.title()
    
        # Compile data for reports
        report_data = {
            "keyword": keyword,
            "domain": domain,
            "target_url": target_url,
            "competitor_urls": competitor_urls,
            "ai_overview_competitors": ai_overview_competitors,
            "domain_found": domain_present,
            "ai_sources_in_organic_count": ai_sources_in_organic_count,
            "ai_overview_content": ai_overview_content,
            "domain_organic_position": domain_organic_position,
            "domain_ai_position": domain_ai_position,
            "content_analysis": content_data,
            "social_channels": social_channels,
            "youtube_results": youtube_results,
            "ranked_linkedin_titles": ranked_linkedin_titles,
            "ranked_reddit_titles": ranked_reddit_titles
        }
        
        st.success("Analysis complete!")

        # Display key results in Streamlit
        st.markdown("#### AI Overview Content")
        st.text(ai_overview_content)

        st.info("Generating PDF Report...")
        pdf_path = generate_pdf_report(report_data)
        with open(pdf_path, "rb") as file:
            st.download_button("Download PDF Report", data=file, file_name="SEO_Report.pdf", mime="application/pdf")

        st.info("Generating DOCX Report")

        # Define output file path
        docx_output_file = "SEO_Report_"+keyword+".docx"

        # Generate DOCX report
        generate_docx_report(report_data, domain, output_file=docx_output_file)

        # Ensure the file exists before trying to download
        if os.path.exists(docx_output_file):
            with open(docx_output_file, "rb") as file:
                st.download_button(
                    label="Download DOCX Report",
                    data=file,
                    file_name=docx_output_file,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.error("Error generating DOCX report. Please try again.")
