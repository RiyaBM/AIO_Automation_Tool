import os
import tempfile
import streamlit as st
import pdfkit
from jinja2 import Template
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from dotenv import load_dotenv
import shutil
from utils import add_hyperlink

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

def generate_docx_report(data,domain, output_file = "aio_report.docx"):
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
    document.add_heading("Relevant 5 Pages from AI Overview Sources", level=2)
    if data.get("ai_overview_competitors"):
        for url in data.get("ai_overview_competitors"):
            p = document.add_paragraph(style="List Bullet")
            add_hyperlink(p, url, url)
    else:
        document.add_paragraph("No AI Overview Competitors found.")

    document.add_heading("Other Pages from AI Overview Sources", level=3)
    document.add_heading("Forbes:", level=4)
    if data.get("ai_overview_forbes"):
        for url in data.get("ai_overview_forbes"):
            p = document.add_paragraph(style="List Bullet")
            add_hyperlink(p, url, url)
    else:
        document.add_paragraph("No Forbes pages found.")
    document.add_heading("PCMag:", level=4)
    if data.get("ai_overview_pcmag"):
        for url in data.get("ai_overview_pcmag"):
            p = document.add_paragraph(style="List Bullet")
            add_hyperlink(p, url, url)
    else:
        document.add_paragraph("No PCMag pages found.")
    document.add_heading("YouTybe:", level=4)
    if data.get("ai_overview_youtube"):
        for url in data.get("ai_overview_youtube"):
            p = document.add_paragraph(style="List Bullet")
            add_hyperlink(p, url, url)
    else:
        document.add_paragraph("No YouTube videos found.")
    document.add_heading("Linkedin:", level=4)
    if data.get("ai_overview_linkedin"):
        for url in data.get("ai_overview_linkedin"):
            p = document.add_paragraph(style="List Bullet")
            add_hyperlink(p, url, url)
    else:
        document.add_paragraph("No LinkedIn pages found.")
    document.add_heading("Reddit:", level=4)
    if data.get("ai_overview_reddit"):
        for url in data.get("ai_overview_reddit"):
            p = document.add_paragraph(style="List Bullet")
            add_hyperlink(p, url, url)
    else:
        document.add_paragraph("No Reddit pages found.")

    p = document.add_paragraph()
    p.add_run("Number of AI Sources in Organic Search (first 20): ").bold = True
    p.add_run(str(data.get("ai_sources_in_organic_count", "")))

    # Competitors Section
    document.add_heading("Competitors Listed", level=2)

    if "competitors" in data:
        table = document.add_table(rows=1, cols=3)
        table.style = "Table Grid"
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Competitor"
        hdr_cells[1].text = "AI Content"
        hdr_cells[2].text = "Source"

        for competitor in data["competitors"]:
            row_cells = table.add_row().cells
            row_cells[0].text = competitor.get("name", "")
            content = competitor.get("content", "")
            if isinstance(content, list):
                row_cells[1].text = "\n".join(f"• {item}" for item in content)  # Format as bullet list
            else:
                row_cells[1].text = content  # Keep as-is if not a list
            row_cells[2].text = competitor.get("source", "")

    # PAA Section
    document.add_heading("People Also Ask", level=2)

    if "peopleAlsoAsk_ai_overview" in data:
        table = document.add_table(rows=1, cols=3)
        table.style = "Table Grid"
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Question"
        hdr_cells[1].text = "AI Content"
        hdr_cells[2].text = "Source"

        for ques in data["peopleAlsoAsk_ai_overview"]:
            row_cells = table.add_row().cells
            row_cells[0].text = ques.get("question", "")
            row_cells[1].text = ques.get("content", "")
            row_cells[2].text = ques.get("link", "")

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
    document.add_heading("Suggetions:", level=4)
    for line in ["Upload videos frequently.", "Write keyword-rich descriptions with timestamps and CTAs."]:
         document.add_paragraph(line, style="List Bullet")
    
    document.add_heading("Social Channels", level=3)
    
    if data.get("social_channels"):
        tbl = document.add_table(rows=1, cols=3)
        tbl.style = 'Table Grid'
        
        # Table headers
        hdr_cells = tbl.rows[0].cells
        hdr_cells[0].text = "Social Channel"
        hdr_cells[1].text = "Relevant Articles / Questions"
        hdr_cells[2].text = "Suggestions"

        # Populate table rows
        for channel in data["social_channels"]:
            row_cells = tbl.add_row().cells
            row_cells[0].text = channel.get("channel", "")
            
            # Process multiple hyperlinks correctly
            relevant_text = channel.get("relevant", "")
            p = row_cells[1].paragraphs[0]
            
            # Extracting links and titles properly
            if "<a href=" in relevant_text:
                import re
                links = re.findall(r"<a href='(.*?)' target='_blank'>(.*?)</a>", relevant_text)
                
                for idx, (url, title) in enumerate(links):
                    add_hyperlink(p, url, title)
                    if idx < len(links) - 1:
                        p.add_run("\n")  # Add line break between links
            else:
                p.add_run(relevant_text)  # If no links, just add plain text
            
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
    <html lang="en">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>AIO Analysis Report</title>
        <style>
            body {
                font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
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
        <p><strong>Keyword:</strong> {{ data.keyword }}</p>
        <p><strong>Target URL:</strong> <a href="{{ data.target_url }}">{{ data.target_url }}</a></p>
        <p><strong>{{data.domain}} Found in AI Overview Sources:</strong> {{ data.domain_found }}</p>

        <h2>{{data.domain}} Ranking</h2>
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
                    <td>{{ data.keyword }}</td>
                    <td>
                        {% if data.domain_organic_position %}
                            {{ data.domain_organic_position }}
                        {% else %}
                            Not Ranking
                        {% endif %}
                    </td>
                    <td>
                        {% if data.domain_ai_position %}
                            {{ data.domain_ai_position }}
                        {% else %}
                            Not Ranking
                        {% endif %}
                    </td>
                </tr>
            </tbody>
        </table>

        <h2>Competitors Listed By</h2>
        <table>
            <tr>
                <th>Competitor</th>
                <th>AI Content</th>
                <th>Source</th>
            </tr>
            {% for competitor in competitors %}
            <tr>
                <td>{{ competitor.name }}</td>
                <td>{{ competitor.content }}</td>
                <td><a href="{{ competitor.source }}">{{ competitor.source }}</a></td>
            </tr>
            {% endfor %}
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
