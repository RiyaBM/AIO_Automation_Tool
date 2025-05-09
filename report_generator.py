"""report_generator.py"""
import os
import tempfile
import streamlit as st
import pdfkit
from jinja2 import Template, Environment, BaseLoader
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from dotenv import load_dotenv
import shutil
import re
from utils import add_hyperlink, process_links_for_template

# Load environment variables from .env if present
load_dotenv()

# Constants
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

def generate_docx_report(data, domain, output_file="aio_report.docx"):
    """
    Generate a Word document report based on the provided SEO analysis data.
    
    Args:
        data (dict): SEO analysis data.
        domain (str): Domain name.
        output_file (str, optional): Output file name. Defaults to "aio_report.docx".
        
    Returns:
        None
    """
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
    # Create the table with header
    table = document.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Keyword"
    hdr_cells[1].text = "Google Search"
    hdr_cells[2].text = "Google - AI Overview"

    # Add row for the primary keyword
    row_cells = table.add_row().cells
    row_cells[0].text = data.get("keyword", "")
    row_cells[1].text = str(data.get("domain_organic_position", "Not Ranking"))
    row_cells[2].text = str(data.get("domain_ai_position", "Not Ranking"))

    # Add rows for secondary keywords
    for kw in data.get("secondary_keywords", []):
        row_cells = table.add_row().cells
        row_cells[0].text = kw
        row_cells[1].text = str(data.get("domain_organic_position_secondary", {}).get(kw, "Not Ranking"))
        row_cells[2].text = str(data.get("domain_ai_position_secondary", {}).get(kw, "Not Ranking"))
    
    document.add_heading("All AI Overview Sources", level=2)
    
    # Get primary keyword AI Overview competitors
    ai_competitors = data.get("ai_overview_competitors", [])
    
    # Get AI Overview competitors from secondary keywords, if available
    secondary_ai_competitors = data.get("secondary_ai_overview_competitors", {})
    
    # Combine all AI Overview competitors
    all_ai_competitors = ai_competitors.copy()
    
    # Add sources from secondary keywords with keyword labeling
    for keyword, competitors in secondary_ai_competitors.items():
        for competitor in competitors:
            # Add keyword info to the competitor
            competitor_with_keyword = competitor.copy()
            competitor_with_keyword["keyword"] = keyword
            all_ai_competitors.append(competitor_with_keyword)
    
    # Remove duplicates based on URL
    unique_urls = set()
    unique_ai_competitors = []
    
    for competitor in all_ai_competitors:
        url = competitor.get("url")
        if url and url not in unique_urls:
            unique_urls.add(url)
            unique_ai_competitors.append(competitor)
    
    if unique_ai_competitors:
        # Create a table with 3 columns to include keyword information
        table = document.add_table(rows=1, cols=4)
        table.style = 'Light List'  # or use 'Table Grid' or any Word style you prefer

        # Header row
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'URL'
        hdr_cells[1].text = 'SERP Position'
        hdr_cells[2].text = 'AI Citation'
        hdr_cells[3].text = 'Source Keyword'

        for item in unique_ai_competitors:
            row_cells = table.add_row().cells
            url = item.get("url")
            # Use actual position if available, otherwise "> 50"
            position = item.get("position")
            position_text = str(position) if position is not None else "> 50"
            
            # Use actual citation if available, otherwise "Not Ranking"
            citation = item.get("citation")
            citation_text = str(citation) if citation is not None else "Not Ranking"
            
            keyword_source = item.get("keyword", data.get("keyword", "Primary"))  # Default to "Primary" for the main keyword

            # Add hyperlink to the first cell
            p = row_cells[0].paragraphs[0]
            add_hyperlink(p, url, url)

            # Add position to second cell
            row_cells[1].text = position_text
            
            # Add citation to third cell
            row_cells[2].text = citation_text
            # Add source keyword to fourth cell
            row_cells[3].text = keyword_source
    else:
        document.add_paragraph("No AI Overview Competitors found.")
    
    # Rest of the code for other sections
    document.add_heading("Other Pages from AI Overview Sources", level=3)

    if data.get("social_ai_overview_sites") and any(data["social_ai_overview_sites"].values()):
        document.add_heading("Social Sites:", level=4)
        table = document.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Site"
        hdr_cells[1].text = "URL"
        
        for site, urls in data.get("social_ai_overview_sites").items():
            for url in urls:
                row_cells = table.add_row().cells
                row_cells[0].text = site.capitalize()
                row_cells[1].text = f'\u2022 {url}'
    else:
        document.add_heading("No Social sites found on AI Overview", level=4)

    if data.get("popular_ai_overview_sites") and any(data["popular_ai_overview_sites"].values()):
        document.add_heading("3rd Party Popular Sites:", level=4)
        table = document.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Site"
        hdr_cells[1].text = "URL"
        
        for site, urls in data.get("popular_ai_overview_sites").items():
            for url in urls:
                row_cells = table.add_row().cells
                row_cells[0].text = site.capitalize()
                row_cells[1].text = f'\u2022 {url}'
    else:
        document.add_heading("No Popular sites found on AI Overview", level=4)

    if data.get("review_ai_overview_sites") and any(data["review_ai_overview_sites"].values()):
        document.add_heading("Review Sites:", level=4)
        table = document.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Site"
        hdr_cells[1].text = "URL"
        
        for site, urls in data.get("review_ai_overview_sites").items():
            for url in urls:
                row_cells = table.add_row().cells
                row_cells[0].text = site.capitalize()
                row_cells[1].text = f'\u2022 {url}'
    else:
        document.add_heading("No Review sites found on AI Overview", level=4)

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
    # document.add_heading("People Also Ask", level=2)

    # if "peopleAlsoAsk_ai_overview" in data:
    #     table = document.add_table(rows=1, cols=3)
    #     table.style = "Table Grid"
    #     hdr_cells = table.rows[0].cells
    #     hdr_cells[0].text = "Question"
    #     hdr_cells[1].text = "AI Content"
    #     hdr_cells[2].text = "Source"

    #     for ques in data["peopleAlsoAsk_ai_overview"]:
    #         row_cells = table.add_row().cells
    #         row_cells[0].text = ques.get("question", "")
    #         row_cells[1].text = ques.get("content", "")
    #         row_cells[2].text = ques.get("link", "")

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
    
    # Content Gap Analysis section - REPLACE Missing Headers section
    document.add_heading("Content Gap Analysis", level=3)
    
    # Check if we have gap analysis data
    if data.get("content_gap_analysis") and data["content_gap_analysis"].get("results"):
        # Create a table for content gap analysis
        gap_table = document.add_table(rows=1, cols=3)
        gap_table.style = 'Table Grid'
        
        # Add headers
        hdr_cells = gap_table.rows[0].cells
        hdr_cells[0].text = "Category"
        hdr_cells[1].text = "Current Status"
        hdr_cells[2].text = "Suggestions"
        
        # Add data rows
        for result in data["content_gap_analysis"]["results"]:
            row_cells = gap_table.add_row().cells
            row_cells[0].text = result.get("category", "")
            row_cells[1].text = result.get("current_status", "")
            row_cells[2].text = result.get("suggestions", "")
    else:
        document.add_paragraph("No content gap analysis available.")
    
    document.add_heading("Missing Headers (compared to AI Overview)", level=3)
    if data.get("content_analysis", {}).get("missing_headers"):
        document.add_paragraph(f"Missing Headers for keyword:: {data.get('keyword', 'Main Keyword')}:", style="List Bullet")
        for mh in data["content_analysis"]["missing_headers"]:
            document.add_paragraph(mh, style="List Bullet 2")
    else:
        document.add_paragraph(f"No missing headers compared to AI Overview for Target keyword: {data.get('keyword', 'Main Keyword')}.")
    
    # Secondary keywords missing headers
    secondary_keywords = data.get("secondary_keywords", [])
    content_data_secondary = data.get("content_data_secondary", {})

    for kw in secondary_keywords:
        missing = content_data_secondary.get(kw, {}).get("missing_headers", [])
        if missing:
            document.add_paragraph(f"Missing Headers for keyword:: {kw}:", style="List Bullet")
            for mh in missing:
                document.add_paragraph(mh, style="List Bullet 2")
        else:
            document.add_paragraph(f"No missing headers compared to AI Overview for secondary keyword: {kw}.")
    
    document.add_heading("Images (After H1 and Before FAQ)", level=3)
    if data.get("content_analysis", {}).get("images"):
        tbl = document.add_table(rows=1, cols=3)  # Changed from 2 to 3 columns
        tbl.style = 'Table Grid'
        hdr_cells = tbl.rows[0].cells
        hdr_cells[0].text = "Image Source"
        hdr_cells[1].text = "Alt Text"
        hdr_cells[2].text = "Suggested Alt Text"  # New column
        
        for image in data["content_analysis"]["images"]:
            row_cells = tbl.add_row().cells
            row_cells[0].text = image.get("src", "")
            row_cells[1].text = image.get("alt", "")
            
            # Get suggested alt text if it's in the data
            if "suggested_alt" in image:
                row_cells[2].text = image.get("suggested_alt", "")
            else:
                row_cells[2].text = "No suggestion available"
    else:
        document.add_paragraph("No images found.")

    document.add_heading("Videos Embedded on Page", level = 3)
    if data.get("content_analysis", {}).get("videos"):
        tbl = document.add_table(rows=1, cols=3)
        tbl.style = 'Table Grid'
        hdr_cells = tbl.rows[0].cells
        hdr_cells[0].text = "Video"
        hdr_cells[1].text = "Source"
        hdr_cells[2].text = "Remarks"
        for row in data["content_analysis"]["videos"]:
            row_cells = tbl.add_row().cells
            row_cells[0].text = row.get("tag", "")
            row_cells[1].text = row.get("src", "")
            row_cells[2].text = "It is good practice to add timestamps for key moments in the description."
    else:
        document.add_paragraph("No video found on page.")
    
    document.add_heading("Relevant Videos on Youtube Channel", level = 3)
    if data.get("relevant_video"):
        tbl = document.add_table(rows=1, cols=2)
        tbl.style = 'Table Grid'
        hdr_cells = tbl.rows[0].cells
        hdr_cells[0].text = "Video"
        hdr_cells[1].text = "Suggestion"
        
        # Check if relevant_video is a list or a single item.
        relevant_video = data["relevant_video"]
        if isinstance(relevant_video, list):
            items = relevant_video
        else:
            # Wrap the non-list value in a list.
            items = [relevant_video]
        
        for item in items:
            row_cells = tbl.add_row().cells
            row_cells[0].text = str(item)
            row_cells[1].text = "Embed this most relevant video on page."
    else:
        document.add_paragraph("There is no relevant video on official channel.")
        document.add_paragraph("Create a video for this topic.")

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
    document.add_heading("Brand Mentions", level=2)
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
    document.add_heading("Suggestions:", level=4)
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
            p.style = "List Bullet"
            
            # Extracting links and titles properly
            if "<a href=" in relevant_text:
                import re
                links = re.findall(r"<a href='(.*?)' target='_blank'>(.*?)</a>", relevant_text)
                
                for idx, (url, title) in enumerate(links):
                    add_hyperlink(p, url, title)
                    if idx < len(links) - 1:
                        p.add_run("\n\n")  # Add line break between links
            else:
                p.add_run(relevant_text)  # If no links, just add plain text
            
            row_cells[2].text = channel.get("suggestions", "")
    else:
        document.add_paragraph("No social channels data found.")

    document.add_heading("AI Overview Competitors Content Analysis", level=2)

    if data.get("aio_competitor_content"):
        for source, content in data["aio_competitor_content"].items():
            document.add_heading(source, level=3)

            # Images
            document.add_heading("Images", level=4)
            images = content.get("images", [])
            if images:
                table = document.add_table(rows=1, cols=2)
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'Alt Text'
                hdr_cells[1].text = 'Image URL'

                for img in images:
                    row_cells = table.add_row().cells
                    row_cells[0].text = img.get("alt", "")
                    row_cells[1].text = img.get("src", "")
            else:
                document.add_paragraph("No images found.", style="BodyText")

            # Videos
            document.add_heading("Videos", level=4)
            videos = content.get("videos", [])
            if videos:
                table = document.add_table(rows=1, cols=2)
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'Tag'
                hdr_cells[1].text = 'Video Source'

                for video in videos:
                    row_cells = table.add_row().cells
                    row_cells[0].text = video.get("tag", "").upper()
                    row_cells[1].text = video.get("src", "")
            else:
                document.add_paragraph("No videos found.", style="BodyText")

            # Schema Table
            document.add_heading("Schema Table", level=4)
            schema_table = content.get("schema_table", [])
            if schema_table:
                table = document.add_table(rows=1, cols=2)
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'Schema'
                hdr_cells[1].text = 'Implemented'

                for row in schema_table:
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(row.get('schema', ''))
                    row_cells[1].text = str(row.get('implemented', ''))
            else:
                document.add_paragraph("No schema data found.", style="BodyText")
    else:
        document.add_paragraph("No citations data found.")

    document.add_heading(f'AI Overview Content for: {data.get("keyword")}', level=2)
    ai_content = data.get("ai_overview_content", "")
    for line in ai_content.split("\n"):
        document.add_paragraph(line)

    secondary_ai_content = data.get("secondary_ai_overview_content")

    for kw in secondary_keywords:
        document.add_heading(f"\nAI Overview Content for: {kw}", level=2)
        ai_content = secondary_ai_content[kw]
        for line in ai_content.split("\n"):
            document.add_paragraph(line)

            
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
    
    Args:
        data (dict): SEO analysis data dictionary.
        
    Returns:
        str: Path to the generated PDF file.
    """
    # Custom filter to process links in the template
    def urlize_links(text):
        links = process_links_for_template(text)
        return links or [("", text)]
        
    # Set up Jinja2 environment with the custom filter
    env = Environment(loader=BaseLoader())
    env.filters['urlize_links'] = urlize_links

    HTML_TEMPLATE = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
    <meta charset="UTF-8">
    <title>SEO Analysis Report</title>
    <style>
        /* --- Page Setup --- */
        @page { margin: 1in; }
        body { font-family: Arial, sans-serif; color: #333; line-height: 1.4; }

        /* --- Headings --- */
        h1 { font-size: 24px; margin-bottom: 0.5em; }
        h2 { font-size: 20px; margin-top: 1.5em; margin-bottom: 0.5em; }
        h3 { font-size: 16px; margin-top: 1em; margin-bottom: 0.4em; }
        h4 { font-size: 14px; margin-top: 0.8em; margin-bottom: 0.3em; }

        /* --- Paragraphs --- */
        p { margin: 0.4em 0; }
        p.small { font-size: 0.9em; color: #666; }

        /* --- Links --- */
        a { color: #1a0dab; text-decoration: none; }
        a:hover { text-decoration: underline; }

        /* --- Tables --- */
        table { width: 100%; border-collapse: collapse; margin: 0.5em 0; }
        th, td { border: 1px solid #ccc; padding: 0.4em; text-align: left; vertical-align: top; }
        th { background-color: #f2f2f2; }

        /* --- Lists --- */
        ul, ol { margin: 0.4em 0 0.4em 1.5em; }

        /* --- Page Breaks --- */
        .page-break { page-break-before: always; }
    </style>
    </head>
    <body>

    <!-- Title -->
    <h1>SEO Analysis Report</h1>

    <!-- Basic Info -->
    <p><strong>Keyword:</strong> {{ data.keyword }}</p>
    <p><strong>Target URL:</strong> <a href="{{ data.target_url }}">{{ data.target_url }}</a></p>
    <p><strong>{{ domain }}</strong> Found in AI Overview Sources: {{ data.domain_found }}</p>

    <!-- Domain Ranking -->
    <h2>Domain Ranking</h2>
    <table>
        <thead>
        <tr>
            <th>Keyword</th>
            <th>Google Search</th>
            <th>Google – AI Overview</th>
        </tr>
        </thead>
        <tbody>
        <!-- Primary -->
        <tr>
            <td>{{ data.keyword }}</td>
            <td>{{ data.domain_organic_position or "Not Ranking" }}</td>
            <td>{{ data.domain_ai_position or "Not Ranking" }}</td>
        </tr>
        <!-- Secondary -->
        {% for kw in data.secondary_keywords %}
        <tr>
            <td>{{ kw }}</td>
            <td>{{ data.domain_organic_position_secondary.get(kw, "Not Ranking") }}</td>
            <td>{{ data.domain_ai_position_secondary.get(kw, "Not Ranking") }}</td>
        </tr>
        {% endfor %}
        </tbody>
    </table>

    <!-- AI Overview Content -->
    <h2>AI Overview Content</h2>
    {% for line in data.ai_overview_content.split('\n') %}
        <p>{{ line }}</p>
    {% endfor %}

    <!-- All AI Overview Sources -->
    <h2>All AI Overview Sources</h2>
    {% set unique_urls = [] %}
    {% set all_competitors = [] %}

    {# First collect all unique competitors #}
    {% for item in data.ai_overview_competitors %}
        {% set _ = all_competitors.append({'url': item.url, 'position': item.position, 'keyword': 'Primary'}) %}
    {% endfor %}

    {# Add competitors from secondary keywords #}
    {% for kw, competitors in data.secondary_ai_overview_competitors.items() %}
        {% for item in competitors %}
            {% set _ = all_competitors.append({'url': item.url, 'position': item.position, 'keyword': kw}) %}
        {% endfor %}
    {% endfor %}

    {# Display the combined unique list #}
    {% if all_competitors %}
        <table>
        <thead>
            <tr><th>URL</th><th>SERP Position</th><th>Source Keyword</th></tr>
        </thead>
        <tbody>
            {% for item in all_competitors %}
                {% if item.url not in unique_urls %}
                    {% set _ = unique_urls.append(item.url) %}
                    <tr>
                    <td><a href="{{ item.url }}">{{ item.url }}</a></td>
                    <td>{{ item.position or "> 50" }}</td>
                    <td>{{ item.keyword }}</td>
                    </tr>
                {% endif %}
            {% endfor %}
        </tbody>
        </table>
    {% else %}
        <p>No AI Overview Competitors found.</p>
    {% endif %}

    <!-- Other Pages -->
    <h3>Other Pages from AI Overview Sources</h3>

    <!-- Social Sites -->
    {% if data.social_ai_overview_sites and data.social_ai_overview_sites.values()|selectattr("length", ">", 0)|list %}
        <h4>Social Sites:</h4>
        <table>
        <thead><tr><th>Site</th><th>URL</th></tr></thead>
        <tbody>
            {% for site, urls in data.social_ai_overview_sites.items() %}
            {% if urls %}
            {% for url in urls %}
            <tr><td>{{ site|capitalize }}</td><td>{{ url }}</td></tr>
            {% endfor %}
            {% endif %}
            {% endfor %}
        </tbody>
        </table>
    {% else %}
        <h4>No Social sites found on AI Overview</h4>
    {% endif %}

    <!-- Popular Sites -->
    {% if data.popular_ai_overview_sites and data.popular_ai_overview_sites.values()|selectattr("length", ">", 0)|list %}
        <h4>3rd Party Popular Sites:</h4>
        <table>
        <thead><tr><th>Site</th><th>URL</th></tr></thead>
        <tbody>
            {% for site, urls in data.popular_ai_overview_sites.items() %}
            {% if urls %}
            {% for url in urls %}
            <tr><td>{{ site|capitalize }}</td><td>{{ url }}</td></tr>
            {% endfor %}
            {% endif %}
            {% endfor %}
        </tbody>
        </table>
    {% else %}
        <h4>No Popular sites found on AI Overview</h4>
    {% endif %}

    <!-- Review Sites -->
    {% if data.review_ai_overview_sites and data.review_ai_overview_sites.values()|selectattr("length", ">", 0)|list %}
        <h4>Review Sites:</h4>
        <table>
        <thead><tr><th>Site</th><th>URL</th></tr></thead>
        <tbody>
            {% for site, urls in data.review_ai_overview_sites.items() %}
            {% if urls %}
            {% for url in urls %}
            <tr><td>{{ site|capitalize }}</td><td>{{ url }}</td></tr>
            {% endfor %}
            {% endif %}
            {% endfor %}
        </tbody>
        </table>
    {% else %}
        <h4>No Review sites found on AI Overview</h4>
    {% endif %}

    <!-- AI Sources Count -->
    <p><strong>Number of AI Sources in Organic Search (first 20):</strong> {{ data.ai_sources_in_organic_count }}</p>

    <div class="page-break"></div>

    <!-- Brand Mentions -->
    <h2>Brand Mentions</h2>
    <h3>YouTube</h3>
    {% if data.youtube_results %}
        <table>
        <thead>
            <tr>
            <th>Title</th><th>Displayed Link</th><th>Source</th><th>Snippet</th><th>Key Moments</th>
            </tr>
        </thead>
        <tbody>
            {% for yt in data.youtube_results %}
            <tr>
            <td><a href="{{ yt.link }}">{{ yt.title }}</a></td>
            <td>{{ yt.displayed_link }}</td>
            <td>{{ yt.source }}</td>
            <td>{{ yt.snippet }}</td>
            <td>{{ yt.key_moments }}</td>
            </tr>
            {% endfor %}
        </tbody>
        </table>
    {% else %}
        <p>No YouTube results found.</p>
    {% endif %}

    <h4>Suggestions:</h4>
    <ul>
        <li>Upload videos frequently.</li>
        <li>Write keyword-rich descriptions with timestamps and CTAs.</li>
    </ul>

    <!-- Social Channels -->
    <h3>Social Channels</h3>
    {% if data.social_channels %}
        <table>
        <thead><tr><th>Social Channel</th><th>Relevant Articles / Questions</th><th>Suggestions</th></tr></thead>
        <tbody>
            {% for ch in data.social_channels %}
            <tr>
            <td>{{ ch.channel }}</td>
            <td>
                <ul>
                {% for url, title in ch.relevant|urlize_links %}
                    <li><a href="{{ url }}">{{ title }}</a></li>
                {% else %}
                    <li>{{ ch.relevant }}</li>
                {% endfor %}
                </ul>
            </td>
            <td>{{ ch.suggestions }}</td>
            </tr>
            {% endfor %}
        </tbody>
        </table>
    {% else %}
        <p>No social channels data found.</p>
    {% endif %}

    <!-- Top SERP URLs -->
    <h2>Top SERP URLs</h2>
    {% if data.competitor_urls %}
        <ul>
        {% for url in data.competitor_urls %}
            <li><a href="{{ url }}">{{ url }}</a></li>
        {% endfor %}
        </ul>
    {% else %}
        <p>No competitor URLs found.</p>
    {% endif %}

    </body>
    </html>
    """
    # Create template from HTML_TEMPLATE using our custom environment
    template = env.from_string(HTML_TEMPLATE)
    
    # Create a new dict with combined data and domain
    template_data = {
        "data": data,
        "domain": data.get("domain", "Domain")
    }
    
    # Render template
    html_report = template.render(**template_data)

    # Auto-detect wkhtmltopdf
    wkhtmltopdf_path = shutil.which("wkhtmltopdf")
    
    # Try common wkhtmltopdf locations if not found in PATH
    if not wkhtmltopdf_path:
        common_paths = [
            r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe",
            r"C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe",
            "/usr/bin/wkhtmltopdf",
            "/usr/local/bin/wkhtmltopdf",
        ]
        for path in common_paths:
            if os.path.exists(path):
                wkhtmltopdf_path = path
                break

    # Verify the path
    if not wkhtmltopdf_path or not os.path.exists(wkhtmltopdf_path):
        st.warning("wkhtmltopdf not found. Attempting to generate PDF without it...")
        try:
            # Try with default configuration
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as pdf_file:
                pdfkit.from_string(html_report, pdf_file.name)
                return pdf_file.name
        except Exception as e:
            st.error(f"Error generating PDF: {str(e)}")
            st.info("Please install wkhtmltopdf from https://wkhtmltopdf.org/downloads.html")
            # Create an HTML file as fallback
            with tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode="w", encoding="utf-8") as html_file:
                html_file.write(html_report)
                st.success(f"Generated HTML report as fallback: {html_file.name}")
                return html_file.name
    
    # Configure pdfkit with the found wkhtmltopdf path
    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
    
    # Generate the PDF with proper options
    options = {
        'page-size': 'A4',
        'margin-top': '20mm',
        'margin-right': '20mm',
        'margin-bottom': '20mm',
        'margin-left': '20mm',
        'encoding': 'UTF-8',
        'no-outline': None,
        'enable-local-file-access': None,
    }
    
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as pdf_file:
            pdfkit.from_string(html_report, pdf_file.name, options=options, configuration=config)
            return pdf_file.name
    except Exception as e:
        st.error(f"Error generating PDF: {str(e)}")
        # Create an HTML file as fallback
        with tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode="w", encoding="utf-8") as html_file:
            html_file.write(html_report)
            st.success(f"Generated HTML report as fallback: {html_file.name}")
            return html_file.name
