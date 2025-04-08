import os
import streamlit as st
from docx.opc.constants import RELATIONSHIP_TYPE
from utils import get_serp_results, extract_domain, search_youtube_video, get_ai_overview_othersites, extract_competitor_urls, get_ai_overview_competitors, get_ai_overview_questions, check_domain_in_ai_overview, find_domain_position_in_organic, find_domain_position_in_ai, trim_url, get_ai_overview_content, analyze_target_content, get_social_results, rank_titles_by_semantic_similarity, get_youtube_results, get_ai_overview_competitors_content
from report_generator import generate_docx_report, generate_pdf_report
import requests

SOCIAL_SITES = ["youtube", "linkedin", "reddit", "quora"]
POPULAR_SITES = ["forbes", "pcmag", "techradar", "businessinsider", "techrepublic", "lifewire", "nytimes", "itpro", "macworld", "zdnet", "thectoclub", "techimply"]
REVIEW_SITES = ["gartner", "trustpilot", "crowdreviews", "capterra", "clutch", "softwarereviews", "softwaresuggest", "g2"]

# -------------------------------
# Streamlit UI Integration
# -------------------------------

st.set_page_config(page_title="SEO AI Analysis Report Generator", layout="wide")

st.title("SEO AI Analysis Report Generator")

social_ai_overviews = {}
popular_ai_overviews = {}
review_ai_overviews = {}
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
        ai_overview_competitors = get_ai_overview_competitors(serp_data, domain)
        domain_present = check_domain_in_ai_overview(serp_data, domain, target_url)
        domain_organic_position = find_domain_position_in_organic(serp_data, domain)
        domain_ai_position = find_domain_position_in_ai(serp_data, domain)
        competitor_urls_first20 = [trim_url(url) for url in competitor_urls[:20]]
        ai_sources_in_organic_count = sum(1 for source in ai_overview_competitors if source in competitor_urls_first20)
        ai_overview_content = get_ai_overview_content(serp_data)
        competitors = get_ai_overview_competitors_content(serp_data, domain)
        for site in SOCIAL_SITES:
            social_ai_overviews[site] = get_ai_overview_othersites(serp_data, site)
        for site in POPULAR_SITES:
            popular_ai_overviews[site] = get_ai_overview_othersites(serp_data, site)
        for site in REVIEW_SITES:
            review_ai_overviews[site] = get_ai_overview_othersites(serp_data, site)
        people_also_ask_ai_overview = get_ai_overview_questions(serp_data)

        st.info("Analyzing target URL content...")
        content_data = analyze_target_content(target_url, serp_data)
        
        st.info("Fetching social results from LinkedIn and Reddit...")

        try:
            r = requests.get("https://huggingface.co", timeout=5)
            st.success("Internet is working! ✅")
        except Exception as e:
            st.error(f"No internet: {e}")


        linkedin_results = get_social_results(keyword, "linkedin.com", limit_max=5, serp_api_key=SERPAPI_KEY)
        reddit_results = get_social_results(keyword, "reddit.com", limit_max=5, serp_api_key=SERPAPI_KEY)
        relevant_video = search_youtube_video(keyword, domain, serp_api_key = SERPAPI_KEY)
        linkedin_titles = [r["title"] for r in linkedin_results]
        reddit_titles = [r["title"] for r in reddit_results]
        # ranked_linkedin_titles = rank_titles_by_semantic_similarity(keyword, linkedin_titles, threshold=0.75)
        # ranked_reddit_titles = rank_titles_by_semantic_similarity(keyword, reddit_titles, threshold=0.75)
        social_channels = [
            {
                "channel": "LinkedIn",
                "relevant": "No relevant LinkedIn discussions found.",
                # "relevant": "<br><br>".join(
                #     [f"<a href='{linkedin_results[i]['link']}' target='_blank'>{title}</a><br><small>{linkedin_results[i]['link']}</small>"
                #      for i, (title, _) in enumerate(ranked_linkedin_titles)]
                # ) if ranked_linkedin_titles else "No relevant LinkedIn discussions found.",
                "suggestions": "Create an official LinkedIn presence and engage in relevant discussions."
            },
            {
                "channel": "Reddit",
                "relevant": "No relevant Reddit discussions found.",
                # "relevant": "<br><br>".join(
                #     [f"<a href='{reddit_results[i]['link']}' target='_blank'>{title}</a><br><small>{reddit_results[i]['link']}</small>"
                #      for i, (title, _) in enumerate(ranked_reddit_titles)]
                # ) if ranked_reddit_titles else "No relevant Reddit discussions found.",
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
            "ranked_reddit_titles": ranked_reddit_titles,
            "competitors": competitors,
            "social_ai_overview_sites": social_ai_overviews,
            "popular_ai_overview_sites": popular_ai_overviews,
            "review_ai_overview_sites": review_ai_overviews,
            "peopleAlsoAsk_ai_overview": people_also_ask_ai_overview,
            "relevant_video": relevant_video
        }
        
        st.success("Analysis complete!")

        # Display key results in Streamlit
        st.markdown("#### AI Overview Content")
        st.text(ai_overview_content)

        st.info("Generating DOCX Report")

        # Define output file path
        docx_output_file = "AIO_Report_"+keyword+".docx"
        pdf_output_file = "AIO_Report_"+keyword+".pdf"

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

        st.info("Generating PDF Report...")
        pdf_path = generate_pdf_report(report_data)
        with open(pdf_path, "rb") as file:
            st.download_button("Download PDF Report", data=file, file_name=pdf_output_file, mime="application/pdf")