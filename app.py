"""app.py"""
import os
import streamlit as st
from docx.opc.constants import RELATIONSHIP_TYPE
from utils import get_serp_results, perform_content_gap_analysis, get_50serp_results, extract_domain, search_youtube_video, analyze_secondary_content, get_ai_overview_othersites, get_competitors_content, extract_competitor_urls, get_ai_overview_competitors, get_ai_overview_questions, check_domain_in_ai_overview, find_domain_position_in_organic, find_domain_position_in_ai, trim_url, get_ai_overview_content, analyze_target_content, get_social_results, get_youtube_results, get_ai_overview_competitors_content, rank_titles_by_semantic_similarity
from report_generator import generate_docx_report

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
    s_Keyword = st.text_input("Secondary Keywords")
    st.info("Please enter maximum 5 Secondary Keywords comma seperated [, seperated]")
    submitted = st.form_submit_button("Run Analysis")
    
if submitted:
    # Initialize secondary_keywords right at the beginning
    secondary_keywords = [kw.strip() for kw in s_Keyword.split(',') if kw.strip()]
    
    if len(secondary_keywords) > 5:
        st.warning("You entered more than 5 secondary keywords. Only the first 5 will be used.")
        secondary_keywords = secondary_keywords[:5]
    
    # Access the API key from Streamlit's secrets
    try:
        SERPAPI_KEY = st.secrets["SERPAPI_KEY"]
        if not SERPAPI_KEY:
            st.error("SERP API Key is required!")
            st.stop()
    except KeyError:
        st.error("SERP API Key is not set in secrets. Please configure it.")
        st.stop()
        
    try:
        # Initialize dictionaries for secondary keyword data
        serp_data_secondary = {}
        domain_present_secondary = {}
        domain_organic_position_secondary = {}
        domain_ai_position_secondary = {}
        content_data_secondary = {}
        secondary_ai_overview_competitors = {} 
        secondary_ai_overview_content = {}
        secondary_serp_data_50 = {}
        
        with st.spinner("Fetching SERP data for keyword: " + keyword):
            serp_data = get_serp_results(keyword, SERPAPI_KEY)
            serp_data_50 = get_50serp_results(keyword, SERPAPI_KEY)
            domain = extract_domain(target_url).lower()
            competitor_urls = extract_competitor_urls(serp_data)
            ai_overview_competitors = get_ai_overview_competitors(serp_data, serp_data_50, domain)
            domain_present = check_domain_in_ai_overview(serp_data, domain, target_url)
            domain_organic_position = find_domain_position_in_organic(serp_data_50, domain)
            domain_ai_position = find_domain_position_in_ai(serp_data, domain)
            competitor_urls_first20 = [trim_url(url) for url in competitor_urls[:20]]
            ai_sources_in_organic_count = sum(1 for source in ai_overview_competitors if source.get("url", "") in competitor_urls_first20)
            ai_overview_content = get_ai_overview_content(serp_data)
            competitors = get_ai_overview_competitors_content(serp_data, domain)
            
            # Process site categories
            for site in SOCIAL_SITES:
                social_ai_overviews[site] = get_ai_overview_othersites(serp_data, site)
            for site in POPULAR_SITES:
                popular_ai_overviews[site] = get_ai_overview_othersites(serp_data, site)
            for site in REVIEW_SITES:
                review_ai_overviews[site] = get_ai_overview_othersites(serp_data, site)
                
            people_also_ask_ai_overview = get_ai_overview_questions(serp_data)

        with st.spinner("Analyzing target URL content..."):
            content_data = analyze_target_content(target_url, serp_data)

        # Initialize all_aio_content with primary keyword content
        all_aio_content = ai_overview_content + "\n\n"
        
        # Process secondary keywords
        for kw in secondary_keywords:
            with st.spinner(f"Fetching SERP data for secondary keyword: {kw}"):
                serp_data_secondary[kw] = get_serp_results(kw, SERPAPI_KEY)
                secondary_serp_data_50[kw] = get_50serp_results(kw, SERPAPI_KEY)
                domain_present_secondary[kw] = check_domain_in_ai_overview(serp_data_secondary[kw], domain, target_url)
                domain_organic_position_secondary[kw] = find_domain_position_in_organic(secondary_serp_data_50[kw], domain)
                domain_ai_position_secondary[kw] = find_domain_position_in_ai(serp_data_secondary[kw], domain)
                content_data_secondary[kw] = analyze_secondary_content(content_data["headers"], serp_data_secondary[kw])
                secondary_ai_overview_competitors[kw] = get_ai_overview_competitors(serp_data_secondary[kw], secondary_serp_data_50[kw], domain)
                
                # Get secondary keyword competitor URLs
                sec_competitor_urls = extract_competitor_urls(serp_data_secondary[kw])
                sec_competitor_urls_first20 = [trim_url(url) for url in sec_competitor_urls[:20]]
                ai_sources_in_organic_count += sum(1 for source in secondary_ai_overview_competitors[kw] if source.get("url", "") in sec_competitor_urls_first20)
                
                # Get and store secondary AI overview content
                secondary_ai_overview_content[kw] = get_ai_overview_content(serp_data_secondary[kw])
                all_aio_content += f"AIO content for {kw}: " + secondary_ai_overview_content.get(kw, "") + "\n\n"

                # Process site categories for secondary keywords
                for site in SOCIAL_SITES:
                    social_ai_overviews[site].extend(get_ai_overview_othersites(serp_data_secondary[kw], site)) 
                for site in POPULAR_SITES:
                    popular_ai_overviews[site].extend(get_ai_overview_othersites(serp_data_secondary[kw], site)) 
                for site in REVIEW_SITES:
                    review_ai_overviews[site].extend(get_ai_overview_othersites(serp_data_secondary[kw], site)) 

        try:
            # Combine all AI Overview competitors
            all_ai_competitors = ai_overview_competitors.copy()
            
            # Add sources from secondary keywords with keyword labeling
            for keyword_sec, competitors in secondary_ai_overview_competitors.items():
                for competitor in competitors:
                    # Add keyword info to the competitor
                    competitor_with_keyword = competitor.copy()
                    competitor_with_keyword["keyword"] = keyword_sec
                    all_ai_competitors.append(competitor_with_keyword)
            
            # Remove duplicates based on URL
            unique_urls = set()
            
            for competitor in all_ai_competitors:
                url = competitor.get("url")
                if url and url not in unique_urls:
                    unique_urls.add(url)
            
            # Access the API key from Streamlit's secrets for OpenAI
            OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY")
            if not OPENAI_API_KEY:
                st.warning("OpenAI API Key not found in secrets. Content Gap Analysis will be skipped.")
                content_gap_analysis = None
            else:
                # Perform content gap analysis for the main keyword using full page content
                with st.spinner("Performing Content Gap Analysis..."):
                    # Use the full page content instead of just headers
                    full_page_content = content_data.get("full_content", "")
                    
                    content_gap_analysis = perform_content_gap_analysis(
                        all_aio_content,
                        full_page_content,
                        unique_urls,
                        OPENAI_API_KEY
                    )
                    
        except Exception as e:
            st.error(f"Error performing content gap analysis: {str(e)}")
            content_gap_analysis = None

        with st.spinner("Fetching social results from LinkedIn and Reddit..."):
            # Get social results
            linkedin_results = get_social_results(keyword, "linkedin.com", limit_max=5, serp_api_key=SERPAPI_KEY)
            reddit_results = get_social_results(keyword, "reddit.com", limit_max=5, serp_api_key=SERPAPI_KEY)
            
            # Get YouTube info
            relevant_video = search_youtube_video(keyword, domain, serp_api_key=SERPAPI_KEY)
            
            # Initialize lists for titles - handle empty results
            linkedin_titles = [r["title"] for r in linkedin_results] if linkedin_results else []
            reddit_titles = [r["title"] for r in reddit_results] if reddit_results else []
            
            # Perform semantic ranking when we have results
            ranked_linkedin_titles = rank_titles_by_semantic_similarity(keyword, linkedin_titles, threshold=0.6) if linkedin_titles else []
            ranked_reddit_titles = rank_titles_by_semantic_similarity(keyword, reddit_titles, threshold=0.6) if reddit_titles else []
            
            # Format social channel data with links
            social_channels = [
                {
                    "channel": "LinkedIn",
                    "relevant": "<br><br>".join(
                        [f"<a href='{linkedin_results[linkedin_titles.index(title)]['link']}' target='_blank'>{title}</a>"
                         for title, _ in ranked_linkedin_titles]
                    ) if ranked_linkedin_titles else "No relevant LinkedIn discussions found.",
                    "suggestions": "Create an official LinkedIn presence and engage in relevant discussions."
                },
                {
                    "channel": "Reddit",
                    "relevant": "<br><br>".join(
                        [f"<a href='{reddit_results[reddit_titles.index(title)]['link']}' target='_blank'>{title}</a>"
                         for title, _ in ranked_reddit_titles]
                    ) if ranked_reddit_titles else "No relevant Reddit discussions found.",
                    "suggestions": "Participate in Reddit discussions to boost engagement."
                }
            ]
            
            youtube_results = get_youtube_results(keyword, limit_max=5, serp_api_key=SERPAPI_KEY)

        with st.spinner("Analyzing competitor content..."):
            ai_overview_competitor_content = get_competitors_content(ai_overview_competitors)
            
        # Format domain name for display
        if domain == "efax":
            domain_display = "eFax"
        else:
            domain_display = domain.title()
    
        # Compile data for reports
        report_data = {
            "keyword": keyword,
            "domain": domain_display,
            "target_url": target_url,
            "competitor_urls": competitor_urls,
            "ai_overview_competitors": ai_overview_competitors,
            "secondary_ai_overview_competitors": secondary_ai_overview_competitors,
            "domain_found": domain_present,
            "ai_sources_in_organic_count": ai_sources_in_organic_count,
            "ai_overview_content": ai_overview_content,
            "domain_organic_position": domain_organic_position,
            "domain_ai_position": domain_ai_position,
            "content_analysis": content_data,
            "social_channels": social_channels,
            "youtube_results": youtube_results,
            "competitors": competitors,
            "social_ai_overview_sites": social_ai_overviews,
            "popular_ai_overview_sites": popular_ai_overviews,
            "review_ai_overview_sites": review_ai_overviews,
            "peopleAlsoAsk_ai_overview": people_also_ask_ai_overview,
            "relevant_video": relevant_video,
            "secondary_keywords": secondary_keywords,
            "domain_organic_position_secondary": domain_organic_position_secondary,
            "domain_ai_position_secondary": domain_ai_position_secondary,
            "content_data_secondary": content_data_secondary,
            "secondary_ai_overview_content": secondary_ai_overview_content,
            "content_gap_analysis": content_gap_analysis,
            "aio_competitor_content": ai_overview_competitor_content
        }
        
        st.success("Analysis complete!")

        # Display key results in Streamlit
        st.markdown("#### AI Overview Content")
        st.text(ai_overview_content)

        with st.spinner("Generating DOCX Report"):
            # Define output file paths
            docx_output_file = f"AIO_Report_{keyword}.docx"
            pdf_output_file = f"AIO_Report_{keyword}.pdf"

            # Generate DOCX report
            generate_docx_report(report_data, domain_display, output_file=docx_output_file)

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

    except Exception as e:
        st.error(f"An error occurred during analysis: {str(e)}")
        st.info("Please check your inputs and try again.")