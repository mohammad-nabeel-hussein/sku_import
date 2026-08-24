import io
import time
from datetime import datetime
import pandas as pd
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# --- ScraperAPI Key Integration ---
SCRAPER_API_KEY = "955f333964f3d5f59e0ed5f4037d6ea1"

# --- Page Config ---
st.set_page_config(
    page_title="Noon SKU Extractor Dashboard", 
    page_icon="🛍️", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS Styling ---
st.markdown("""
<style>
    [data-testid="collapsedControl"] { display: none !important; }
    section[data-testid="stSidebar"] { width: 320px !important; min-width: 320px !important; }
    header[data-testid="stHeader"] { display: none !important; }
    footer { visibility: hidden !important; height: 0px !important; }
    #MainMenu { visibility: hidden !important; }
</style>
""", unsafe_allow_html=True)

CATEGORY_MAP = {
    "Toys & Games": "toys-and-games",
    "Health & Nutrition": "health",
    "Home & Kitchen": "home-and-kitchen",
    "Other / Custom Slug": "custom"
}

if "scrape_history" not in st.session_state:
    st.session_state.scrape_history = []

# --- STATIC SIDEBAR ---
st.sidebar.header("⚙️ Search Configuration")

country = st.sidebar.selectbox("Store Country", options=["Saudi Arabia", "UAE"], index=0)
selected_category_label = st.sidebar.selectbox("Category", options=list(CATEGORY_MAP.keys()), index=0)

if CATEGORY_MAP[selected_category_label] == "custom":
    category_slug = st.sidebar.text_input("Custom Category Slug", value="electronics")
else:
    category_slug = CATEGORY_MAP[selected_category_label]

search_term = st.sidebar.text_input("Search Query", value="saudi national day balloon")

p_col1, p_col2 = st.sidebar.columns(2)
with p_col1:
    start_page = st.sidebar.number_input("Start Page", min_value=1, value=1, step=1)
with p_col2:
    end_page = st.sidebar.number_input("End Page", min_value=1, value=2, step=1)

country_code = "saudi-en" if country == "Saudi Arabia" else "uae-en"

st.sidebar.write("")
start_btn = st.sidebar.button("🚀 Start Extraction", type="primary", use_container_width=True)

st.sidebar.divider()
st.sidebar.subheader("📜 Run History")
if st.session_state.scrape_history:
    history_df = pd.DataFrame(st.session_state.scrape_history)
    st.sidebar.dataframe(history_df, use_container_width=True, hide_index=True)
else:
    st.sidebar.info("No runs recorded yet.")

# --- MAIN DASHBOARD AREA ---
st.title("🛍️ Noon SKU Extractor Dashboard")
st.caption("Configure parameters in the sidebar and run extraction.")

progress_placeholder = st.empty()
status_placeholder = st.empty()

metrics_cols = st.columns(3)
m_page = metrics_cols[0].empty()
m_skus = metrics_cols[1].empty()
m_eta = metrics_cols[2].empty()

m_page.metric("Page Progress", f"- / -")
m_skus.metric("SKUs Collected", "0")
m_eta.metric("Estimated ETA", "00:00")

download_area = st.container()

# --- Proxy Scraper Engine (Fixes HTTP 500 Error) ---
def run_proxy_scraper(country_path, cat_slug, query, start_p, end_p):
    sku_to_page = {}
    total_pages = (end_p - start_p) + 1
    start_time = time.time()
    pages_done = 0

    session = requests.Session()

    for current_page in range(start_p, end_p + 1):
        status_placeholder.info(f"Fetching Page {current_page} of {end_p} via Proxy...")

        # Formulate clean search URL
        formatted_query = query.strip().replace(" ", "+")
        if cat_slug and cat_slug != "custom":
            target_noon_url = f"https://www.noon.com/{country_path}/{cat_slug}/?page={current_page}&q={formatted_query}"
        else:
            target_noon_url = f"https://www.noon.com/{country_path}/search/?page={current_page}&q={formatted_query}"
        
        # ScraperAPI call without JS render parameter to avoid 500 errors
        proxy_url = f"http://api.scraperapi.com?api_key={SCRAPER_API_KEY}&url={target_noon_url}&country_code=sa"

        try:
            res = session.get(proxy_url, timeout=45)

            if res.status_code == 200:
                html_text = res.text
                
                # Extract product SKUs matching Noon's catalog pattern (e.g. N12345678A or Z12345678A)
                found_skus = set(re.findall(r"/([A-Z0-9]{10,})/p/", html_text))
                
                new_items_found = 0
                for sku in found_skus:
                    if sku not in sku_to_page:
                        sku_to_page[sku] = current_page
                        new_items_found += 1

                if new_items_found == 0 and len(sku_to_page) > 0:
                    status_placeholder.warning(f"No new SKUs found on page {current_page}. Stopping.")
                    break
            else:
                status_placeholder.error(f"Page {current_page} returned HTTP {res.status_code}.")
                break

        except Exception as e:
            status_placeholder.error(f"Error loading page {current_page}: {e}")
            break

        pages_done += 1
        pct = pages_done / total_pages
        progress_placeholder.progress(pct)

        elapsed = time.time() - start_time
        avg_time = elapsed / pages_done
        eta_seconds = int((total_pages - pages_done) * avg_time)
        eta_str = time.strftime("%M:%S", time.gmtime(eta_seconds)) if eta_seconds > 0 else "00:00"

        m_page.metric("Page Progress", f"{current_page} / {end_p}")
        m_skus.metric("SKUs Collected", len(sku_to_page))
        m_eta.metric("Estimated ETA", eta_str)

        time.sleep(1)

    return sku_to_page
# --- Excel Generator ---
def generate_excel_export(sku_to_page_map, country_path):
    wb = openpyxl.Workbook()
    ws_detailed = wb.active
    ws_detailed.title = "SKU Catalog"
    ws_summary = wb.create_sheet(title="Summary")

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    border_thin = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
                         top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    headers = ["Item #", "Product SKU", "Source Page", "Direct Link"]
    ws_detailed.append(headers)
    ws_detailed.row_dimensions[1].height = 24

    for c in range(1, 5):
        cell = ws_detailed.cell(row=1, column=c)
        cell.fill, cell.font, cell.alignment, cell.border = header_fill, header_font, align_center, border_thin

    sorted_skus = sorted(sku_to_page_map.items(), key=lambda x: (x[1], x[0]))

    for idx, (sku, p_num) in enumerate(sorted_skus, start=1):
        row_idx = idx + 1
        product_url = f"https://www.noon.com/{country_path}/{sku}/p/"
        ws_detailed.append([idx, sku, f"Page {p_num}", product_url])
        ws_detailed.row_dimensions[row_idx].height = 20

        c1, c2, c3, c4 = [ws_detailed.cell(row=row_idx, column=i) for i in range(1, 5)]
        c1.alignment = c2.alignment = c3.alignment = align_center
        c4.alignment = align_left
        c1.border = c2.border = c3.border = c4.border = border_thin

        if row_idx % 2 == 0:
            zebra = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
            c1.fill = c2.fill = c3.fill = c4.fill = zebra

    for col in ws_detailed.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_detailed.column_dimensions[col_letter].width = max(max_len + 4, 15)

    ws_summary.append(["Page Number", "Products Count"])
    ws_summary.row_dimensions[1].height = 24
    for c in (1, 2):
        cell = ws_summary.cell(row=1, column=c)
        cell.fill, cell.font, cell.alignment, cell.border = header_fill, header_font, align_center, border_thin

    page_counts = {}
    for _, p_num in sorted_skus:
        page_counts[p_num] = page_counts.get(p_num, 0) + 1

    s_row = 2
    for p_num in sorted(page_counts.keys()):
        ws_summary.append([f"Page {p_num}", page_counts[p_num]])
        ws_summary.row_dimensions[s_row].height = 20
        for c in (1, 2):
            cell = ws_summary.cell(row=s_row, column=c)
            cell.border, cell.alignment = border_thin, align_center
            if s_row % 2 == 0:
                cell.fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
        s_row += 1

    ws_summary.append(["Total Unique SKUs", len(sku_to_page_map)])
    ws_summary.row_dimensions[s_row].height = 22
    for c in (1, 2):
        cell = ws_summary.cell(row=s_row, column=c)
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.border, cell.alignment = border_thin, align_center

    for col in ws_summary.columns:
        col_letter = get_column_letter(col[0].column)
        ws_summary.column_dimensions[col_letter].width = 20

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

# --- Execution Trigger ---
if start_btn:
    if start_page > end_page:
        st.error("Start Page cannot be greater than End Page.")
    elif not search_term.strip():
        st.warning("Please enter a search term.")
    else:
        st_start_time = time.time()

        extracted_skus = run_proxy_scraper(
            country_code, category_slug, search_term, start_page, end_page
        )

        elapsed_sec = round(time.time() - st_start_time, 2)

        if not extracted_skus:
            st.error("No SKUs found. Check your search query or page range.")
        else:
            st.balloons()
            status_placeholder.success(f"Successfully extracted {len(extracted_skus)} products in {elapsed_sec} seconds!")

            st.session_state.scrape_history.append({
                "Time": datetime.now().strftime("%H:%M:%S"),
                "Store": country,
                "Query": search_term,
                "Pages": f"{start_page}-{end_page}",
                "SKUs": len(extracted_skus)
            })

            sorted_items = sorted(extracted_skus.items(), key=lambda x: (x[1], x[0]))
            csv_rows = []
            for idx, (sku, p_num) in enumerate(sorted_items, start=1):
                csv_rows.append({
                    "Item #": idx,
                    "Product SKU": sku,
                    "Source Page": f"Page {p_num}",
                    "Direct Link": f"https://www.noon.com/{country_code}/{sku}/p/"
                })
            df_csv = pd.DataFrame(csv_rows)

            excel_bytes = generate_excel_export(extracted_skus, country_code)
            fn_base = f"noon_{country_code}_{category_slug}_p{start_page}_to_p{end_page}_{search_term.replace(' ', '_')}"

            with download_area:
                st.divider()
                st.subheader("📬 Export Downloads")
                col_exp1, col_exp2 = st.columns(2)

                with col_exp1:
                    st.download_button(
                        label="📊 Download Excel (.xlsx)",
                        data=excel_bytes,
                        file_name=f"{fn_base}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )

                csv_data = df_csv.to_csv(index=False).encode('utf-8')
                with col_exp2:
                    st.download_button(
                        label="📄 Download CSV (.csv)",
                        data=csv_data,
                        file_name=f"{fn_base}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
