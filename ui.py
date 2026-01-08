import streamlit as st
import pandas as pd
import json
import os
from datetime import date, datetime
import plotly.express as px
from main_script import run_analysis_logic

# --- PAGE SETTINGS ---
st.set_page_config(page_title="AI Email Analyzer", layout="wide", initial_sidebar_state="expanded")

# --- LOGIC FUNCTIONS ---
def load_config():
    if not os.path.exists('config.json'):
        return {
            "api_keys": {"openai": ""}, 
            "email_accounts": [], 
            "company_info": {"name": "", "industry": "", "target_complaints": ""}, 
            "filtering": {"subject_blacklist": [], "sender_blacklist": []}, 
            "settings": {
                "excel_filename": "Analysis.xlsx", 
                "ai_model": "gpt-4o-mini",
                "last_run_date": date.today().strftime("%Y-%m-%d")
            }
        }
    # Open config file
    with open('config.json', 'r', encoding='utf-8') as f:
        conf = json.load(f)
        # Safety fallback if last_run_date is missing
        if 'last_run_date' not in conf['settings']:
            conf['settings']['last_run_date'] = date.today().strftime("%Y-%m-%d")
        return conf
def save_config(config):
    with open('config.json', 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

# Load config into session state
if 'config' not in st.session_state:
    st.session_state.config = load_config()

# --- SIDEBAR (Navigation) ---
with st.sidebar:
    st.divider()
    if st.button("🛑 Quit Application", use_container_width=True):
        st.warning("The application is closing... You can close this tab.")
        # Use os._exit to fully terminate the EXE process
        os._exit(0)
    st.title("📧 Email AI Menu")
    menu = st.radio("Select mode:", ["🏠 Dashboard", "⚙️ Settings"])
    st.info("Tip: Save settings before the first run.")

# --- MAIN CONTENT ---
if menu == "🏠 Dashboard":
    st.title("🚀 Run Analysis")

# Πλαίσιο Ελέγχου Ημερομηνίας και Εκτέλεσης
    with st.container(border=True):
        col_run, col_date = st.columns([1, 1])
        
        with col_run:
            if st.button("▶️ START SCAN NOW", type="primary", use_container_width=True):
                if not st.session_state.config['api_keys']['openai']:
                    st.error("❌ Please enter the OpenAI Key in Settings.")
                else:
                    with st.status("🔍 Searching and analyzing...", expanded=True) as status:
                        # Call analysis logic (general_dokimi.py)
                        results, final_stats = run_analysis_logic(st.session_state.config, status)
                        status.update(label="✅ Analysis completed!", state="complete", expanded=False)

                    # Save results for display in metrics/charts
                    st.session_state.last_results = results
                    st.session_state.last_stats = final_stats
                    # Reload config
                    st.session_state.config = load_config()

    with col_date:
        # Date selector on the right column
        date_str = st.session_state.config['settings'].get('last_run_date', date.today().strftime("%Y-%m-%d"))
        last_date_val = datetime.strptime(date_str, "%Y-%m-%d").date()

        selected_date = st.date_input("📅 Check emails from:", last_date_val)

        st.session_state.config['settings']['last_run_date'] = selected_date.strftime("%Y-%m-%d")
        st.caption(f"ℹ️ Analysis will start from: **{selected_date.strftime('%d/%m/%Y')}**")
    # Εμφάνιση Αποτελεσμάτων
    if 'last_stats' in st.session_state:
        st.divider()
        st.subheader("📊 Last Scan Results")

        m1, m2, m3 = st.columns(3)
        m1.metric("New Complaints", st.session_state.last_stats["relevant"])
        m2.metric("Irrelevant/Spam", st.session_state.last_stats["irrelevant"])
        m3.metric("Errors", st.session_state.last_stats["errors"])

        col_chart, col_btns = st.columns([1, 1])
        
        with col_chart:
            if st.session_state.last_stats["relevant"] + st.session_state.last_stats["irrelevant"] > 0:
                df_pie = pd.DataFrame({
                    "Category": ["Complaints", "Irrelevant"],
                    "Count": [st.session_state.last_stats["relevant"], st.session_state.last_stats["irrelevant"]]
                })
                fig = px.pie(df_pie, values='Count', names='Category', hole=0.4,
                             color_discrete_sequence=['#2ecc71', '#bdc3c7'])
                st.plotly_chart(fig, use_container_width=True)

        with col_btns:
            st.write("### 📂 File Management")
            fname = st.session_state.config['settings']['excel_filename']

            if st.session_state.last_results:
                st.success(f"Found {len(st.session_state.last_results)} new complaints!")

                # Button to open Excel (local)
                if st.button("📂 OPEN EXCEL FILE", use_container_width=True):
                    if os.path.exists(fname):
                        os.startfile(fname)
                    else:
                        st.error("The file has not been created yet.")

                # Download button
                if os.path.exists(fname):
                    with open(fname, "rb") as f:
                        st.download_button("📥 DOWNLOAD FILE", data=f, file_name=fname, use_container_width=True)
            else:
                st.info("No new complaints were found during the last scan.")
elif menu == "⚙️ Settings":
    st.title("⚙️ Settings Management")
    
    # Χρήση Tabs για οργάνωση των ρυθμίσεων
    # Use tabs to organize settings
    tab1, tab2, tab3 = st.tabs(["🔑 Access & Emails", "🏢 Company & AI", "🛡️ Filters & GDPR"])

    with tab1:
        st.subheader("OpenAI Authentication")
        st.session_state.config['api_keys']['openai'] = st.text_input("OpenAI Key", st.session_state.config['api_keys']['openai'], type="password")

        st.divider()
        st.subheader("Email Accounts")
        for i, acc in enumerate(st.session_state.config['email_accounts']):
            with st.expander(f"📧 {acc['user'] if acc['user'] else 'New Account'}"):
                acc['user'] = st.text_input(f"Email", acc['user'], key=f"u{i}")
                acc['pass'] = st.text_input(f"Password", acc['pass'], type="password", key=f"p{i}")
                acc['server'] = st.text_input(f"IMAP Server", acc['server'], key=f"s{i}")
                if st.button(f"🗑️ Delete {i}", key=f"del{i}"):
                    st.session_state.config['email_accounts'].pop(i)
                    st.rerun()

        if st.button("➕ Add New Account"):
            st.session_state.config['email_accounts'].append({"user": "", "pass": "", "server": "imap.gmail.com"})
            st.rerun()

    with tab2:
        col_n, col_i = st.columns(2)
        st.session_state.config['company_info']['name'] = col_n.text_input("Company Name", st.session_state.config['company_info']['name'])
        st.session_state.config['company_info']['industry'] = col_i.text_input("Industry", st.session_state.config['company_info']['industry'])
        st.session_state.config['company_info']['target_complaints'] = st.text_area("Complaint Description (for AI)", st.session_state.config['company_info']['target_complaints'], help="Explain to the AI what you consider a complaint.")

    with tab3:
        f1, f2 = st.columns(2)
        with f1:
            sub = st.text_area("Subject Blacklist (comma-separated)", ", ".join(st.session_state.config['filtering']['subject_blacklist']))
            st.session_state.config['filtering']['subject_blacklist'] = [x.strip() for x in sub.split(",") if x.strip()]
        with f2:
            snd = st.text_area("Sender Blacklist (comma-separated)", ", ".join(st.session_state.config['filtering']['sender_blacklist']))
            st.session_state.config['filtering']['sender_blacklist'] = [x.strip() for x in snd.split(",") if x.strip()]

        names = st.text_area("Names to Anonymize (comma-separated)", ", ".join(st.session_state.config['company_info'].get('anonymize_names', [])))
        st.session_state.config['company_info']['anonymize_names'] = [x.strip() for x in names.split(",") if x.strip()]

    # Persistent save button at the bottom of settings
    st.divider()
    if st.button("💾 SAVE ALL SETTINGS", type="primary", use_container_width=True):
        save_config(st.session_state.config)
        st.success("✅ Settings saved successfully!")