import os
import json
import re
import spacy
import sys
import unicodedata
import pandas as pd
import openai
from datetime import date, datetime
from imap_tools import MailBox, AND

# --- NLP INITIALIZATION ---
try:
    # Multilingual model for names/locations
    nlp = spacy.load("xx_ent_wiki_sm")
except:
    nlp = None

def anonymize_text(text):
    """Automatic anonymization for multiple languages (GDPR)"""
    if not nlp or not text:
        return text
    doc = nlp(text)
    entities = sorted(doc.ents, key=lambda x: len(x.text), reverse=True)
    for ent in entities:
        if ent.label_ in ["PER", "PERSON", "ORG", "GPE", "LOC"]:
            label = "PERSON" if ent.label_ in ["PER", "PERSON"] else "LOCATION"
            text = text.replace(ent.text, f"[{label}]")
    return text

def strip_accents(s):
    if not s: return ""
    return ''.join(c for c in unicodedata.normalize('NFD', s.lower())
                  if unicodedata.category(c) != 'Mn')

def clean_email_body(subject, text):
    """Cleaning and anonymization"""
    text = re.sub(r'https?://\S+|www\.\S+', '[LINK]', text)
    text = re.sub(r'\S+@\S+\.\S+', '[EMAIL]', text)
    subject = re.sub(r'\S+@\S+\.\S+', '[EMAIL]', subject)
    
    phone_pattern = r'\b(\+30|0030)?\s?[26789]\d{1,2}[\s\-]?\d{3,4}[\s\-]?\d{3,4}\b'
    text = re.sub(phone_pattern, '[PHONE]', text)
    
    signature_markers = ["--", "Kind regards", "Sincerely", "Regards", "Thank you", "Best regards"]
    lines = text.splitlines()
    clean_lines = []
    for line in lines:
        if any(marker.lower() in line.lower() for marker in signature_markers):
            break
        clean_lines.append(line)
    
    c_sub = anonymize_text(subject)
    c_txt = anonymize_text("\n".join(clean_lines))
    
    final_body = "\n".join([l.strip() for l in c_txt.splitlines() if l.strip()])
    return c_sub, final_body

def is_sender_blacklisted(email_sender, blacklist):
    if not email_sender: return False
    sender = email_sender.lower().strip()
    for blocked in blacklist:
        if blocked.lower().strip() in sender: return True
    return False

def analyze_with_ai(subject, body, company_info, api_key, model_name):
    client = openai.OpenAI(api_key=api_key)
    prompt_system = f"""
    You are a strict support assistant for the company {company_info['name']}.
    GOAL: Identify ONLY genuine complaints in the sector {company_info['industry']}.
    Focus on: {company_info['target_complaints']}.
    REPLY ONLY IN JSON: {{ "is_relevant": "YES" or "NO", "summary": "Summary" }}
    """
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "system", "content": prompt_system}, 
                      {"role": "user", "content": f"Subject: {subject}\nBody: {body}"}],
            response_format={"type": "json_object"},
            temperature=0
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        return {"is_relevant": "ERROR", "summary": str(e)}

def run_analysis_logic(config, status_container):
    settings = config.get('settings', {})
    api_key = config.get('api_keys', {}).get('openai', '')
    accounts = config.get('email_accounts', [])
    company = config.get('company_info', {})
    filters = config.get('filtering', {})
    fname = settings.get('excel_filename', 'Analysis.xlsx')
    
    date_str = config['settings'].get('last_run_date', date.today().strftime("%Y-%m-%d"))
    start_date = datetime.strptime(date_str, "%Y-%m-%d").date()
    existing_ids = set()
    old_df = pd.DataFrame()
    if os.path.exists(fname):
        try:
            old_df = pd.read_excel(fname)
            if 'ID' in old_df.columns:
                # Convert IDs to set for fast O(1) lookup
                existing_ids = set(old_df['ID'].astype(str).tolist())
        except Exception as e:
            status_container.warning(f"⚠️ Failed to read existing Excel file: {e}")
    
    if not api_key or not accounts:
        status_container.error("❌ Missing API keys or email accounts.")
        return [], {"relevant": 0, "irrelevant": 0, "errors": 0}

    if not nlp:
        status_container.warning("⚠️ NLP model not found. Name anonymization will be limited.")

    final_list = []
    stats = {"relevant": 0, "irrelevant": 0, "errors": 0}
    subj_blacklist = [strip_accents(w) for w in filters['subject_blacklist']]
    
    # UI setup for logs
    status_container.write(f"### 📅 Starting scan from: {start_date.strftime('%d/%m/%Y')}")
    log_window = status_container.empty() 
    log_buffer = []

    def update_logs(message, color="grey"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_buffer.append(f":{color}[[{timestamp}] {message}]")
        if len(log_buffer) > 5: log_buffer.pop(0)
        log_window.markdown("\n\n".join(log_buffer))

    # --- EMAIL CHECK ---
    for acc in accounts:
        update_logs(f"🔄 Connecting: {acc['user']}", "blue")
        try:
            with MailBox(acc['server']).login(acc['user'], acc['pass']) as mailbox:
                msg_uids = mailbox.uids(AND(date_gte=start_date))
                consecutive_errors = 0
                
                for idx, msg in enumerate(mailbox.fetch(AND(uid=msg_uids)), 1):
                    try:
                        if str(msg.uid) in existing_ids:
                            continue # Skip without logs to avoid flooding the screen
                        if is_sender_blacklisted(msg.from_, filters['sender_blacklist']) or \
                           any(word in strip_accents(msg.subject) for word in subj_blacklist):
                            stats["irrelevant"] += 1
                            continue

                        c_sub, c_txt = clean_email_body(msg.subject, msg.text or msg.html)
                        ans = analyze_with_ai(c_sub, c_txt, company, api_key, settings['ai_model'])
                        
                        if ans.get("is_relevant") == "ERROR":
                            update_logs(f"❗ AI error on email {idx}", "red")
                            stats["errors"] += 1
                            consecutive_errors += 1
                        else:
                            consecutive_errors = 0
                            if ans.get("is_relevant") == "YES":
                                update_logs(f"✅ COMPLAINT: {c_sub[:30]}...", "green")
                                stats["relevant"] += 1
                                final_list.append({
                                    "ID": msg.uid, 
                                    "Email Date": msg.date.strftime("%d-%m-%Y %H:%M"),
                                    "Account": acc['user'], 
                                    "Customer Email": msg.from_,
                                    "Subject": c_sub, 
                                    "Summary (AI)": ans.get("summary"), 
                                    "Status": "REVIEWED"
                                })
                            else:
                                update_logs(f"⚪ Irrelevant: {c_sub[:30]}...", "grey")
                                stats["irrelevant"] += 1

                        if consecutive_errors >= 3:
                            update_logs("🛑 Stop: Too many AI errors.", "red")
                            break

                    except Exception as e:
                        update_logs(f"⚠️ Error on email {idx}", "red")
                        stats["errors"] += 1
                        continue
        except Exception as e:
            status_container.error(f"❌ Connection error for {acc['user']}: {e}")

    # --- SAVE TO EXCEL ---
    if final_list:
        df_new = pd.DataFrame(final_list)
        fname = settings['excel_filename']
        
        if os.path.exists(fname):
            try:
                old_df = pd.read_excel(fname)
                df_final = pd.concat([old_df, df_new], ignore_index=True)
            except:
                df_final = df_new
        else:
            df_final = df_new

        try:
            # Use XlsxWriter for formatting
            writer = pd.ExcelWriter(fname, engine='xlsxwriter')
            df_final.to_excel(writer, index=False, sheet_name='Complaints')
            
            workbook  = writer.book
            worksheet = writer.sheets['Complaints']

            # Formats
            header_fmt = workbook.add_format({
                'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'
            })

            # Column setup (matching previous layout)
            worksheet.set_column('A:A', None, None, {'hidden': True}) # ID
            worksheet.set_column('B:B', 18) # Email Date
            worksheet.set_column('C:C', 25) # Account
            worksheet.set_column('D:D', 30) # Customer Email
            worksheet.set_column('E:E', 40) # Subject
            worksheet.set_column('F:F', 80) # Summary (AI)
            worksheet.set_column('G:G', 18) # Status

            # Εφαρμογή Header Format
            for col_num, value in enumerate(df_final.columns.values):
                worksheet.write(0, col_num, value, header_fmt)
            
            writer.close()
            update_logs(f"💾 File updated: {fname}", "blue")

        except PermissionError:
            # ERROR: file is open in Excel
            status_container.error(f"❌ ACCESS ERROR: The file '{fname}' is open in Excel. Close it and try again.")
            return [], stats 

    # Update date in config
    config['settings']['last_run_date'] = date.today().strftime("%Y-%m-%d")
    with open('config.json', 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4, ensure_ascii=False)
        
    return final_list, stats    