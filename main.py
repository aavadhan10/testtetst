import streamlit as st
import pandas as pd
import pdfplumber
import anthropic
import json
from docx import Document

st.set_page_config(page_title="Cap Table Audit", page_icon="📊", layout="wide")

class CapTableAuditor:
    def __init__(self):
        try:
            self.client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
        except:
            st.error("❌ Add ANTHROPIC_API_KEY to secrets.toml")
            self.client = None

    def extract_text(self, file):
        ext = file.name.split('.')[-1].lower()
        try:
            if ext == 'pdf':
                with pdfplumber.open(file) as pdf:
                    return "\n".join([p.extract_text() for p in pdf.pages])
            elif ext in ['doc', 'docx']:
                return "\n".join([p.text for p in Document(file).paragraphs])
            elif ext in ['xls', 'xlsx']:
                df = pd.read_excel(file, sheet_name=None)
                return "\n".join([f"Sheet {k}:\n{v.to_string()}" for k,v in df.items()])
            elif ext == 'csv':
                return pd.read_csv(file).to_string()
        except Exception as e:
            st.error(f"Error reading {file.name}: {e}")
            return ""

    def load_cap_table(self, files):
        dfs = []
        for file in files:
            ext = file.name.split('.')[-1].lower()
            try:
                if ext == 'csv':
                    df = pd.read_csv(file)
                elif ext in ['xls', 'xlsx']:
                    df = pd.read_excel(file)
                else:
                    continue
                df.columns = df.columns.str.strip().str.lower()
                dfs.append(df)
                st.success(f"✅ {file.name} ({len(df)} rows)")
            except Exception as e:
                st.error(f"❌ {file.name}: {e}")
        return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

    def analyze_with_claude(self, text, doc_name):
        if not self.client: return {"error": "No client"}
        
        prompt = f"""You are a lawyer conducting a capitalization table tie out of a company on behalf of an investor.

Analyze this legal document and extract ALL grant information:

Document: {doc_name}
Content: {text[:8000]}

Extract for each grant:
- Stockholder/Grantee name
- Number of shares/options granted
- Grant date (for board consents: the LAST date any director signed the consent, or the explicitly written effective date of the board approval. For board minutes: the date the meeting was held)
- Vesting start date
- Vesting schedule details
- Security type (options, shares, warrants, etc.)
- Exercise price if applicable

Return JSON format:
{{
  "grants": [
    {{
      "stockholder": "Full Name",
      "shares": "number",
      "grant_date": "YYYY-MM-DD",
      "vesting_start": "YYYY-MM-DD",
      "vesting_schedule": "description",
      "security_type": "options/shares/warrant",
      "exercise_price": "price if applicable",
      "document_reference": "specific section reference"
    }}
  ],
  "document_type": "board_consent/board_minutes/option_agreement/share_purchase_agreement/warrant/note/other"
}}"""

        try:
            response = self.client.messages.create(
                model="claude-3-sonnet-20240229",
                max_tokens=2000,
                messages=[{"role": "user", "content": prompt}]
            )
            text = response.content[0].text
            start, end = text.find('{'), text.rfind('}') + 1
            return json.loads(text[start:end]) if start != -1 else {"error": "No JSON"}
        except Exception as e:
            return {"error": str(e)}

    def compare_with_claude(self, cap_df, legal_docs):
        if not self.client: return {"error": "No client"}
        
        prompt = f"""You are a lawyer conducting a capitalization table tie out of a company on behalf of an investor.

1. Compare the company's capitalization table against the legal documents. The legal documents are the ultimate source of truth, and you are auditing the capitalization table to make sure it reflects the legal documents.

2. For each stockholder's grant in the capitalization table, confirm that the grant details, including the grant date, number of shares issued, vesting start date, and vesting schedule match what is approved in the corresponding board consent, board minutes, or other grant documents, including any share purchase agreement, option grant agreement, warrant, note, or other convertible securities.

3. The grant date in any board consent is the last date a director signed the consent, or the explicitly written effective date of the board approval. The grant date in any board minutes is the date the meeting was held.

4. Please list out all discrepancies, and provide a brief summary of each discrepancy, including which stockholder is impacted, what is incorrect on the capitalization table, and what it should be based on the legal documents, making reference to the specific legal document.

CAPITALIZATION TABLE:
{cap_df.head(20).to_dict('records')}

LEGAL DOCUMENTS ANALYSIS:
{legal_docs}

Return JSON format:
{{
  "discrepancies": [
    {{
      "stockholder": "Name",
      "discrepancy_type": "shares_mismatch/grant_date_mismatch/vesting_start_mismatch/missing_legal_doc/missing_cap_entry",
      "description": "Brief summary of discrepancy with specific details",
      "cap_table_value": "what the cap table shows",
      "legal_document_value": "what the legal document shows",
      "source_document": "specific legal document name",
      "severity": "high/medium/low"
    }}
  ],
  "summary": {{
    "total_discrepancies": 0,
    "high_severity_count": 0,
    "medium_severity_count": 0,
    "low_severity_count": 0,
    "overall_assessment": "Summary of cap table accuracy"
  }}
}}"""

        try:
            response = self.client.messages.create(
                model="claude-3-sonnet-20240229",
                max_tokens=3000,
                messages=[{"role": "user", "content": prompt}]
            )
            text = response.content[0].text
            start, end = text.find('{'), text.rfind('}') + 1
            return json.loads(text[start:end]) if start != -1 else {"error": "No JSON"}
        except Exception as e:
            return {"error": str(e)}

def main():
    st.title("📊 Cap Table Audit Tool")
    
    auditor = CapTableAuditor()
    if not auditor.client: st.stop()
    
    # File uploads
    col1, col2 = st.columns(2)
    with col1:
        cap_files = st.file_uploader("Cap Table Files", type=['csv','xlsx','xls'], accept_multiple_files=True)
    with col2:
        legal_files = st.file_uploader("Legal Documents", type=['pdf','doc','docx','xlsx','xls','csv'], accept_multiple_files=True)
    
    if cap_files and legal_files:
        # Load cap table
        cap_df = auditor.load_cap_table(cap_files)
        if cap_df.empty: return
        
        with st.expander("Cap Table Preview"):
            st.dataframe(cap_df.head())
        
        # Analyze legal docs
        legal_analysis = []
        for i, file in enumerate(legal_files):
            st.write(f"📄 {file.name}")
            text = auditor.extract_text(file)
            if text:
                analysis = auditor.analyze_with_claude(text, file.name)
                if "error" not in analysis:
                    legal_analysis.append(analysis)
                    st.success("✅")
                else:
                    st.error(f"❌ {analysis['error']}")
        
        # Compare
        if legal_analysis:
            result = auditor.compare_with_claude(cap_df, legal_analysis)
            
            if "error" not in result:
                discrepancies = result.get("discrepancies", [])
                st.metric("Discrepancies", len(discrepancies))
                
                if discrepancies:
                    for d in discrepancies:
                        severity = {"high": "🔴", "medium": "🟡", "low": "🟢"}.get(d.get("severity"), "⚪")
                        with st.expander(f"{severity} {d.get('stockholder', 'Unknown')} - {d.get('discrepancy_type', 'Issue')}"):
                            st.write(f"**Description:** {d.get('description', 'No details')}")
                            st.write(f"**Cap Table Value:** {d.get('cap_table_value', 'N/A')}")
                            st.write(f"**Legal Document Value:** {d.get('legal_document_value', 'N/A')}")
                            st.write(f"**Source Document:** {d.get('source_document', 'N/A')}")
                    
                    # Show summary metrics
                    summary = result.get("summary", {})
                    if summary:
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("High Severity", summary.get("high_severity_count", 0))
                        with col2:
                            st.metric("Medium Severity", summary.get("medium_severity_count", 0))
                        with col3:
                            st.metric("Low Severity", summary.get("low_severity_count", 0))
                        
                        if summary.get("overall_assessment"):
                            st.info(f"**Assessment:** {summary['overall_assessment']}")
                    
                    # Download
                    csv = pd.DataFrame(discrepancies).to_csv(index=False)
                    st.download_button("📥 Download Report", csv, "audit.csv")
                else:
                    st.success("✅ No issues found!")
            else:
                st.error(f"Error: {result['error']}")

if __name__ == "__main__":
    main()
