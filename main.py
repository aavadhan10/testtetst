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
        
        prompt = f"""Extract grant info from this legal document:

{text[:8000]}

Return JSON:
{{
  "grants": [
    {{"stockholder": "Name", "shares": "123", "grant_date": "2024-01-01", "vesting_start": "2024-01-01"}}
  ]
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
        
        prompt = f"""Compare cap table vs legal docs. Find discrepancies:

CAP TABLE: {cap_df.head(20).to_dict('records')}
LEGAL DOCS: {legal_docs}

Return JSON:
{{
  "discrepancies": [
    {{"stockholder": "Name", "issue": "description", "severity": "high/medium/low"}}
  ],
  "summary": {{"total": 0}}
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
        cap_files = st.file_uploader("Cap Table", type=['csv','xlsx','xls'], accept_multiple_files=True)
    with col2:
        legal_files = st.file_uploader("Legal Docs", type=['pdf','docx','doc','xlsx','csv'], accept_multiple_files=True)
    
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
                        with st.expander(f"{severity} {d.get('stockholder', 'Unknown')}"):
                            st.write(d.get('issue', 'No details'))
                    
                    # Download
                    csv = pd.DataFrame(discrepancies).to_csv(index=False)
                    st.download_button("📥 Download Report", csv, "audit.csv")
                else:
                    st.success("✅ No issues found!")
            else:
                st.error(f"Error: {result['error']}")

if __name__ == "__main__":
    main()
