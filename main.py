import streamlit as st
import pandas as pd
import pdfplumber
import anthropic
import json
from typing import Dict, List

st.set_page_config(page_title="Cap Table Audit Tool", page_icon="📊", layout="wide")

class CapTableAuditor:
    def __init__(self, api_key=None):
        self.client = anthropic.Anthropic(api_key=api_key) if api_key else None
    
    def extract_pdf_text(self, pdf_file) -> str:
        try:
            with pdfplumber.open(pdf_file) as pdf:
                return "\n".join([page.extract_text() for page in pdf.pages])
        except Exception as e:
            st.error(f"Error extracting PDF: {e}")
            return ""
    
    def analyze_with_claude(self, text: str, doc_name: str) -> Dict:
        if not self.client:
            return {"error": "No API key"}
        
        prompt = f"""You are a lawyer conducting a capitalization table tie out. Analyze this legal document and extract ALL grant information:

Document: {doc_name}
Content: {text}

Extract for EACH grant:
1. Stockholder/Grantee name
2. Number of shares/options
3. Grant date (for board consents: LAST signature date or explicit effective date; for board minutes: meeting date)
4. Vesting start date
5. Vesting schedule
6. Security type
7. Exercise price

Return JSON format:
{{
  "grants": [
    {{
      "stockholder": "Name",
      "shares": "number",
      "grant_date": "YYYY-MM-DD", 
      "vesting_start": "YYYY-MM-DD",
      "security_type": "options/shares/warrant",
      "exercise_price": "price",
      "notes": "details"
    }}
  ]
}}"""

        try:
            response = self.client.messages.create(
                model="claude-3-sonnet-20240229",
                max_tokens=4000,
                temperature=0,
                messages=[{"role": "user", "content": prompt}]
            )
            
            text = response.content[0].text
            start = text.find('{')
            end = text.rfind('}') + 1
            
            if start != -1 and end != -1:
                return json.loads(text[start:end])
            return {"error": "No JSON found"}
        except Exception as e:
            return {"error": str(e)}
    
    def compare_with_claude(self, cap_table_df: pd.DataFrame, legal_analysis: List[Dict]) -> Dict:
        if not self.client:
            return {"error": "No API key"}
        
        prompt = f"""Compare cap table against legal documents. Legal docs are source of truth.

CAP TABLE:
{cap_table_df.to_dict('records')}

LEGAL DOCUMENTS:
{legal_analysis}

Find discrepancies in: grant dates, share counts, vesting dates, missing entries.

Return JSON:
{{
  "discrepancies": [
    {{
      "stockholder": "Name",
      "discrepancy_type": "shares_mismatch/date_mismatch/missing_doc",
      "description": "Detailed issue",
      "severity": "high/medium/low",
      "cap_table_value": "what cap table shows",
      "legal_doc_value": "what legal doc shows",
      "source_document": "document name"
    }}
  ],
  "summary": {{
    "total_discrepancies": 0,
    "assessment": "overall findings"
  }}
}}"""

        try:
            response = self.client.messages.create(
                model="claude-3-sonnet-20240229",
                max_tokens=4000,
                temperature=0,
                messages=[{"role": "user", "content": prompt}]
            )
            
            text = response.content[0].text
            start = text.find('{')
            end = text.rfind('}') + 1
            
            if start != -1 and end != -1:
                return json.loads(text[start:end])
            return {"error": "No JSON found"}
        except Exception as e:
            return {"error": str(e)}

def main():
    st.title("📊 AI Cap Table Audit Tool")
    
    # API Key
    api_key = st.sidebar.text_input("Claude API Key:", type="password")
    if not api_key:
        st.sidebar.warning("Enter Claude API key for AI analysis")
    
    auditor = CapTableAuditor(api_key)
    
    # File uploads
    col1, col2 = st.columns(2)
    with col1:
        cap_file = st.file_uploader("Cap Table (CSV)", type=['csv'])
    with col2:
        legal_files = st.file_uploader("Legal Documents (PDFs)", type=['pdf'], accept_multiple_files=True)
    
    if cap_file and legal_files and api_key:
        # Process cap table
        cap_df = pd.read_csv(cap_file)
        st.success(f"✅ Loaded {len(cap_df)} cap table entries")
        
        with st.expander("Cap Table Preview"):
            st.dataframe(cap_df.head())
        
        # Analyze legal documents
        legal_analysis = []
        progress = st.progress(0)
        
        for i, file in enumerate(legal_files):
            st.write(f"📄 Analyzing: {file.name}")
            text = auditor.extract_pdf_text(file)
            analysis = auditor.analyze_with_claude(text, file.name)
            
            if "error" not in analysis:
                legal_analysis.append(analysis)
                st.success(f"✅ {file.name}")
            else:
                st.error(f"❌ {file.name}: {analysis['error']}")
            
            progress.progress((i + 1) / len(legal_files))
        
        # Compare with Claude
        if legal_analysis:
            comparison = auditor.compare_with_claude(cap_df, legal_analysis)
            
            if "error" not in comparison:
                discrepancies = comparison.get("discrepancies", [])
                summary = comparison.get("summary", {})
                
                # Results
                st.header("🔍 Audit Results")
                st.metric("Total Discrepancies", summary.get("total_discrepancies", 0))
                
                if summary.get("assessment"):
                    st.info(summary["assessment"])
                
                if discrepancies:
                    st.error(f"❌ Found {len(discrepancies)} issues")
                    
                    for i, disc in enumerate(discrepancies, 1):
                        severity_icon = {"high": "🔴", "medium": "🟡", "low": "🟢"}.get(disc.get("severity"), "⚪")
                        
                        with st.expander(f"{severity_icon} {disc['stockholder']} - {disc.get('discrepancy_type', 'Issue')}", expanded=disc.get('severity') == 'high'):
                            st.write(f"**Issue:** {disc['description']}")
                            st.write(f"**Cap Table:** {disc.get('cap_table_value', 'N/A')}")
                            st.write(f"**Legal Doc:** {disc.get('legal_doc_value', 'N/A')}")
                            st.write(f"**Source:** {disc.get('source_document', 'N/A')}")
                    
                    # Download report
                    report_df = pd.DataFrame(discrepancies)
                    csv = report_df.to_csv(index=False)
                    st.download_button("📥 Download Report", csv, "audit_report.csv", "text/csv")
                else:
                    st.success("✅ No discrepancies found!")
            else:
                st.error(f"Comparison error: {comparison['error']}")
    
    elif cap_file and legal_files and not api_key:
        st.warning("⚠️ Enter Claude API key for full analysis")

if __name__ == "__main__":
    main()
