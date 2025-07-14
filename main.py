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
                    # Try different CSV parsing options
                    try:
                        df = pd.read_csv(file)
                    except:
                        # Try with different separator
                        file.seek(0)
                        try:
                            df = pd.read_csv(file, sep=';')
                        except:
                            # Try with different quote handling
                            file.seek(0)
                            try:
                                df = pd.read_csv(file, quotechar='"', quoting=1)
                            except:
                                # Last resort - skip bad lines
                                file.seek(0)
                                df = pd.read_csv(file, error_bad_lines=False, warn_bad_lines=True)
                elif ext in ['xls', 'xlsx']:
                    df = pd.read_excel(file)
                else:
                    continue
                
                df.columns = df.columns.str.strip().str.lower()
                dfs.append(df)
                st.success(f"✅ {file.name} ({len(df)} rows)")
            except Exception as e:
                st.error(f"❌ {file.name}: {e}")
                st.info("💡 Try saving as Excel (.xlsx) format if CSV continues to fail")
        return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

    def analyze_with_claude(self, text, doc_name):
        if not self.client: return {"error": "No client"}
        
        # More aggressive text cleaning
        import re
        clean_text = re.sub(r'[^\w\s\-\.\,\:\;\(\)\[\]\/]', ' ', text)  # Remove special chars
        clean_text = ' '.join(clean_text.split())  # Normalize whitespace
        
        prompt = f"""You are a lawyer conducting a capitalization table tie out. Analyze this document and extract EXACT grant details.

Document: {doc_name}
Content: {clean_text[:8000]}

Extract PRECISE details for each grant:
- Stockholder name
- EXACT share count (e.g., "10000", "5000")
- Grant date (board consent: LAST signature date; board minutes: meeting date)
- Vesting start date
- DETAILED vesting schedule (e.g., "1/48th monthly over 4 years", "25% after 1 year then 1/36th monthly", "100% immediate")
- Security type
- Exercise price if applicable
- Any other specific terms

IMPORTANT: Return ONLY valid JSON. Be extremely specific with vesting schedules.

{{
  "grants": [
    {{
      "stockholder": "Full Name",
      "shares": "exact_number",
      "grant_date": "YYYY-MM-DD",
      "vesting_start": "YYYY-MM-DD",
      "vesting_schedule_detailed": "exact vesting description from document",
      "security_type": "options/shares/warrant",
      "exercise_price": "price_if_applicable",
      "source_text": "relevant text from document showing these details"
    }}
  ],
  "document_type": "board_consent"
}}"""

        try:
            response = self.client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=2000,
                temperature=0,
                messages=[{"role": "user", "content": prompt}]
            )
            
            response_text = response.content[0].text.strip()
            
            # Find JSON boundaries more carefully
            json_start = -1
            json_end = -1
            brace_count = 0
            
            for i, char in enumerate(response_text):
                if char == '{':
                    if json_start == -1:
                        json_start = i
                    brace_count += 1
                elif char == '}':
                    brace_count -= 1
                    if brace_count == 0 and json_start != -1:
                        json_end = i + 1
                        break
            
            if json_start != -1 and json_end != -1:
                json_str = response_text[json_start:json_end]
                
                # Additional cleaning
                json_str = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', json_str)  # Remove control chars
                json_str = json_str.replace('\n', ' ').replace('\r', ' ')
                
                try:
                    return json.loads(json_str)
                except json.JSONDecodeError as e:
                    # If JSON still fails, return a basic structure
                    st.warning(f"JSON parsing failed for {doc_name}, using fallback extraction")
                    return {
                        "grants": [],
                        "document_type": "unknown",
                        "error_details": f"Could not parse JSON: {str(e)}",
                        "raw_response": response_text[:500]
                    }
            else:
                return {"error": "No valid JSON structure found", "raw_response": response_text[:500]}
                
        except Exception as e:
            return {"error": f"API call failed: {str(e)}"}

    def compare_with_claude(self, cap_df, legal_docs):
        if not self.client: return {"error": "No client"}
        
        prompt = f"""You are a lawyer conducting a capitalization table tie out. Compare EVERY SINGLE DETAIL between cap table and legal documents.

FIND ALL ISSUES - DO NOT STOP AT ONE TYPE. For each stockholder, check ALL of these and report EVERY discrepancy found:

1. SHARE COUNT MISMATCHES (exact numbers must match)
2. GRANT DATE MISMATCHES (exact dates must match - check character by character)
3. VESTING START DATE MISMATCHES (exact dates must match)
4. VESTING SCHEDULE MISMATCHES (exact wording must match - "1/48th monthly" ≠ "monthly vesting")
5. SECURITY TYPE MISMATCHES (options vs shares vs warrants)
6. EXERCISE PRICE MISMATCHES
7. MISSING ENTRIES (in cap table but not legal docs, or vice versa)
8. FORMATTING/DETAIL DIFFERENCES (generic vs specific descriptions)

IMPORTANT: Report MULTIPLE issues per stockholder if they exist. Do not limit to one issue per person.

Examples of multiple issues for same person:
- John Doe: Grant date wrong AND vesting schedule too generic AND share count off by 500
- Jane Smith: Missing from cap table entirely 
- Bob Johnson: Exercise price missing AND vesting start date wrong

CAPITALIZATION TABLE:
{cap_df.head(20).to_dict('records')}

LEGAL DOCUMENTS:
{legal_docs}

Return JSON with ALL discrepancies found (expect 10-20+ issues typically):
{{
  "discrepancies": [
    {{
      "stockholder": "Name",
      "discrepancy_type": "shares_mismatch/grant_date_mismatch/vesting_schedule_mismatch/vesting_start_mismatch/missing_detail/etc",
      "specific_issue": "Extremely specific description of what doesn't match",
      "cap_table_value": "exact value from cap table",
      "legal_document_value": "exact value from legal document", 
      "source_document": "document name",
      "severity": "high/medium/low",
      "legal_text_evidence": "exact text from legal document proving the correct value"
    }}
  ],
  "summary": {{
    "total_discrepancies": 0,
    "assessment": "detailed assessment noting that most cap tables have 10-20+ issues when compared forensically"
  }}
}}

Be exhaustive. Check every stockholder against every field. Report everything that doesn't match perfectly."""

        try:
            response = self.client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=4000,
                temperature=0,
                messages=[{"role": "user", "content": prompt}]
            )
            response_text = response.content[0].text.strip()
            
            # Same robust JSON extraction as analyze_with_claude
            json_start = -1
            json_end = -1
            brace_count = 0
            
            for i, char in enumerate(response_text):
                if char == '{':
                    if json_start == -1:
                        json_start = i
                    brace_count += 1
                elif char == '}':
                    brace_count -= 1
                    if brace_count == 0 and json_start != -1:
                        json_end = i + 1
                        break
            
            if json_start != -1 and json_end != -1:
                json_str = response_text[json_start:json_end]
                
                # Additional cleaning
                import re
                json_str = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', json_str)
                json_str = json_str.replace('\n', ' ').replace('\r', ' ')
                
                try:
                    return json.loads(json_str)
                except json.JSONDecodeError as e:
                    return {"error": f"JSON parsing error in comparison: {str(e)}", "raw_response": response_text[:500]}
            else:
                return {"error": "No valid JSON found in comparison response", "raw_response": response_text[:500]}
                
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
        st.subheader("📑 Analyzing Legal Documents")
        legal_analysis = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i, file in enumerate(legal_files):
            # Update status
            status_text.write(f"🔄 Processing document {i+1} of {len(legal_files)}: **{file.name}**")
            
            with st.spinner(f"Extracting text from {file.name}..."):
                text = auditor.extract_text(file)
            
            if text:
                with st.spinner(f"AI analyzing {file.name}..."):
                    analysis = auditor.analyze_with_claude(text, file.name)
                
                if "error" not in analysis:
                    legal_analysis.append(analysis)
                    st.success(f"✅ Successfully analyzed {file.name}")
                    
                    # Show grants found
                    grants_found = len(analysis.get('grants', []))
                    if grants_found > 0:
                        st.info(f"📋 Found {grants_found} grant(s) in {file.name}")
                else:
                    st.error(f"❌ Error analyzing {file.name}: {analysis['error']}")
            else:
                st.warning(f"⚠️ No text extracted from {file.name}")
            
            # Update progress
            progress_bar.progress((i + 1) / len(legal_files))
        
        # Clear status text when done
        status_text.empty()
        
        if legal_analysis:
            total_grants = sum(len(doc.get("grants", [])) for doc in legal_analysis)
            st.success(f"🎉 Analysis complete! Extracted **{total_grants} total grants** from **{len(legal_analysis)} documents**")
        # Compare
        if legal_analysis:
            st.subheader("🔍 Comparing Cap Table vs Legal Documents")
            with st.spinner("AI performing detailed discrepancy analysis..."):
                result = auditor.compare_with_claude(cap_df, legal_analysis)
            
            if "error" not in result:
                discrepancies = result.get("discrepancies", [])
                summary = result.get("summary", {})
                
                # Results header with better layout
                st.header("📋 Audit Results")
                
                # Summary metrics in prominent position
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Issues", len(discrepancies))
                with col2:
                    st.metric("🔴 High", summary.get("high_severity_count", 0))
                with col3:
                    st.metric("🟡 Medium", summary.get("medium_severity_count", 0))
                with col4:
                    st.metric("🟢 Low", summary.get("low_severity_count", 0))
                
                if summary.get("overall_assessment"):
                    st.info(f"**📊 Overall Assessment:** {summary['overall_assessment']}")
                
                # Discrepancies section
                if discrepancies:
                    st.subheader("🔍 Detailed Discrepancies")
                    
                    for i, d in enumerate(discrepancies, 1):
                        severity = {"high": "🔴", "medium": "🟡", "low": "🟢"}.get(d.get("severity"), "⚪")
                        
                        # Use a container to ensure proper scrolling
                        with st.container():
                            with st.expander(f"{severity} Issue {i}: {d.get('stockholder', 'Unknown')} - {d.get('discrepancy_type', 'Issue')}", expanded=d.get('severity') == 'high'):
                                
                                # Main discrepancy description
                                st.markdown(f"**🎯 Specific Issue:** {d.get('specific_issue', d.get('detailed_description', d.get('description', 'No details')))}")
                                
                                # Side-by-side comparison with exact values
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.markdown("**📊 Cap Table Shows:**")
                                    st.code(d.get('cap_table_value', 'N/A'), language=None)
                                
                                with col2:
                                    st.markdown("**📄 Legal Document Shows:**")
                                    st.code(d.get('legal_document_value', 'N/A'), language=None)
                                
                                # Legal evidence
                                if d.get('legal_text_evidence'):
                                    st.markdown("**🔍 Legal Document Evidence:**")
                                    st.code(d['legal_text_evidence'], language=None)
                                
                                # Source and correction needed
                                st.write(f"**📂 Source Document:** {d.get('source_document', 'N/A')}")
                                
                                if d.get('correction_required'):
                                    st.error(f"**✏️ Correction Needed:** {d['correction_required']}")
                                elif d.get('specific_issue'):
                                    st.warning(f"**⚠️ Action Required:** Review and update cap table to match legal document exactly")
                    
                    # Enhanced download with more details
                    st.subheader("📥 Download Report")
                    csv = pd.DataFrame(discrepancies).to_csv(index=False)
                    st.download_button("📊 Download Detailed Audit Report", csv, "detailed_audit_report.csv")
                    
                else:
                    st.success("🎉 No discrepancies found! Cap table matches legal documents perfectly.")
            else:
                st.error(f"❌ Comparison error: {result['error']}")
        else:
            st.warning("⚠️ No legal documents were successfully analyzed. Please check your files and try again.")

if __name__ == "__main__":
    main()
