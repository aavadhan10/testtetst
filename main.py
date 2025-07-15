import streamlit as st
import pandas as pd
import openpyxl
from docx import Document
import io
from datetime import datetime
import base64
import anthropic
import os
from typing import List, Dict

# Page configuration
st.set_page_config(
    page_title="Cap Table Tie-Out Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

class LLMCapTableAnalyzer:
    def __init__(self, api_key: str):
        self.client = anthropic.Anthropic(api_key=api_key)
        self.uploaded_files = {}
        
    def read_docx_content(self, file_content: bytes, filename: str) -> str:
        """Read DOCX content and return as plain text"""
        try:
            doc = Document(io.BytesIO(file_content))
            full_text = []
            for paragraph in doc.paragraphs:
                full_text.append(paragraph.text)
            return '\n'.join(full_text)
        except Exception as e:
            st.error(f"Error reading {filename}: {str(e)}")
            return ""
    
    def excel_to_text_preview(self, file_content: bytes, filename: str) -> str:
        """Convert Excel to text preview for LLM analysis"""
        try:
            df = pd.read_excel(io.BytesIO(file_content), engine='openpyxl', header=None)
            
            # Create text representation
            text_preview = f"Excel file: {filename}\n\n"
            text_preview += "Raw data structure:\n"
            
            for i in range(min(15, len(df))):  # First 15 rows
                row_data = df.iloc[i].tolist()
                # Clean up the row data
                cleaned_row = []
                for cell in row_data:
                    if pd.isna(cell):
                        cleaned_row.append("")
                    else:
                        cleaned_row.append(str(cell))
                text_preview += f"Row {i + 1}: {cleaned_row}\n"
            
            return text_preview
        except Exception as e:
            return f"Error reading Excel file {filename}: {str(e)}"
    
    def create_analysis_prompt(self, board_docs: Dict[str, str], cap_table_text: str) -> str:
        """Create the enhanced prompt that catches all discrepancies"""
        
        prompt = """You are a lawyer conducting a comprehensive capitalization table tie out of a company on behalf of an investor. You must be extremely thorough and catch EVERY discrepancy, no matter how small.

CRITICAL INSTRUCTIONS:
1. Compare the company's capitalization table against the legal documents. The legal documents are the ultimate source of truth.

2. For EACH stockholder's grant in the capitalization table, verify:
   - Grant date matches board approval date
   - Number of shares issued matches board approval
   - Price per share is correct (calculate from cost basis ÷ shares)
   - Vesting start date matches board documents
   - Vesting schedule matches board documents
   - Issue date matches board approval date
   - Board approval date is accurate

3. PHANTOM EQUITY DETECTION:
   - Flag ANY cap table entry that lacks supporting board documentation
   - Every grant must have a corresponding board consent, resolution, or legal document
   - If you cannot find board approval for a grant, it's a HIGH severity phantom equity issue

4. VESTING SCHEDULE VERIFICATION:
   - Check if vesting schedule format matches between cap table and board documents
   - Look for discrepancies like "monthly" vs "annual" vesting
   - Flag if board says "1/48th monthly" but cap table shows different vesting frequency
   - Verify vesting schedule descriptions match exactly (e.g., "1/48 monthly" vs "25% annually")
   - Flag vesting schedule mismatches as HIGH severity

5. REPURCHASE/CANCELLATION VERIFICATION:
   - Check if cap table reflects any share repurchases or cancellations from board documents
   - Verify remaining share counts after repurchases
   - Check if repurchase pricing matches original grant pricing

6. GRANULAR ANALYSIS REQUIRED:
   - List each discrepancy separately (don't group multiple issues)
   - Check dates, quantities, pricing, and math independently
   - Be as detailed as the most thorough legal review

7. The grant date in any board consent is the last date a director signed the consent, or the explicitly written effective date.

Here are the documents to analyze:

BOARD DOCUMENTS:
"""
        
        # Add each board document
        for filename, content in board_docs.items():
            prompt += f"\n--- {filename} ---\n{content}\n"
        
        prompt += f"\nSECURITIES LEDGER / CAP TABLE:\n{cap_table_text}\n"
        
        prompt += """

ANALYSIS REQUIREMENTS:

1. DOCUMENT MAPPING: First, create a list of all board-approved grants from the legal documents
2. CAP TABLE REVIEW: Analyze each cap table entry against this list
3. PHANTOM EQUITY: Identify entries with no board support
4. MATHEMATICAL VERIFICATION: Calculate and verify all vesting amounts
5. DISCREPANCY IDENTIFICATION: List every single discrepancy found

For each discrepancy, provide:
- Discrepancy #[number]
- Severity: HIGH/MEDIUM/LOW
- Stockholder: [name]
- Security ID: [ID from cap table]
- Issue: [brief title]
- Cap Table Shows: [specific value]
- Legal Documents Show: [what it should be]
- Description: [detailed explanation]
- Source Document: [specific filename]

SEVERITY GUIDELINES:
- HIGH: Phantom equity, wrong share counts, incorrect pricing, missing repurchases, wrong vesting calculations
- MEDIUM: Date discrepancies, documentation gaps
- LOW: Minor formatting or non-material issues

Be extremely thorough - this is for investor due diligence and every discrepancy matters. Aim to find 8-12 discrepancies if the cap table has significant issues."""
        
        return prompt
    
    def analyze_with_llm(self, board_docs: Dict[str, str], cap_table_text: str) -> str:
        """Send documents to LLM for analysis"""
        
        prompt = self.create_analysis_prompt(board_docs, cap_table_text)
        
        try:
            response = self.client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=4000,
                temperature=0,
                messages=[{
                    "role": "user", 
                    "content": prompt
                }]
            )
            
            return response.content[0].text
            
        except Exception as e:
            return f"Error analyzing documents: {str(e)}"

def main():
    st.title("📊 Cap Table Tie-Out Analysis")
    st.markdown("*LLM-powered analysis replicating expert legal review*")
    
    # Initialize analyzer with API key from secrets
    if 'analyzer' not in st.session_state:
        try:
            api_key = st.secrets["ANTHROPIC_API_KEY"]
            st.session_state.analyzer = LLMCapTableAnalyzer(api_key)
        except KeyError:
            st.error("❌ Anthropic API key not found in Streamlit secrets. Please configure ANTHROPIC_API_KEY in your .streamlit/secrets.toml file")
            st.stop()
    
    # File upload section
    with st.sidebar:
        st.header("📁 Upload Documents")
        
        # Board documents upload
        st.subheader("Board Documents")
        board_files = st.file_uploader(
            "Upload board consents, minutes, and legal docs",
            type=['docx', 'doc'],
            accept_multiple_files=True,
            key="board_docs",
            help="Upload DOCX files containing board resolutions, consents, and other legal documents"
        )
        
        # Securities ledger upload
        st.subheader("Securities Ledger")
        cap_table_file = st.file_uploader(
            "Upload cap table (Excel format)",
            type=['xlsx', 'xls'],
            key="cap_table",
            help="Upload Excel file containing the company's capitalization table"
        )
        
        # Analysis button
        st.markdown("---")
        run_analysis = st.button(
            "🔍 Run LLM Tie-Out Analysis", 
            type="primary", 
            use_container_width=True
        )
    
    # Main content area
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📋 Uploaded Documents")
        
        if board_files:
            st.write("**Board Documents:**")
            for file in board_files:
                st.write(f"✅ {file.name}")
                
            # Show preview of first document
            if st.checkbox("Show document preview"):
                first_file = board_files[0]
                first_file.seek(0)
                analyzer = st.session_state.get('analyzer')
                if analyzer:
                    content = analyzer.read_docx_content(first_file.read(), first_file.name)
                    st.text_area("Document content preview:", content[:1000] + "...", height=200)
        else:
            st.info("No board documents uploaded yet")
        
        if cap_table_file:
            st.write("**Securities Ledger:**")
            st.write(f"✅ {cap_table_file.name}")
            
            # Show preview of cap table
            if st.checkbox("Show cap table preview"):
                cap_table_file.seek(0)
                try:
                    df_preview = pd.read_excel(io.BytesIO(cap_table_file.read()), engine='openpyxl')
                    st.dataframe(df_preview.head(10))
                except Exception as e:
                    st.error(f"Error previewing cap table: {e}")
        else:
            st.info("No securities ledger uploaded yet")
    
    with col2:
        st.subheader("⚙️ Analysis Status")
        
        if not board_files and not cap_table_file:
            st.warning("📄 Please upload documents to begin analysis")
        elif not board_files:
            st.warning("📋 Please upload board documents")
        elif not cap_table_file:
            st.warning("📊 Please upload securities ledger")
        else:
            st.success("✅ Ready for LLM analysis!")
            st.info("Click 'Run LLM Tie-Out Analysis' to start")
    
    # Run analysis when button is clicked
    if run_analysis:
        if not board_files or not cap_table_file:
            st.error("Please upload both board documents and securities ledger before running analysis")
            return
        
        with st.spinner("🤖 LLM is analyzing your documents... This may take 30-60 seconds"):
            try:
                analyzer = st.session_state.analyzer
                
                # Process board documents
                board_docs = {}
                for file in board_files:
                    file.seek(0)  # Reset file pointer
                    content = analyzer.read_docx_content(file.read(), file.name)
                    board_docs[file.name] = content
                
                # Process cap table
                cap_table_file.seek(0)
                cap_table_text = analyzer.excel_to_text_preview(cap_table_file.read(), cap_table_file.name)
                
                # Send to LLM for analysis
                analysis_result = analyzer.analyze_with_llm(board_docs, cap_table_text)
                
                # Display results - clean and simple like the original Claude analysis
                st.markdown("---")
                st.header("🤖 LLM Analysis Results")
                
                # Display the analysis in a clean format
                st.markdown(analysis_result)
                
                # Add some basic metrics if we can extract them
                high_count = analysis_result.lower().count('high')
                medium_count = analysis_result.lower().count('medium') 
                low_count = analysis_result.lower().count('low')
                
                if high_count > 0 or medium_count > 0 or low_count > 0:
                    st.markdown("---")
                    st.subheader("📊 Quick Summary")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("🔴 High Issues", high_count)
                    with col2:
                        st.metric("🟡 Medium Issues", medium_count)
                    with col3:
                        st.metric("🟢 Low Issues", low_count)
                
                # Create downloadable report
                st.markdown("---")
                st.subheader("📤 Download Report")
                
                # Create downloadable text file
                report_filename = f"cap_table_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                
                report_content = f"""Cap Table Tie-Out Analysis Report
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Board Documents Analyzed:
{', '.join(board_docs.keys())}

Cap Table File: {cap_table_file.name}

ANALYSIS RESULTS:
{analysis_result}
"""
                
                st.download_button(
                    label="📄 Download Analysis Report",
                    data=report_content,
                    file_name=report_filename,
                    mime="text/plain"
                )
                
            except Exception as e:
                st.error(f"Error during analysis: {str(e)}")
                st.error("Please check your API key and try again")
    
    # Information section
    with st.expander("ℹ️ How this works"):
        st.markdown("""
        **This tool replicates the exact process of expert legal document analysis:**
        
        1. **Document Upload**: You upload board documents and cap table (just like dragging files to Claude)
        2. **LLM Processing**: The documents are sent to Claude (same AI that did the original analysis)
        3. **Expert Analysis**: Claude performs the same thorough legal review methodology
        4. **Detailed Report**: You get the same quality discrepancy analysis and recommendations
        
        **What the LLM analyzes:**
        - Compares cap table entries against board approvals
        - Identifies missing documentation
        - Checks share quantities, pricing, and dates
        - Detects phantom equity grants
        - Finds missing transactions (repurchases, etc.)
        - Provides severity assessment and recommendations
        
        **API Key**: Configured via Streamlit secrets (ANTHROPIC_API_KEY in .streamlit/secrets.toml)
        """)
    
    # Footer
    st.markdown("---")
    st.markdown("*Powered by Claude AI - Professional legal document analysis*")

if __name__ == "__main__":
    main()
