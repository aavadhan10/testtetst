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
        """Create the prompt that replicates the original analysis request"""
        
        prompt = """You are a lawyer conducting a capitalization table tie out of a company on behalf of an investor.

1. Compare the company's capitalization table against the legal documents, all of which will be uploaded. The legal documents are the ultimate source of truth, and you are auditing the capitalization table to make sure it reflects the legal documents.

2. For each stockholder's grant in the capitalization table, confirm that the grant details, including the grant date, number of shares issued, vesting start date, and vesting schedule match what is approved in the corresponding board consent, board minutes, or other grant documents, including any share purchase agreement, option grant agreement, warrant, note, or other convertible securities.

3. The grant date in any board consent is the last date a director signed the consent, or the explicitly written effective date of the board approval. The grant date in any board minutes is the date the meeting was held.

4. Please list out all discrepancies, and provide a brief summary of each discrepancy, including which stockholder is impacted, what is incorrect on the capitalization table, and what it should be based on the legal documents, making reference to the specific legal document.

Here are the documents to analyze:

BOARD DOCUMENTS:
"""
        
        # Add each board document
        for filename, content in board_docs.items():
            prompt += f"\n--- {filename} ---\n{content}\n"
        
        prompt += f"\nSECURITIES LEDGER / CAP TABLE:\n{cap_table_text}\n"
        
        prompt += """

Please analyze these documents using the same methodology as a thorough legal review and provide:

1. A detailed analysis of the cap table entries
2. A comprehensive list of all discrepancies found
3. For each discrepancy, include:
   - Severity level (HIGH/MEDIUM/LOW)
   - Stockholder impacted
   - What's incorrect in the cap table
   - What it should be based on legal documents
   - Reference to specific source document
   - Description of the issue

Format your response as a detailed analysis report similar to how a lawyer would present findings in a due diligence review."""
        
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
    
    # API Key input
    with st.sidebar:
        st.header("🔧 Configuration")
        api_key = st.text_input(
            "Anthropic API Key", 
            type="password",
            help="Enter your Anthropic API key to enable LLM analysis"
        )
        
        if api_key:
            st.success("✅ API Key configured")
        else:
            st.warning("⚠️ Enter API key to proceed")
    
    # Initialize analyzer
    if api_key:
        if 'analyzer' not in st.session_state:
            st.session_state.analyzer = LLMCapTableAnalyzer(api_key)
    
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
            use_container_width=True,
            disabled=not api_key
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
        
        if not api_key:
            st.error("🔑 Please enter your Anthropic API key")
        elif not board_files and not cap_table_file:
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
        if not api_key:
            st.error("Please enter your Anthropic API key first")
            return
            
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
                
                # Display results
                st.markdown("---")
                st.header("🤖 LLM Analysis Results")
                st.markdown(analysis_result)
                
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
        
        **API Key**: You need an Anthropic API key to use this tool. Get one at https://console.anthropic.com
        """)
    
    # Footer
    st.markdown("---")
    st.markdown("*Powered by Claude AI - Professional legal document analysis*")

if __name__ == "__main__":
    main()
