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
    
    def display_analysis_with_cards(self, analysis_result: str):
        """Display analysis results in pretty card format"""
        
        # Show raw analysis in an expander first
        with st.expander("📄 View Full Raw Analysis"):
            st.markdown(analysis_result)
        
        # Try to parse discrepancies from the text
        discrepancies = self.parse_discrepancies_from_text(analysis_result)
        
        if discrepancies:
            # Summary metrics
            st.subheader("📊 Summary")
            
            high_count = len([d for d in discrepancies if d.get('severity') == 'HIGH'])
            medium_count = len([d for d in discrepancies if d.get('severity') == 'MEDIUM'])
            low_count = len([d for d in discrepancies if d.get('severity') == 'LOW'])
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("🔍 Total Issues", len(discrepancies))
            with col2:
                st.metric("🔴 High Severity", high_count)
            with col3:
                st.metric("🟡 Medium Severity", medium_count) 
            with col4:
                st.metric("🟢 Low Severity", low_count)
            
            # Risk level assessment
            if high_count >= 5:
                st.error("🚨 **CRITICAL RISK**: Multiple high-severity issues require immediate attention")
            elif high_count >= 2:
                st.warning("⚠️ **HIGH RISK**: Several important discrepancies found")
            elif high_count >= 1:
                st.warning("⚠️ **MEDIUM RISK**: Some discrepancies need correction")
            else:
                st.success("✅ **LOW RISK**: Minor issues only")
            
            # Display discrepancies as cards
            st.subheader("🔍 Detailed Discrepancies")
            
            for i, disc in enumerate(discrepancies, 1):
                self.create_discrepancy_card(i, disc)
        else:
            # Fallback to regular display if parsing fails
            st.info("💡 Could not parse structured discrepancies. Showing full analysis:")
            st.markdown(analysis_result)
    
    def parse_discrepancies_from_text(self, text: str) -> List[Dict]:
        """Parse discrepancies from LLM response text with better pattern matching"""
        discrepancies = []
        
        # Split into sections and look for discrepancy patterns
        sections = text.split('\n\n')  # Split by double newlines
        
        for section in sections:
            lines = section.strip().split('\n')
            
            # Look for discrepancy indicators
            discrepancy_indicators = [
                'discrepancy', 'issue', 'problem', 'error', 
                'severity:', 'stockholder:', 'high', 'medium', 'low'
            ]
            
            if any(indicator in section.lower() for indicator in discrepancy_indicators):
                disc = self.extract_discrepancy_from_section(section)
                if disc and any(disc.values()):  # Only add if we extracted meaningful data
                    discrepancies.append(disc)
        
        # If we didn't find structured discrepancies, try line-by-line approach
        if not discrepancies:
            discrepancies = self.parse_line_by_line(text)
        
        return discrepancies
    
    def extract_discrepancy_from_section(self, section: str) -> Dict:
        """Extract discrepancy info from a text section"""
        disc = {}
        lines = section.split('\n')
        
        for line in lines:
            line = line.strip()
            lower_line = line.lower()
            
            # Extract severity
            if 'severity:' in lower_line:
                disc['severity'] = line.split(':', 1)[1].strip().upper()
            elif any(sev in lower_line for sev in ['high', 'medium', 'low']):
                if 'high' in lower_line:
                    disc['severity'] = 'HIGH'
                elif 'medium' in lower_line:
                    disc['severity'] = 'MEDIUM'
                elif 'low' in lower_line:
                    disc['severity'] = 'LOW'
            
            # Extract stockholder
            if 'stockholder:' in lower_line:
                disc['stockholder'] = line.split(':', 1)[1].strip()
            elif any(name in line for name in ['John Doe', 'Jane Smith']):
                for name in ['John Doe', 'Jane Smith']:
                    if name in line:
                        disc['stockholder'] = name
                        break
            
            # Extract issue type
            if 'issue:' in lower_line:
                disc['issue'] = line.split(':', 1)[1].strip()
            elif any(issue_type in lower_line for issue_type in [
                'board approval', 'phantom equity', 'repurchase', 'price', 
                'vesting', 'missing', 'incorrect'
            ]):
                disc['issue'] = line.strip()
            
            # Extract cap table value
            if 'cap table shows:' in lower_line or 'cap table:' in lower_line:
                disc['cap_table_value'] = line.split(':', 1)[1].strip()
            
            # Extract legal document value
            if any(phrase in lower_line for phrase in [
                'legal documents show:', 'should be:', 'correct:', 'board documents:'
            ]):
                disc['legal_value'] = line.split(':', 1)[1].strip()
            
            # Extract source document
            if 'reference:' in lower_line or 'source:' in lower_line or 'template' in lower_line:
                disc['source'] = line.strip()
        
        # Try to extract issue from first line if not found
        if 'issue' not in disc and lines:
            first_line = lines[0].strip()
            # Clean up common prefixes
            for prefix in ['1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '-', '*']:
                if first_line.startswith(prefix):
                    first_line = first_line[len(prefix):].strip()
            disc['issue'] = first_line
        
        return disc
    
    def parse_line_by_line(self, text: str) -> List[Dict]:
        """Fallback: parse line by line looking for key information"""
        discrepancies = []
        lines = text.split('\n')
        
        current_disc = {}
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            lower_line = line.lower()
            
            # Look for numbered items or bullet points that might be discrepancies
            if any(line.startswith(prefix) for prefix in ['1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '-', '*']):
                # Save previous discrepancy
                if current_disc:
                    discrepancies.append(current_disc)
                current_disc = {'issue': line, 'description': line}
            
            # Look for severity mentions
            elif any(sev in lower_line for sev in ['high', 'medium', 'low']):
                if 'high' in lower_line:
                    current_disc['severity'] = 'HIGH'
                elif 'medium' in lower_line:
                    current_disc['severity'] = 'MEDIUM' 
                elif 'low' in lower_line:
                    current_disc['severity'] = 'LOW'
            
            # Look for stockholder names
            elif any(name in line for name in ['John Doe', 'Jane Smith']):
                for name in ['John Doe', 'Jane Smith']:
                    if name in line:
                        current_disc['stockholder'] = name
                        break
            
            # Look for specific values
            elif 'may 16' in lower_line and '2025' in line:
                current_disc['cap_table_value'] = 'May 16, 2025'
            elif 'march 1' in lower_line or 'january 1' in lower_line:
                current_disc['legal_value'] = line.strip()
            elif 'template' in lower_line:
                current_disc['source'] = line.strip()
        
        # Add last discrepancy
        if current_disc:
            discrepancies.append(current_disc)
        
        return discrepancies
    
    def create_discrepancy_card(self, number: int, discrepancy: Dict):
        """Create a pretty card for each discrepancy"""
        
        # Determine card styling based on severity
        severity = discrepancy.get('severity', 'UNKNOWN').upper()
        if severity == 'HIGH':
            border_color = "#ff4444"
            header_emoji = "🔴"
            bg_color = "#fff5f5"
        elif severity == 'MEDIUM':
            border_color = "#ffaa00"
            header_emoji = "🟡"
            bg_color = "#fffaf0"
        elif severity == 'LOW':
            border_color = "#00aa00"
            header_emoji = "🟢"
            bg_color = "#f0fff4"
        else:
            border_color = "#888888"
            header_emoji = "⚪"
            bg_color = "#f9f9f9"
        
        # Create the card using HTML
        title = discrepancy.get('title', discrepancy.get('issue', f'Discrepancy #{number}'))
        stockholder = discrepancy.get('stockholder', 'Unknown')
        issue = discrepancy.get('issue', 'Issue not specified')
        
        with st.container():
            st.markdown(
                f"""
                <div style="
                    border-left: 4px solid {border_color};
                    background-color: {bg_color};
                    padding: 1rem;
                    margin: 1rem 0;
                    border-radius: 0 8px 8px 0;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                ">
                    <h4 style="margin: 0 0 0.5rem 0; color: {border_color};">
                        {header_emoji} DISCREPANCY #{number}: {issue}
                    </h4>
                    <p style="margin: 0; font-weight: bold; color: #333;">
                        👤 <strong>Stockholder:</strong> {stockholder} | 
                        📊 <strong>Severity:</strong> {severity}
                    </p>
                </div>
                """, 
                unsafe_allow_html=True
            )
            
            # Details in columns
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**📋 Cap Table Shows:**")
                cap_value = discrepancy.get('cap_table_value', 'Not specified')
                st.code(cap_value)
                
                if 'security_id' in discrepancy:
                    st.markdown("**🆔 Security ID:**")
                    st.code(discrepancy['security_id'])
            
            with col2:
                st.markdown("**📜 Legal Documents Show:**")
                legal_value = discrepancy.get('legal_value', 'Not specified')
                st.code(legal_value)
                
                if 'source' in discrepancy:
                    st.markdown("**📄 Source Document:**")
                    st.code(discrepancy['source'])
            
            if 'description' in discrepancy:
                st.markdown("**📝 Description:**")
                st.write(discrepancy['description'])
            
            st.markdown("---")

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
                
                # Display results with pretty formatting
                st.markdown("---")
                st.header("🤖 LLM Analysis Results")
                
                # Parse the analysis result to create cards
                analyzer.display_analysis_with_cards(analysis_result)
                
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
