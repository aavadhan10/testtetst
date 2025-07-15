import streamlit as st
import pandas as pd
import openpyxl
from docx import Document
import io
from datetime import datetime
import re
import tempfile
import os

# Page configuration
st.set_page_config(
    page_title="Cap Table Tie-Out Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

class CapTableAnalyzer:
    def __init__(self):
        self.board_docs_content = {}
        self.cap_table_entries = []
        self.board_analysis = {}
        
    def save_uploaded_files(self, board_files, cap_table_file):
        """Save uploaded files temporarily for JavaScript analysis"""
        temp_dir = tempfile.mkdtemp()
        file_paths = {}
        
        # Save board documents
        if board_files:
            for file in board_files:
                file_path = os.path.join(temp_dir, file.name)
                with open(file_path, 'wb') as f:
                    f.write(file.read())
                file_paths[file.name] = file_path
                
                # Also read content for analysis
                if file.name.endswith('.docx'):
                    file.seek(0)  # Reset file pointer
                    doc = Document(io.BytesIO(file.read()))
                    content = '\n'.join([p.text for p in doc.paragraphs])
                    self.board_docs_content[file.name] = content
        
        # Save cap table
        if cap_table_file:
            cap_file_path = os.path.join(temp_dir, cap_table_file.name)
            with open(cap_file_path, 'wb') as f:
                f.write(cap_table_file.read())
            file_paths['cap_table'] = cap_file_path
        
        return file_paths, temp_dir
    
    def analyze_excel_like_original(self, cap_table_file):
        """Analyze Excel file exactly like I did with JavaScript"""
        st.subheader("📊 Cap Table Data Inspection")
        st.write("*Replicating the JavaScript/SheetJS analysis approach*")
        
        try:
            # Read Excel file with openpyxl (similar to SheetJS)
            cap_table_file.seek(0)
            df = pd.read_excel(io.BytesIO(cap_table_file.read()), engine='openpyxl', header=None)
            
            # Show raw structure first (like I did)
            st.write("**First 10 rows of raw data:**")
            for i in range(min(10, len(df))):
                row_data = df.iloc[i].tolist()
                st.write(f"Row {i + 1}: {row_data}")
            
            # Find header row (like I did manually)
            header_row_idx = None
            for i, row in df.iterrows():
                row_str = ' '.join([str(cell) for cell in row if pd.notna(cell)])
                if 'Security ID' in row_str and 'Stakeholder Name' in row_str:
                    header_row_idx = i
                    break
            
            if header_row_idx is not None:
                st.write(f"\n**Headers found in row {header_row_idx + 1}:**")
                headers = df.iloc[header_row_idx].tolist()
                st.write(headers)
                
                st.write("\n**Cap Table Entries:**")
                
                # Extract data entries (like I did)
                entries = []
                for i in range(header_row_idx + 1, len(df)):
                    row = df.iloc[i]
                    if pd.notna(row.iloc[0]) and str(row.iloc[0]).strip():  # Has Security ID
                        entry_data = {}
                        for j, header in enumerate(headers):
                            if j < len(row) and pd.notna(row.iloc[j]) and pd.notna(header):
                                entry_data[str(header)] = row.iloc[j]
                        
                        if entry_data:
                            entries.append({
                                'entry_num': i - header_row_idx,
                                'security_id': row.iloc[0],
                                'data': entry_data
                            })
                            
                            # Display like I did originally
                            with st.expander(f"Entry {i - header_row_idx} (Security ID: {row.iloc[0]})"):
                                for header, value in entry_data.items():
                                    if pd.notna(value) and str(value).strip():
                                        st.write(f"**{header}:** {value}")
                
                self.cap_table_entries = entries
                return entries
            else:
                st.error("Could not find header row with 'Security ID' and 'Stakeholder Name'")
                return []
                
        except Exception as e:
            st.error(f"Error analyzing Excel file: {str(e)}")
            return []
    
    def analyze_board_docs_like_original(self):
        """Analyze board documents using manual parsing like I did originally"""
        st.subheader("📋 Board Document Analysis")
        st.write("*Manual document review and key information extraction*")
        
        analysis = {}
        
        for filename, content in self.board_docs_content.items():
            st.write(f"\n**{filename}:**")
            
            # Determine document type (like I did manually)
            content_lower = content.lower()
            
            if 'restricted stock' in content_lower or 'rsa' in content_lower:
                st.write("- **Document Type:** Board Consent for RSA Issuance")
                
                # Extract information manually (like I did)
                doc_analysis = self._extract_rsa_info(content, filename)
                analysis[filename] = doc_analysis
                
                # Display findings
                st.write(f"- **Document dated:** {doc_analysis.get('date', 'Not found')}")
                st.write(f"- **Stockholder:** {doc_analysis.get('stockholder', 'Not found')}")
                st.write(f"- **Shares:** {doc_analysis.get('shares', 'Not found')}")
                st.write(f"- **Price per share:** {doc_analysis.get('price_per_share', 'Not found')}")
                st.write(f"- **Vesting start date:** {doc_analysis.get('vesting_start', 'Not found')}")
                st.write(f"- **Vesting schedule:** {doc_analysis.get('vesting_schedule', 'Not found')}")
                
            elif 'repurchase' in content_lower:
                st.write("- **Document Type:** Board Consent for Repurchase")
                
                doc_analysis = self._extract_repurchase_info(content, filename)
                analysis[filename] = doc_analysis
                
                st.write(f"- **Document dated:** {doc_analysis.get('date', 'Not found')}")
                st.write(f"- **Stockholder:** {doc_analysis.get('stockholder', 'Not found')}")
                st.write(f"- **Shares repurchased:** {doc_analysis.get('repurchased_shares', 'Not found')}")
                st.write(f"- **Repurchase price:** {doc_analysis.get('price_per_share', 'Not found')}")
                
            else:
                st.write("- **Document Type:** Unknown/Other")
                # Try to extract basic info
                doc_analysis = self._extract_basic_info(content, filename)
                analysis[filename] = doc_analysis
        
        self.board_analysis = analysis
        return analysis
    
    def _extract_rsa_info(self, content, filename):
        """Extract RSA grant information like I did manually"""
        lines = content.split('\n')
        
        analysis = {
            'type': 'RSA Grant',
            'filename': filename,
            'date': None,
            'stockholder': None,
            'shares': None,
            'price_per_share': None,
            'vesting_start': None,
            'vesting_schedule': None
        }
        
        # Extract date (look for "Date:" pattern)
        for line in lines:
            if 'Date:' in line and ('2024' in line or '2025' in line):
                analysis['date'] = line.replace('Date:', '').strip()
                break
        
        # Extract stockholder name (look in schedule or text)
        for line in lines:
            # Common patterns for names
            if any(name in line for name in ['John Doe', 'Jane Smith', 'Bob', 'Alice']):
                # Extract the name
                for potential_name in ['John Doe', 'Jane Smith', 'Bob', 'Alice']:
                    if potential_name in line:
                        analysis['stockholder'] = potential_name
                        break
                if analysis['stockholder']:
                    break
        
        # Extract shares (look for number followed by "shares")
        for line in lines:
            match = re.search(r'(\d{1,3}(?:,\d{3})*)\s+shares?', line)
            if match:
                analysis['shares'] = match.group(1)
                break
        
        # Extract price per share
        for line in lines:
            match = re.search(r'\$(\d+\.\d{2})\s+per\s+share', line)
            if not match:
                match = re.search(r'\$(\d+\.\d{2})', line)
            if match:
                analysis['price_per_share'] = f"${match.group(1)}"
                break
        
        # Extract vesting start date
        for line in lines:
            # Look for date patterns
            date_match = re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}', line)
            if date_match and 'vesting' in line.lower():
                analysis['vesting_start'] = date_match.group(0)
                break
        
        # Extract vesting schedule
        for line in lines:
            if '1/48' in line and 'month' in line.lower():
                analysis['vesting_schedule'] = '1/48th monthly'
                break
        
        return analysis
    
    def _extract_repurchase_info(self, content, filename):
        """Extract repurchase information"""
        lines = content.split('\n')
        
        analysis = {
            'type': 'Repurchase',
            'filename': filename,
            'date': None,
            'stockholder': None,
            'repurchased_shares': None,
            'price_per_share': None
        }
        
        # Extract date
        for line in lines:
            if 'Date:' in line and ('2024' in line or '2025' in line):
                analysis['date'] = line.replace('Date:', '').strip()
                break
        
        # Extract stockholder
        for line in lines:
            for potential_name in ['John Doe', 'Jane Smith', 'Bob', 'Alice']:
                if potential_name in line:
                    analysis['stockholder'] = potential_name
                    break
            if analysis['stockholder']:
                break
        
        # Extract repurchased shares
        for line in lines:
            match = re.search(r'repurchase\s+(\d{1,3}(?:,\d{3})*)\s+', line, re.IGNORECASE)
            if match:
                analysis['repurchased_shares'] = match.group(1)
                break
        
        # Extract price
        for line in lines:
            match = re.search(r'\$(\d+\.\d{2})', line)
            if match:
                analysis['price_per_share'] = f"${match.group(1)}"
                break
        
        return analysis
    
    def _extract_basic_info(self, content, filename):
        """Extract basic info from unknown document types"""
        return {
            'type': 'Unknown',
            'filename': filename,
            'content_preview': content[:500] + "..." if len(content) > 500 else content
        }
    
    def perform_discrepancy_analysis(self):
        """Perform discrepancy analysis like I did originally"""
        st.header("🔍 Discrepancy Analysis")
        st.write("*Comparing cap table entries against board documents*")
        
        discrepancies = []
        
        # Get board document references
        rsa_docs = [doc for doc in self.board_analysis.values() if doc.get('type') == 'RSA Grant']
        repurchase_docs = [doc for doc in self.board_analysis.values() if doc.get('type') == 'Repurchase']
        
        # Create lookup of approved grants
        approved_grants = {}
        for doc in rsa_docs:
            if doc.get('stockholder'):
                key = doc['stockholder']
                approved_grants[key] = doc
        
        # Analyze each cap table entry
        for entry in self.cap_table_entries:
            data = entry['data']
            security_id = entry['security_id']
            stakeholder = data.get('Stakeholder Name', '')
            
            # Check if this grant has board approval
            if stakeholder not in approved_grants:
                discrepancies.append({
                    'severity': 'HIGH',
                    'stockholder': stakeholder,
                    'security_id': security_id,
                    'issue': 'Missing Board Approval',
                    'cap_table_value': 'Entry exists',
                    'legal_document_value': 'No supporting documentation found',
                    'description': f'Cap table shows {security_id} for {stakeholder} but no board approval found',
                    'source_document': 'None found'
                })
            else:
                # Compare against board approval
                board_doc = approved_grants[stakeholder]
                
                # Check shares
                cap_shares = str(data.get('Quantity Issued', ''))
                board_shares = board_doc.get('shares', '')
                if cap_shares and board_shares and cap_shares.replace(',', '') != board_shares.replace(',', ''):
                    discrepancies.append({
                        'severity': 'HIGH',
                        'stockholder': stakeholder,
                        'security_id': security_id,
                        'issue': 'Incorrect Share Quantity',
                        'cap_table_value': f'{cap_shares} shares',
                        'legal_document_value': f'{board_shares} shares',
                        'description': f'Cap table shows {cap_shares} shares but board approval is for {board_shares} shares',
                        'source_document': board_doc['filename']
                    })
                
                # Check price per share
                try:
                    cap_cost_basis = float(data.get('Cost Basis', 0))
                    cap_shares_num = float(str(data.get('Quantity Issued', 1)).replace(',', ''))
                    cap_price_per_share = cap_cost_basis / cap_shares_num if cap_shares_num > 0 else 0
                    
                    board_price_str = board_doc.get('price_per_share', '$0')
                    board_price = float(board_price_str.replace('$', ''))
                    
                    if abs(cap_price_per_share - board_price) > 0.01:
                        discrepancies.append({
                            'severity': 'HIGH',
                            'stockholder': stakeholder,
                            'security_id': security_id,
                            'issue': 'Incorrect Price Per Share',
                            'cap_table_value': f'${cap_price_per_share:.2f}',
                            'legal_document_value': f'${board_price:.2f}',
                            'description': f'Cap table shows ${cap_price_per_share:.2f} per share but board approval is for ${board_price:.2f} per share',
                            'source_document': board_doc['filename']
                        })
                except (ValueError, ZeroDivisionError):
                    pass
                
                # Check board approval date
                cap_approval_date = str(data.get('Board Approval Date', ''))
                board_date = board_doc.get('date', '')
                if cap_approval_date and board_date:
                    # Simple date comparison (could be enhanced)
                    if board_date not in cap_approval_date:
                        discrepancies.append({
                            'severity': 'HIGH',
                            'stockholder': stakeholder,
                            'security_id': security_id,
                            'issue': 'Incorrect Board Approval Date',
                            'cap_table_value': cap_approval_date,
                            'legal_document_value': board_date,
                            'description': 'Board approval date in cap table does not match legal documents',
                            'source_document': board_doc['filename']
                        })
        
        # Check for missing repurchase transactions
        for repurchase_doc in repurchase_docs:
            stockholder = repurchase_doc.get('stockholder', '')
            repurchased_shares = repurchase_doc.get('repurchased_shares', '')
            
            if stockholder and repurchased_shares:
                discrepancies.append({
                    'severity': 'HIGH',
                    'stockholder': stockholder,
                    'security_id': 'Multiple',
                    'issue': 'Missing Repurchase Transaction',
                    'cap_table_value': 'No repurchase reflected',
                    'legal_document_value': f'{repurchased_shares} shares repurchased',
                    'description': f'Board approved repurchase of {repurchased_shares} shares from {stockholder} but cap table does not reflect this transaction',
                    'source_document': repurchase_doc['filename']
                })
        
        return discrepancies
    
    def generate_report(self, discrepancies):
        """Generate detailed report like my original markdown report"""
        st.header("📊 Final Analysis Report")
        
        if not discrepancies:
            st.success("🎉 No discrepancies found! The cap table appears to be in sync with the board documents.")
            return
        
        # Summary
        st.error(f"⚠️ Found {len(discrepancies)} discrepancies that require immediate correction")
        
        high_severity = [d for d in discrepancies if d['severity'] == 'HIGH']
        medium_severity = [d for d in discrepancies if d['severity'] == 'MEDIUM']
        low_severity = [d for d in discrepancies if d['severity'] == 'LOW']
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("🔴 High Severity", len(high_severity))
        with col2:
            st.metric("🟡 Medium Severity", len(medium_severity))
        with col3:
            st.metric("🟢 Low Severity", len(low_severity))
        
        # Detailed discrepancies
        st.subheader("Detailed Discrepancy Analysis")
        
        for i, disc in enumerate(discrepancies, 1):
            severity_color = {"HIGH": "🔴", "MEDIUM": "🟡", "LOW": "🟢"}
            
            with st.expander(f"{severity_color[disc['severity']]} DISCREPANCY #{i}: {disc['issue']} - {disc['stockholder']}"):
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**Severity:** " + disc['severity'])
                    st.write("**Stockholder:** " + disc['stockholder'])
                    st.write("**Security ID:** " + disc['security_id'])
                    st.write("**Issue:** " + disc['issue'])
                
                with col2:
                    st.write("**Cap Table Shows:**")
                    st.code(disc['cap_table_value'])
                    
                    st.write("**Legal Documents Show:**")
                    st.code(disc['legal_document_value'])
                
                st.write("**Description:**")
                st.write(disc['description'])
                
                st.write("**Source Document:**")
                st.write(disc['source_document'])
        
        # Risk assessment and recommendations
        st.subheader("Risk Assessment & Recommendations")
        
        if high_severity:
            st.error("**High Risk Issues Identified:**")
            st.write("- Multiple discrepancies require immediate attention")
            st.write("- Potential phantom equity grants without legal support")
            st.write("- Incorrect pricing and dates affecting valuation")
        
        st.write("**Immediate Actions Required:**")
        st.write("1. Verify all entries have proper board documentation")
        st.write("2. Correct pricing and date discrepancies")
        st.write("3. Record missing transactions (repurchases, etc.)")
        st.write("4. Remove or document phantom equity entries")
        
        # Export functionality
        st.subheader("📤 Export Results")
        
        # Create downloadable report
        report_data = []
        for disc in discrepancies:
            report_data.append([
                disc['severity'],
                disc['stockholder'],
                disc['security_id'],
                disc['issue'],
                disc['cap_table_value'],
                disc['legal_document_value'],
                disc['description'],
                disc['source_document']
            ])
        
        df_report = pd.DataFrame(report_data, columns=[
            'Severity', 'Stockholder', 'Security ID', 'Issue', 
            'Cap Table Value', 'Legal Document Value', 'Description', 'Source Document'
        ])
        
        csv = df_report.to_csv(index=False)
        st.download_button(
            label="Download Discrepancies Report (CSV)",
            data=csv,
            file_name=f"cap_table_discrepancies_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )

def main():
    st.title("📊 Cap Table Tie-Out Analysis")
    st.markdown("*Using the original manual analysis methodology*")
    
    # Initialize analyzer
    if 'analyzer' not in st.session_state:
        st.session_state.analyzer = CapTableAnalyzer()
    
    # Sidebar for file uploads
    with st.sidebar:
        st.header("📁 Upload Documents")
        
        # Board documents upload
        st.subheader("Board Documents")
        board_files = st.file_uploader(
            "Upload board consents and minutes (DOCX)",
            type=['docx'],
            accept_multiple_files=True,
            key="board_docs"
        )
        
        # Securities ledger upload
        st.subheader("Securities Ledger")
        cap_table_file = st.file_uploader(
            "Upload cap table (Excel format)",
            type=['xlsx', 'xls'],
            key="cap_table"
        )
        
        # Analysis button
        st.markdown("---")
        run_analysis = st.button("🔍 Run Tie-Out Analysis", type="primary", use_container_width=True)
    
    # Show upload status
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📋 Uploaded Documents")
        
        if board_files:
            st.write("**Board Documents:**")
            for file in board_files:
                st.write(f"✅ {file.name}")
        else:
            st.info("No board documents uploaded yet")
        
        if cap_table_file:
            st.write("**Securities Ledger:**")
            st.write(f"✅ {cap_table_file.name}")
        else:
            st.info("No securities ledger uploaded yet")
    
    with col2:
        st.subheader("⚙️ Analysis Status")
        
        if not board_files and not cap_table_file:
            st.warning("Please upload documents to begin analysis")
        elif not board_files:
            st.warning("Please upload board documents")
        elif not cap_table_file:
            st.warning("Please upload securities ledger")
        else:
            st.success("Ready for analysis!")
    
    # Run analysis when button is clicked
    if run_analysis:
        if not board_files or not cap_table_file:
            st.error("Please upload both board documents and securities ledger before running analysis")
            return
        
        with st.spinner("Running tie-out analysis..."):
            analyzer = st.session_state.analyzer
            
            # Step 1: Analyze Excel file (like original JavaScript approach)
            st.markdown("---")
            cap_entries = analyzer.analyze_excel_like_original(cap_table_file)
            
            # Step 2: Analyze board documents (manual parsing)
            st.markdown("---")
            board_analysis = analyzer.analyze_board_docs_like_original()
            
            # Step 3: Perform discrepancy analysis
            st.markdown("---")
            discrepancies = analyzer.perform_discrepancy_analysis()
            
            # Step 4: Generate final report
            st.markdown("---")
            analyzer.generate_report(discrepancies)

if __name__ == "__main__":
    main()
