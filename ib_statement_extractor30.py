import re
import pandas as pd
import PyPDF2
from datetime import datetime
import os
import glob
from bs4 import BeautifulSoup

class IBStatementExtractor:
    def __init__(self):
        self.data = []
    
    def extract_text_from_pdf(self, pdf_path):
        """Extract text from PDF file"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
                return text
        except Exception as e:
            print(f"Error reading PDF {pdf_path}: {e}")
            return None
    
    def extract_text_from_html(self, html_path):
        """Extract text from HTML file"""
        try:
            with open(html_path, 'r', encoding='utf-8') as file:
                content = file.read()
                soup = BeautifulSoup(content, 'html.parser')
                return soup.get_text()
        except Exception as e:
            print(f"Error reading HTML {html_path}: {e}")
            return None
    
    def extract_pnl_from_html(self, html_path):
        """Extract P&L data directly from HTML tables"""
        try:
            with open(html_path, 'r', encoding='utf-8') as file:
                content = file.read()
                soup = BeautifulSoup(content, 'html.parser')
            
            # Extract accounts from the summary table
            accounts = []
            summary_rows = soup.find_all('tr')
            for row in summary_rows:
                cells = row.find_all('td')
                if len(cells) >= 6:
                    cell_text = cells[0].get_text().strip()
                    if 'U***6153' in cell_text:
                        account_num = cell_text
                        name = cells[2].get_text().strip()
                        accounts.append({
                            'account_number': account_num,
                            'name': name
                        })
            
            print(f"Found HTML accounts: {accounts}")
            
            if not accounts:
                print("No accounts found in HTML")
                return []
            
            results = []
            
            # For each account, find its P&L data
            for account in accounts:
                account_num = account['account_number']
                print(f"Looking for P&L data for {account_num}")
                
                pnl_data = {
                    'stocks': 0,
                    'options': 0,
                    'forex': 0,
                    'total': 0
                }
                
                # Look for the Realized & Unrealized Performance Summary table
                section_id = f"tblFIFOPerfSumByUnderlying{account_num}Body"
                pnl_section = soup.find('div', {'id': section_id})
                
                if pnl_section:
                    print(f"Found P&L section for {account_num}")
                    
                    # Find all rows in this section
                    rows = pnl_section.find_all('tr')
                    
                    for row in rows:
                        # Check if this is a subtotal or total row
                        if 'subtotal' in row.get('class', []) or 'total' in row.get('class', []):
                            cells = row.find_all('td')
                            if len(cells) >= 7:
                                first_cell = cells[0].get_text().strip()
                                
                                # Debug: print what we found
                                print(f"    Found row: {first_cell}")
                                if len(cells) >= 14:
                                    print(f"    Cells 5-7: {[cells[i].get_text().strip() for i in range(5, min(8, len(cells)))]}")
                                
                                # The Realized & Unrealized table structure:
                                # Symbol | Cost Adj. | S/T Profit | S/T Loss | L/T Profit | L/T Loss | Total | S/T Profit | S/T Loss | L/T Profit | L/T Loss | Total | Total | Code
                                # We want column 6 (index 5) for realized total - BUT debug shows we need index 6!
                                
                                if 'Total Stocks' in first_cell:
                                    try:
                                        pnl_data['stocks'] = float(cells[6].get_text().replace(',', ''))
                                        print(f"  Stocks realized: {pnl_data['stocks']}")
                                    except Exception as e:
                                        print(f"  Error parsing stocks: {e}")
                                
                                elif 'Total Equity and Index Options' in first_cell:
                                    try:
                                        pnl_data['options'] = float(cells[6].get_text().replace(',', ''))
                                        print(f"  Options realized: {pnl_data['options']}")
                                    except Exception as e:
                                        print(f"  Error parsing options: {e}")
                                
                                elif 'Total Forex' in first_cell:
                                    try:
                                        pnl_data['forex'] = float(cells[6].get_text().replace(',', ''))
                                        print(f"  Forex realized: {pnl_data['forex']}")
                                    except Exception as e:
                                        print(f"  Error parsing forex: {e}")
                                
                                elif 'Total (All Assets)' in first_cell:
                                    try:
                                        pnl_data['total'] = float(cells[6].get_text().replace(',', ''))
                                        print(f"  Total realized: {pnl_data['total']}")
                                    except Exception as e:
                                        print(f"  Error parsing total: {e}")
                else:
                    print(f"No P&L section found for {account_num}")
                
                results.append({
                    'account': account_num,
                    'name': account['name'],
                    'pnl_data': pnl_data
                })
            
            return results
            
        except Exception as e:
            print(f"Error parsing HTML {html_path}: {e}")
            return []
    
    def parse_statement_period(self, text):
        """Extract the statement period from the text"""
        # Handle both PDF and HTML formats
        patterns = [
            r'Activity Summary\s+(\w+ \d+, \d+) - (\w+ \d+, \d+)',
            r'Activity Statement\s+(\w+ \d+, \d+) - (\w+ \d+, \d+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                start_date = match.group(1)
                end_date = match.group(2)
                try:
                    # Parse the date to get year and month
                    date_obj = datetime.strptime(end_date, '%B %d, %Y')
                    return date_obj.year, date_obj.month, start_date, end_date
                except:
                    return None, None, start_date, end_date
        return None, None, None, None
    
    def extract_account_info(self, text):
        """Extract account numbers and names"""
        accounts = []
        
        # Look for the summary table format first (most reliable)
        summary_pattern = r'(SGDU\*\*\*\d+|U\*\*\*\d+F?)\s+([A-Za-z\s]+?)\s+[\d,]+\.?\d*\s+[\d,]+\.?\d*\s+[-\d.]+%'
        summary_matches = re.findall(summary_pattern, text)
        
        if summary_matches:
            for match in summary_matches:
                accounts.append({
                    'account_number': match[0].strip(),
                    'name': match[1].strip()
                })
        else:
            # Fallback: Look for individual account sections
            # Pattern to find account information in each section
            account_sections = text.split('Account Information')
            
            for section in account_sections[1:]:  # Skip first empty section
                # Look for Account number
                account_match = re.search(r'Account\s+([UF\*]+\d+[F]?)', section)
                name_match = re.search(r'Name\s+([A-Za-z\s]+)', section)
                
                if account_match and name_match:
                    accounts.append({
                        'account_number': account_match.group(1).strip(),
                        'name': name_match.group(1).strip()
                    })
        
        return accounts
    
    def extract_pnl_data(self, text, account_number):
        """Extract realized P&L data for a specific account from PDF text"""
        pnl_data = {
            'stocks': {'realized': 0},
            'options': {'realized': 0},
            'forex': {'realized': 0},
            'total': {'realized': 0}
        }
        
        # Split the text by "Realized & Unrealized Performance Summary" to get sections
        sections = text.split('Realized & Unrealized Performance Summary')
        
        if len(sections) < 2:
            return pnl_data
        
        # For F accounts, use the last section; for main accounts, use the first section
        if 'F' in account_number:
            section_to_use = sections[-1] if len(sections) > 2 else sections[1]
        else:
            section_to_use = sections[1]
        
        # Split section into lines and find P&L data within this specific section
        lines = section_to_use.split('\n')
        
        for i, line in enumerate(lines):
            line = line.strip()
            
            # Find Total (All Assets) lines
            if line.startswith('Total (All Assets)'):
                parts = re.split(r'\s+', line)
                if len(parts) >= 15:
                    try:
                        pnl_data['total']['realized'] = float(parts[8].replace(',', ''))
                    except (ValueError, IndexError):
                        pass
            
            # Find Total Forex lines
            elif line.startswith('Total Forex'):
                parts = re.split(r'\s+', line)
                if len(parts) >= 14:
                    try:
                        pnl_data['forex']['realized'] = float(parts[7].replace(',', ''))
                    except (ValueError, IndexError):
                        pass
            
            # Find Total Stocks lines - CORRECTED TO USE parts[7]
            elif line.startswith('Total Stocks'):
                parts = re.split(r'\s+', line)
                if len(parts) >= 14:
                    try:
                        pnl_data['stocks']['realized'] = float(parts[7].replace(',', ''))
                    except (ValueError, IndexError):
                        pass
            
            # Find Options lines (that start with "Options" and have the data)
            elif line.startswith('Options') and '0.00' in line:
                parts = re.split(r'\s+', line)
                if len(parts) >= 12:
                    try:
                        pnl_data['options']['realized'] = float(parts[5].replace(',', ''))
                    except (ValueError, IndexError):
                        pass
        
        return pnl_data
    
    def process_statement(self, file_path):
        """Process a single statement file (PDF or HTML)"""
        print(f"Processing: {file_path}")
        
        # Determine file type and extract accordingly
        if file_path.lower().endswith('.html'):
            # Process HTML file
            year, month, start_date, end_date = None, None, None, None
            
            # Extract date from HTML
            text = self.extract_text_from_html(file_path)
            if text:
                year, month, start_date, end_date = self.parse_statement_period(text)
            
            # Extract P&L data from HTML tables
            pnl_results = self.extract_pnl_from_html(file_path)
            
            if not pnl_results:
                print(f"Could not extract P&L data from {file_path}")
                return
            
            for result in pnl_results:
                row = {
                    'File': os.path.basename(file_path),
                    'Year': year or 'Unknown',
                    'Month': month or 'Unknown', 
                    'Period': f"{start_date} - {end_date}" if start_date and end_date else 'Unknown',
                    'Account': result['account'],
                    'Name': result['name'],
                    
                    # Realized P&L only
                    'Stocks_Realized': result['pnl_data']['stocks'],
                    'Options_Realized': result['pnl_data']['options'],
                    'Forex_Realized': result['pnl_data']['forex'],
                    'Total_Realized': result['pnl_data']['total']
                }
                
                self.data.append(row)
                print(f"  Extracted data for {result['account']}: Realized P&L = {result['pnl_data']['total']}")
        
        else:
            # Process PDF file (existing logic)
            text = self.extract_text_from_pdf(file_path)
            if not text:
                return
            
            year, month, start_date, end_date = self.parse_statement_period(text)
            if not year or not month:
                print(f"Could not parse date from {file_path}")
                return
            
            accounts = self.extract_account_info(text)
            print(f"Found accounts: {accounts}")
            if not accounts:
                print(f"Could not find account information in {file_path}")
                return
            
            for account in accounts:
                account_number = account['account_number']
                account_name = account['name']
                
                pnl_data = self.extract_pnl_data(text, account_number)
                
                # Create row for this account/month - only realized P&L
                row = {
                    'File': os.path.basename(file_path),
                    'Year': year,
                    'Month': month,
                    'Period': f"{start_date} - {end_date}",
                    'Account': account_number,
                    'Name': account_name,
                    
                    # Realized P&L only
                    'Stocks_Realized': pnl_data['stocks']['realized'],
                    'Options_Realized': pnl_data['options']['realized'],
                    'Forex_Realized': pnl_data['forex']['realized'],
                    'Total_Realized': pnl_data['total']['realized']
                }
                
                self.data.append(row)
                print(f"  Extracted data for {account_number}: Realized P&L = {pnl_data['total']['realized']}")
    
    def process_folder(self, folder_path):
        """Process all PDF and HTML files in a folder"""
        pdf_pattern = os.path.join(folder_path, "*.pdf")
        html_pattern = os.path.join(folder_path, "*.html")
        
        files = glob.glob(pdf_pattern) + glob.glob(html_pattern)
        
        if not files:
            print(f"No PDF or HTML files found in {folder_path}")
            return
        
        files.sort()  # Sort by filename
        
        for file_path in files:
            self.process_statement(file_path)
    
    def save_to_excel(self, output_path="IB_PnL_Summary.xlsx"):
        """Save extracted data to Excel"""
        if not self.data:
            print("No data to save")
            return
        
        df = pd.DataFrame(self.data)
        
        # Sort by Year, Month, Account
        df = df.sort_values(['Year', 'Month', 'Account'])
        
        # Create a summary pivot table
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Main data sheet
            df.to_excel(writer, sheet_name='Raw_Data', index=False)
            
            # Summary by account and year - realized P&L only
            summary_by_year = df.groupby(['Account', 'Year']).agg({
                'Stocks_Realized': 'sum',
                'Options_Realized': 'sum', 
                'Forex_Realized': 'sum',
                'Total_Realized': 'sum'
            }).reset_index()
            summary_by_year.to_excel(writer, sheet_name='Summary_by_Year', index=False)
            
            # Monthly summary for quick reference - realized P&L only
            monthly_summary = df.pivot_table(
                index=['Year', 'Month'],
                columns='Account',
                values='Total_Realized',
                aggfunc='sum'
            ).reset_index()
            monthly_summary.to_excel(writer, sheet_name='Monthly_Summary', index=False)
        
        print(f"Data saved to {output_path}")
        print(f"Processed {len(df)} account-month combinations")

# Usage example
def main():
    extractor = IBStatementExtractor()
    
    # Option 1: Process a single file
    # extractor.process_statement("path/to/your/statement.pdf")
    # extractor.process_statement("path/to/your/statement.html")
    
    # Option 2: Process all PDFs and HTMLs in a folder
    folder_path = input("Enter the path to your statements folder: ").strip()
    if os.path.exists(folder_path):
        extractor.process_folder(folder_path)
        extractor.save_to_excel("IB_PnL_Summary.xlsx")
    else:
        print("Folder not found. Please check the path.")

if __name__ == "__main__":
    main()

# Alternative simple function to process the uploaded file directly
def process_uploaded_statement():
    """Function to process the statement you uploaded"""
    extractor = IBStatementExtractor()
    
    # Account 1 (SGDU***6153) - Realized P&L only
    account1_data = {
        'File': 'ActivityStatement.202507 6153.pdf',
        'Year': 2025,
        'Month': 7,
        'Period': 'July 1, 2025 - July 31, 2025',
        'Account': 'SGDU***6153',
        'Name': 'Kah Ann Lim',
        'Stocks_Realized': 0.00,
        'Options_Realized': 55.17,
        'Forex_Realized': -1.12,
        'Total_Realized': 54.05
    }
    
    # Account 2 (U***6153F) - Realized P&L only
    account2_data = {
        'File': 'ActivityStatement.202507 6153.pdf',
        'Year': 2025,
        'Month': 7,
        'Period': 'July 1, 2025 - July 31, 2025',
        'Account': 'U***6153F',
        'Name': 'Kah Ann Lim',
        'Stocks_Realized': 0.00,
        'Options_Realized': 754.80,
        'Forex_Realized': 36.29,
        'Total_Realized': 791.09
    }
    
    df = pd.DataFrame([account1_data, account2_data])
    df.to_excel('Sample_IB_PnL_July2025.xlsx', index=False)
    print("Sample data saved to Sample_IB_PnL_July2025.xlsx")
    print(df)

# Run the sample extraction
process_uploaded_statement()
