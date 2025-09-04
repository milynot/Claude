import re
import pandas as pd
from datetime import datetime
import os
from bs4 import BeautifulSoup

class IBStatementExtractor:
    def __init__(self):
        self.data = []
    
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
    
    def detect_html_format(self, soup):
        """Detect the format of the HTML statement"""
        # Check for 2013 format indicators
        if soup.find('div', {'id': re.compile(r'tblFIFOPerfSumByUnderlying.*Body')}):
            return '2013'
        
        # Check for newer format indicators
        summary_rows = soup.find_all('tr')
        for row in summary_rows:
            cells = row.find_all('td')
            if len(cells) >= 1:
                cell_text = cells[0].get_text().strip()
                if 'U***' in cell_text:
                    return 'new'
                elif re.match(r'U\d+F?$', cell_text):
                    return 'old'
        
        return 'unknown'
    
    def extract_accounts_from_html(self, soup):
        """Extract accounts from HTML (works for all formats)"""
        accounts = []
        
        # Try account summary table first
        summary_rows = soup.find_all('tr')
        for row in summary_rows:
            cells = row.find_all('td')
            if len(cells) >= 6:
                cell_text = cells[0].get_text().strip()
                # Look for both masked (U***6153) and full (U1046153) account numbers
                if 'U***' in cell_text or re.match(r'U\d+F?$', cell_text):
                    account_num = cell_text
                    name = cells[2].get_text().strip()
                    accounts.append({
                        'account_number': account_num,
                        'name': name
                    })
        
        # If no accounts found, try account information sections (2013 format)
        if not accounts:
            account_info_sections = soup.find_all('div', {'id': re.compile(r'tblAccountInformation_.*Body')})
            for section in account_info_sections:
                table = section.find('table')
                if table:
                    rows = table.find_all('tr')
                    account_num = None
                    name = None
                    for row in rows:
                        cells = row.find_all('td')
                        if len(cells) >= 2:
                            if cells[0].get_text().strip() == 'Account':
                                account_num = cells[1].get_text().strip()
                            elif cells[0].get_text().strip() == 'Name':
                                name = cells[1].get_text().strip()
                    
                    if account_num and name:
                        accounts.append({
                            'account_number': account_num,
                            'name': name
                        })
        
        return accounts
    
    def extract_pnl_from_html_2013(self, soup, account_num):
        """Extract P&L data from 2013 HTML format"""
        pnl_data = {
            'stocks': 0,
            'options': 0,
            'forex': 0,
            'total': 0
        }
        
        section_id = f"tblFIFOPerfSumByUnderlying{account_num}Body"
        pnl_section = soup.find('div', {'id': section_id})
        
        if pnl_section:
            print(f"Found 2013 P&L section for {account_num}")
            table = pnl_section.find('table')
            
            if table:
                rows = table.find_all('tr')
                
                for row in rows:
                    cells = row.find_all('td')
                    if not cells:
                        continue
                        
                    first_cell_text = cells[0].get_text().strip()
                    
                    # Look for the "Total" row - realized P&L is in column 6
                    if first_cell_text == 'Total' and len(cells) > 6:
                        try:
                            cell_value = cells[6].get_text().replace(',', '').strip()
                            pnl_data['total'] = float(cell_value)
                            print(f"  2013: Extracted realized P&L: {pnl_data['total']}")
                            break
                        except (ValueError, IndexError) as e:
                            print(f"  Error parsing 2013 total: {e}")
        else:
            print(f"No 2013 P&L section found for {account_num}")
        
        return pnl_data
    
    def extract_pnl_from_html_2021(self, soup, account_num):
        """Extract P&L data from 2021+ HTML format"""
        pnl_data = {
            'stocks': 0,
            'options': 0,
            'forex': 0,
            'total': 0
        }
        
        # Look for the Realized & Unrealized Performance Summary table (2021+ format)
        section_id = f"tblFIFOPerfSumByUnderlying{account_num}Body"
        pnl_section = soup.find('div', {'id': section_id})
        
        if pnl_section:
            print(f"Found 2021+ P&L section for {account_num}")
            rows = pnl_section.find_all('tr')
            
            for row in rows:
                row_classes = row.get('class', [])
                if 'subtotal' in row_classes or 'total' in row_classes:
                    cells = row.find_all('td')
                    if len(cells) >= 7:
                        first_cell = cells[0].get_text().strip()
                        
                        # Column 6 (index 6) contains realized total for 2021+ formats
                        if 'Total Stocks' in first_cell:
                            try:
                                pnl_data['stocks'] = float(cells[6].get_text().replace(',', ''))
                                print(f"  2021+: Stocks realized: {pnl_data['stocks']}")
                            except Exception as e:
                                print(f"  Error parsing stocks: {e}")
                        
                        elif 'Total Equity and Index Options' in first_cell:
                            try:
                                pnl_data['options'] = float(cells[6].get_text().replace(',', ''))
                                print(f"  2021+: Options realized: {pnl_data['options']}")
                            except Exception as e:
                                print(f"  Error parsing options: {e}")
                        
                        elif 'Total Forex' in first_cell:
                            try:
                                pnl_data['forex'] = float(cells[6].get_text().replace(',', ''))
                                print(f"  2021+: Forex realized: {pnl_data['forex']}")
                            except Exception as e:
                                print(f"  Error parsing forex: {e}")
                        
                        elif 'Total (All Assets)' in first_cell:
                            try:
                                pnl_data['total'] = float(cells[6].get_text().replace(',', ''))
                                print(f"  2021+: Total realized: {pnl_data['total']}")
                            except Exception as e:
                                print(f"  Error parsing total: {e}")
        else:
            print(f"No 2021+ P&L section found for {account_num}")
        
        return pnl_data
    
    def extract_pnl_from_html(self, html_path):
        """Extract P&L data directly from HTML tables"""
        try:
            with open(html_path, 'r', encoding='utf-8') as file:
                content = file.read()
                soup = BeautifulSoup(content, 'html.parser')
            
            # Detect format
            format_type = self.detect_html_format(soup)
            print(f"Detected format: {format_type}")
            
            # Extract accounts
            accounts = self.extract_accounts_from_html(soup)
            print(f"Found HTML accounts: {accounts}")
            
            if not accounts:
                print("No accounts found in HTML")
                return []
            
            results = []
            
            # For each account, find its P&L data using the appropriate format
            for account in accounts:
                account_num = account['account_number']
                print(f"Looking for P&L data for {account_num}")
                
                # Use format-specific extraction
                if format_type == '2013':
                    pnl_data = self.extract_pnl_from_html_2013(soup, account_num)
                else:
                    # Use 2021+ format for 'new', 'old', and 'unknown' formats
                    pnl_data = self.extract_pnl_from_html_2021(soup, account_num)
                
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
                    date_obj = datetime.strptime(end_date, '%B %d, %Y')
                    return date_obj.year, date_obj.month, start_date, end_date
                except:
                    return None, None, start_date, end_date
        return None, None, None, None
    
    def process_statement(self, file_path):
        """Process a single statement file (HTML only)"""
        print(f"Processing: {file_path}")
        
        if file_path.lower().endswith('.html'):
            # Process HTML file
            text = self.extract_text_from_html(file_path)
            year, month, start_date, end_date = None, None, None, None
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
                    'Stocks_Realized': result['pnl_data']['stocks'],
                    'Options_Realized': result['pnl_data']['options'],
                    'Forex_Realized': result['pnl_data']['forex'],
                    'Total_Realized': result['pnl_data']['total']
                }
                
                self.data.append(row)
                print(f"  FINAL RESULT for {result['account']}: Realized P&L = {result['pnl_data']['total']}")
    
    def process_folder(self, folder_path):
        """Process all HTML files in a folder"""
        html_pattern = os.path.join(folder_path, "*.html")
        import glob
        files = glob.glob(html_pattern)
        
        if not files:
            print(f"No HTML files found in {folder_path}")
            return
        
        files.sort()
        
        for file_path in files:
            self.process_statement(file_path)
    
    def save_to_excel(self, output_path="IB_PnL_Summary.xlsx"):
        """Save extracted data to Excel"""
        if not self.data:
            print("No data to save")
            return
        
        df = pd.DataFrame(self.data)
        df = df.sort_values(['Year', 'Month', 'Account'])
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Raw_Data', index=False)
            
            # Create summary sheets
            summary_by_year = df.groupby(['Account', 'Year']).agg({
                'Stocks_Realized': 'sum',
                'Options_Realized': 'sum', 
                'Forex_Realized': 'sum',
                'Total_Realized': 'sum'
            }).reset_index()
            summary_by_year.to_excel(writer, sheet_name='Summary_by_Year', index=False)
            
            monthly_summary = df.pivot_table(
                index=['Year', 'Month'],
                columns='Account',
                values='Total_Realized',
                aggfunc='sum'
            ).reset_index()
            monthly_summary.to_excel(writer, sheet_name='Monthly_Summary', index=False)
        
        print(f"Data saved to {output_path}")
        print(f"Processed {len(df)} account-month combinations")

def test_universal_extraction():
    """Test the universal extraction with both 2013 and 2021 files"""
    print("=== TESTING UNIVERSAL EXTRACTION ===")
    
    extractor = IBStatementExtractor()
    test_files = [
        "ActivityStatement.201311.html",  # 2013 format
        "ActivityStatement.202111.html"   # 2021 format
    ]
    
    for test_file in test_files:
        print(f"\n--- Testing {test_file} ---")
        try:
            extractor.process_statement(test_file)
        except FileNotFoundError:
            print(f"File {test_file} not found, skipping...")
        except Exception as e:
            print(f"Error processing {test_file}: {e}")
    
    # Show all results
    if extractor.data:
        print(f"\n=== FINAL SUMMARY ===")
        print(f"Total records extracted: {len(extractor.data)}")
        print("\nBy file:")
        for row in extractor.data:
            print(f"  {row['File']} - {row['Account']}: Total=${row['Total_Realized']}, Options=${row['Options_Realized']}, Stocks=${row['Stocks_Realized']}")
        
        extractor.save_to_excel("Test_Universal_Extract.xlsx")
        print("\nData saved to Test_Universal_Extract.xlsx")
        
    else:
        print("FAILED: No data extracted from any file")

def main():
    """Main function to process statements"""
    extractor = IBStatementExtractor()
    
    folder_path = input("Enter the path to your statements folder: ").strip()
    if os.path.exists(folder_path):
        extractor.process_folder(folder_path)
        extractor.save_to_excel("IB_PnL_Summary.xlsx")
    else:
        print("Folder not found. Please check the path.")

if __name__ == "__main__":
    # Test with the 2013 file first
    test_universal_extraction()
    
    # Uncomment the line below to run the main function for processing folders
    # main()
