import os
import email
from email import policy
from bs4 import BeautifulSoup
import pandas as pd
import re
import glob

def parse_eml(file_path):
    with open(file_path, 'rb') as f:
        msg = email.message_from_binary_file(f, policy=policy.default)
    
    date_str = msg.get('Date', '')
    
    html_content = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'text/html':
                payload = part.get_payload(decode=True)
                if payload:
                    html_content = payload.decode(part.get_content_charset() or 'utf-8', errors='ignore')
                break
    else:
        if msg.get_content_type() == 'text/html':
            payload = msg.get_payload(decode=True)
            if payload:
                html_content = payload.decode(msg.get_content_charset() or 'utf-8', errors='ignore')

    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Extract Policy Number
    policy_number = ""
    # Try to find text containing "The following online transaction"
    for element in soup.find_all(text=re.compile(r'The following online transaction.*', re.IGNORECASE)):
        match = re.search(r'\((P\d+)\)', element)
        if match:
            policy_number = match.group(1)
            break

    # Helper to find tables and types
    buy_amount = ""
    buy_type = ""
    rsp_amount = ""
    rsp_type = ""

    # Find Buy section
    buy_header = soup.find(lambda tag: tag.name in ['b', 'strong'] and 'Buy' in tag.get_text(strip=True))
    if buy_header:
        # Find the next p tag for type
        p_tag = buy_header.find_next('p')
        if p_tag:
            buy_type = p_tag.get_text(strip=True)
            
        # Find the next table
        buy_table = buy_header.find_next('table')
        if buy_table:
            rows = buy_table.find_all('tr')
            if len(rows) > 1:
                headers = [th.get_text(strip=True) for th in rows[0].find_all(['th', 'td'])]
                cols = [td.get_text(strip=True) for td in rows[1].find_all('td')]
                
                # Find Investment Amount
                for i, h in enumerate(headers):
                    if 'Investment Amount' in h and i < len(cols):
                        buy_amount = cols[i]
                        break

    # Find RSP section
    rsp_header = soup.find(lambda tag: tag.name in ['b', 'strong'] and 'RSP Application' in tag.get_text(strip=True))
    if rsp_header:
        p_tag = rsp_header.find_next('p')
        if p_tag:
            rsp_type = p_tag.get_text(strip=True)
            
        rsp_table = rsp_header.find_next('table')
        if rsp_table:
            rows = rsp_table.find_all('tr')
            if len(rows) > 1:
                headers = [th.get_text(strip=True) for th in rows[0].find_all(['th', 'td'])]
                cols = [td.get_text(strip=True) for td in rows[1].find_all('td')]
                
                for i, h in enumerate(headers):
                    if 'RSP Amount' in h and i < len(cols):
                        rsp_amount = cols[i]
                        break

    return {
        'File Name': os.path.basename(file_path),
        'Date': date_str,
        'Policy Number': policy_number,
        'Buy Wording': buy_type,
        'Buy Amount': buy_amount,
        'RSP Wording': rsp_type,
        'RSP Amount': rsp_amount
    }

def main():
    folder_path = 'attached_assets' # Defaulting to the folder with the uploaded file
    eml_files = glob.glob(os.path.join(folder_path, '*.eml'))
    
    if not eml_files:
        print(f"No .eml files found in '{folder_path}' directory.")
        return

    data = []
    for file in eml_files:
        try:
            row = parse_eml(file)
            data.append(row)
        except Exception as e:
            print(f"Error parsing {file}: {e}")
            
    if data:
        df = pd.DataFrame(data)
        output_file = 'extracted_data.xlsx'
        df.to_excel(output_file, index=False)
        print(f"Successfully extracted data from {len(data)} files to {output_file}")
    else:
        print("No data was extracted.")

if __name__ == "__main__":
    main()
