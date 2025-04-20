import requests
from bs4 import BeautifulSoup

def get_bse500_pe_ratio():
    """
    Extracts only the PE ratio for BSE 500 from screener.in
    
    Returns:
        str: The PE ratio value as a string
    """
    url = "https://www.screener.in/company/1005/"
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Find the ratios list
        ratios_list = soup.find('ul', id='top-ratios')
        if not ratios_list:
            return None
        
        # Look for the P/E item in the list
        for item in ratios_list.find_all('li'):
            name_span = item.find('span', class_='name')
            value_span = item.find('span', class_='value')
            
            if name_span and value_span and 'P/E' in name_span.text.strip():
                # Extract the numeric part from the value
                value_number = value_span.find('span', class_='number')
                if value_number:
                    return value_number.text.strip()
        
        return None
        
    except Exception:
        return None