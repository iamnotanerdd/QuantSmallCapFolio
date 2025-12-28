import os
import requests
import datetime

# Configuration
# URL Pattern: https://quantmutual.com/Admin/disclouser/quant_Small_Cap_Fund_Jan_2025.xlsx
BASE_URL = "https://quantmutual.com/statutory-disclosures" # For reference
DL_BASE_URL = "https://quantmutual.com/Admin/disclouser/"
SCHEME_NAME_PREFIX = "quant_Small_Cap_Fund"
# DOWNLOAD_DIR = os.path.join(os.path.dirname(__file__), "2025_Disclosures") # OLD
DOWNLOAD_DIR = os.path.join(os.path.dirname(__file__), "Disclosures")

def get_monthly_urls(year):
    """
    Generates URLs for the target scheme for the given year.
    Returns a list of dictionaries with month, year, and url.
    """
    months = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", 
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
    ]
    
    current_date = datetime.date.today()
    results = []
    
    for i, month in enumerate(months):
        # Stop if we are looking for future months
        if year > current_date.year:
            break
        if year == current_date.year and (i + 1) > current_date.month:
            break
            
        # Construct URL
        # Pattern observed: quant_Small_Cap_Fund_Jan_2025.xlsx
        filename = f"{SCHEME_NAME_PREFIX}_{month}_{year}.xlsx"
        url = f"{DL_BASE_URL}{filename}"
        
        results.append({
            "month": month,
            "year": year,
            "url": url,
            "filename": filename
        })
        
    return results

def download_file(url, filepath):
    # Ensure directory exists
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    
    # Check if exists
    if os.path.exists(filepath):
        print(f"Skipping {os.path.basename(filepath)}, already exists.")
        return

    print(f"Downloading {os.path.basename(filepath)} from {url}...")
    try:
        # Use a session to look like a browser if needed, but simple get might work
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        r = requests.get(url, headers=headers, stream=True)
        
        # If 404, maybe try 'Sept' instead of 'Sep' or other variations if needed
        if r.status_code == 404 and "Sep" in url:
             print("404 for Sep, trying Sept...")
             url = url.replace("Sep", "Sept")
             r = requests.get(url, headers=headers, stream=True)

        r.raise_for_status()
        
        with open(filepath, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
        print(f"Saved to {filepath}")
        
    except requests.RequestException as e:
        print(f"Failed to download {os.path.basename(filepath)}: {e}")

def main():
    start_year = 2025
    current_year = datetime.date.today().year
    
    print(f"Starting downloader for {SCHEME_NAME_PREFIX}...")
    print(f"Years: {start_year} to {current_year}")
    print(f"Target Directory: {DOWNLOAD_DIR}")
    
    all_links = []
    for year in range(start_year, current_year + 1):
        print(f"Generating links for {year}...")
        links = get_monthly_urls(year)
        all_links.extend(links)
    
    if not all_links:
        print("No links generated.")
        return

    print(f"Found {len(all_links)} potential download links.")
    
    for item in all_links:
        filepath = os.path.join(DOWNLOAD_DIR, item['filename'])
        download_file(item['url'], filepath)

if __name__ == "__main__":
    main()
