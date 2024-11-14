import pandas as pd
import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
import openpyxl
import time

# Set your Google Maps API key here
GOOGLE_MAPS_API_KEY = 'YOUR_GOOGLE_MAPS_API_KEY'
WHEELING_ZIP = '60090'

def load_companies_from_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
        required_columns = ['Name', 'City', 'ST', 'ZIP']
        if not set(required_columns).issubset(df.columns):
            raise ValueError(f"Required columns {required_columns} not found in the Excel file.")
        print("Successfully loaded company data from Excel.")
        return df[required_columns]
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()

def get_google_search_url(company_name, city, state):
    query = f'"{company_name}" careers data analyst OR data engineer OR data manager {city} {state}'
    return f"https://www.google.com/search?q={requests.utils.quote(query)}"

def scrape_job_listings(company_name, city, state):
    search_url = get_google_search_url(company_name, city, state)
    response = requests.get(search_url, headers={"User-Agent": "Mozilla/5.0"})
    soup = BeautifulSoup(response.text, 'html.parser')
    
    results = []
    for g in soup.select('.tF2Cxc'):
        title = g.select_one('.DKV0Md').text if g.select_one('.DKV0Md') else "Not specified"
        link = g.select_one('.yuRUbf a')['href'] if g.select_one('.yuRUbf a') else "Not specified"
        description = g.select_one('.IsZvec').text if g.select_one('.IsZvec') else "Not specified"
        
        results.append({
            'Company Name': company_name,
            'Job Title': title,
            'Job URL': link,
            'Job Description': description,
            'Location': f"{city}, {state}",
            'Distance (miles)': calculate_distance(f"{city}, {state}")
        })
    return results

def calculate_distance(destination):
    """Calculate distance from Wheeling, IL using Google Maps API"""
    try:
        url = f"https://maps.googleapis.com/maps/api/distancematrix/json?units=imperial&origins={WHEELING_ZIP}&destinations={destination}&key={GOOGLE_MAPS_API_KEY}"
        response = requests.get(url)
        data = response.json()

        if data['status'] == 'OK':
            distance = data['rows'][0]['elements'][0]['distance']['text']
            return distance
        else:
            return "Distance not found"
    except Exception as e:
        print(f"Error calculating distance: {e}")
        return "Distance not found"

def save_to_excel(data, output_file):
    try:
        df = pd.DataFrame(data)
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Job Listings')
        print(f"Successfully saved results to {output_file}")
    except Exception as e:
        print(f"Error saving to Excel: {e}")

def process_company(row):
    company_name = row['Name'].strip()
    city = row['City'].strip()
    state = row['ST'].strip()
    return scrape_job_listings(company_name, city, state)

def main():
    input_file = "C:/Users/jzacaria/OneDrive - True Value Company/Desktop/Python Scripts/VendorPR_IL.xlsx"
    output_file = "C:/Users/jzacaria/OneDrive - True Value Company/Desktop/Python Scripts/JobListings.xlsx"
    
    # Step 1: Load companies from Excel
    companies_df = load_companies_from_excel(input_file)
    
    if companies_df.empty:
        print("No valid company data found. Exiting.")
        return

    job_results = []
    terminal_display_count = 0  # To limit terminal output
    
    # Step 2: Use ThreadPoolExecutor for multi-threading
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(process_company, row): row for idx, row in companies_df.iterrows()}

        for future in as_completed(futures):
            listings = future.result()
            job_results.extend(listings)
            
            # Print only the first 5 results to the terminal
            if terminal_display_count < 5 and listings:
                for listing in listings[:5]:
                    print(f"Title: {listing['Job Title']}\nLink: {listing['Job URL']}\nLocation: {listing['Location']}\nDistance: {listing['Distance (miles)']}\n----")
                terminal_display_count += 1

    # Step 3: Save results to Excel
    if job_results:
        save_to_excel(job_results, output_file)
    else:
        print("No job listings found.")

if __name__ == "__main__":
    start_time = time.time()
    main()
    print(f"Script completed in {time.time() - start_time:.2f} seconds")
