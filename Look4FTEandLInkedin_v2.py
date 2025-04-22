import pandas as pd
from googlesearch import search
import time
import requests
import re

# File paths
input_file = r"\Users.xlsx"
output_file = r"\Users.xlsx"

# Load the Excel file into a DataFrame
df = pd.read_excel(input_file)

# Iterate over each row where the "FTE" is empty
for idx, row in df.iterrows():
    fte_value = row.get("FTE")
    
    if not (pd.isnull(fte_value) or str(fte_value).strip() == ""):
        continue  # Skip if FTE is already filled

    company_name = row.get("Company name")
    if pd.isnull(company_name) or str(company_name).strip() == "":
        continue  # Skip if no company name is available

    query = f"{company_name} LinkedIn employees"
    print(f"Searching for: {query}")
    
    success = False
    backoff = 10  # initial delay in seconds for retrying on errors
    found_employee_count = None
    found_linkedin_url = None

    # Perform search and extraction in a retry loop for resilience
    while not success:
        try:
            count = 0
            # Iterate up to 10 results
            for url in search(query, num_results=10):
                count += 1
                # Consider only LinkedIn company pages
                if "linkedin.com/company" in url.lower():
                    print(f"Trying URL: {url}")
                    headers = {
                        "User-Agent": (
                            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                            "AppleWebKit/537.36 (KHTML, like Gecko) "
                            "Chrome/115.0.0.0 Safari/537.36"
                        )
                    }
                    response = requests.get(url, headers=headers, timeout=10)
                    if response.status_code == 200:
                        # Only assign the LinkedIn URL if the row's 'Company linkedin' is empty.
                        current_linkedin = row.get("Company linkedin ")
                        if pd.isnull(current_linkedin) or str(current_linkedin).strip() == "":
                            found_linkedin_url = url

                        # Search for a pattern like "1,234 employees"
                        match = re.search(r'([\d,]+)\s+employees', response.text, re.IGNORECASE)
                        if match:
                            num_str = match.group(1).replace(",", "")
                            try:
                                found_employee_count = int(num_str)
                                print(f"Found for '{company_name}': {found_employee_count} employees")
                            except Exception as e:
                                print(f"Error converting '{num_str}' to int for '{company_name}': {e}")
                        else:
                            print(f"No employee count found in the URL for '{company_name}'")
                    else:
                        print(f"Failed to fetch URL for '{company_name}', status code: {response.status_code}")

                # Exit if we've checked 10 results or if employee count has been found
                if count >= 10 or found_employee_count is not None:
                    break
                
                # Pause between each request to reduce the chance of getting blocked
                time.sleep(5)
            
            success = True
        except requests.exceptions.HTTPError as http_err:
            print(f"HTTP error for '{company_name}': {http_err}. Waiting {backoff} seconds before retrying.")
            time.sleep(backoff)
            backoff *= 2  # exponential backoff
        except Exception as err:
            print(f"Error for '{company_name}': {err}. Waiting {backoff} seconds before retrying.")
            time.sleep(backoff)
            backoff *= 2

    # Update the DataFrame:
    if found_employee_count is not None:
        df.at[idx, "FTE"] = found_employee_count
    else:
        print(f"No employee count found for '{company_name}'")
    
    # Only update 'Company linkedin ' if it's currently empty and a link was found
    current_linkedin = row.get("Company linkedin ")
    if (pd.isnull(current_linkedin) or str(current_linkedin).strip() == "") and found_linkedin_url is not None:
        df.at[idx, "Company linkedin "] = found_linkedin_url
    else:
        if found_linkedin_url is None:
            print(f"No LinkedIn URL found for '{company_name}'")
    
    # Brief pause before processing the next company
    time.sleep(3)

# Save the updated DataFrame to the output Excel file
df.to_excel(output_file, index=False)
print(f"Updated file saved to: {output_file}")
