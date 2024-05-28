import requests
from bs4 import BeautifulSoup
import logging
import pandas as pd
import os
import traceback

def scrape_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')

    email_data = []
    for row in soup.find_all('tr'):
        cells = row.find_all('td')
        if len(cells) == 15:  # Check if it's a valid row with 15 cells
            case_number = cells[1].text.strip()
            email_time = cells[8].text.strip()
            case_owner = cells[5].text.strip()
            manager = cells[10].text.strip()
            severity = cells[3].text.strip()
            
            if manager == "Maria olivia Lobo":
                email_data.append({
                    'CaseNumber': case_number,
                    'Email Time': email_time,
                    'Case Owner': case_owner,
                    'Manager': manager,
                    'Severity': severity
                })

    return email_data

def save_to_log(data):
    logging.debug("Starting save_to_log function.")
    for email in data:
        logging.info(f"Case Number: {email['CaseNumber']}, Time: {email['Email Time']}, Case Owner: {email['Case Owner']}, Manager: {email['Manager']}, Severity: {email['Severity']}")

def save_to_excel(data, excel_file):
    df = pd.DataFrame(data)
    
    if not os.path.exists(excel_file):
        df.to_excel(excel_file, index=False)
        logging.debug(f"Created new Excel file {excel_file}.")
    else:
        try:
            existing_df = pd.read_excel(excel_file)
            if not existing_df.equals(df):
                existing_df = df
                existing_df.to_excel(excel_file, index=False)
                logging.debug(f"Updated data in existing Excel file {excel_file}.")
            else:
                logging.debug("No new data found. Excel file remains unchanged.")
        except Exception as e:
            logging.error(f"Error occurred while updating Excel file {excel_file}: {e}")
            logging.error(traceback.format_exc())

def main():
    # Configure logging
    logging.basicConfig(filename='scraping_log.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

    try:
        # URL of amsall.html
        html_url = 'http://megds55hi0.asiapacific.hpqcorp.net/AMSEmail/AMSAll.html'

        # Download amsall.html
        response = requests.get(html_url)
        response.raise_for_status()  # Raise an HTTPError for bad responses

        # Extract data from amsall.html
        email_data = scrape_from_html(response.content)

        # Save data to log file
        save_to_log(email_data)

        # Save data to Excel file
        excel_file = 'scraped_emails.xlsx'
        save_to_excel(email_data, excel_file)
    except Exception as e:
        logging.error(f"Error occurred during scraping and saving process: {e}")
        logging.error(traceback.format_exc())

# Call the main function directly
main()
