import os
import requests
import openpyxl
from bs4 import BeautifulSoup

# Open the Excel file and select the active worksheet
workbook = openpyxl.load_workbook('input.xlsx')
worksheet = workbook.active

# Loop through each row in the worksheet
for row in worksheet.iter_rows(min_row=2, values_only=True):
    url_id, url = row

    # Make a request to the URL
    response = requests.get(url)

    # Check if the response status code is 200 (OK)
    if response.status_code == 200:
        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(response.content, 'html.parser')

        # Extract the article title
        try:
            title = soup.find('h1', class_='entry-title').text.strip()
        except AttributeError:
            print(f"No title found for URL {url}")
            continue

        # Extract the article text
        text = '\n'
        for p in soup.select('div.td-post-content.tagdiv-type p'):
            text += p.text.strip() + '\n'

        # Define the filename
        filename = f"{url_id}.txt"

        # Save the text to a file
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(title)
            f.write(text)

        # Print a message to confirm that the file was saved
        print(f'The article text was saved to "{os.getcwd()}\\{filename}"')
    else:
        print(f"Error: Response status code {response.status_code} for URL {url}")
