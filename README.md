# Email-and-Phone-Scraper-EnP-
It extracts email addresses and any country mobile phone numbers from websites, normalizes and deduplicates them, and then displays the results in an intuitive GUI

**EnP**
EnP is an advanced web crawler application built with Python and PyQt5. It extracts email addresses and Italian mobile phone numbers from websites, normalizes and deduplicates them, and then displays the results in an intuitive GUI. With clickable WhatsApp links for each mobile number and bulk email functionality, EnP makes it easy to contact your leads—all in one app.

**Features**
Email & Mobile Extraction:
Extracts email addresses and Italian mobile phone numbers from websites.

Normalization & Deduplication:
Converts various mobile number formats (e.g., with spaces, dashes, or differing country-code representations) into a consistent format (always starting with +39) to avoid duplicates.

Clickable WhatsApp Links:
Each detected mobile number is displayed as an individual clickable link that opens WhatsApp (Web or Desktop) using the format https://wa.me/393311234567.

Manual Website Input & Excel Upload:
Add websites manually using an input field or load a list of websites from an Excel file.

Bulk Email Functionality:
Easily send bulk emails using your default email client. Simply click the "Send Bulk Email" button to open a pre-populated mailto link with all extracted email addresses.

Export Results:
Save your results as an Excel file for further analysis or record keeping.

User-Friendly GUI:
Built with PyQt5, EnP offers a clean and responsive interface with real-time logging and an elapsed time counter.

**Requirements**
Python 3.x
PyQt5
Pandas
Requests
BeautifulSoup4
(Optional) PyInstaller – if you wish to package the app as a standalone executable.

**Install Dependencies**
You can install the required packages using pip:

pip install PyQt5 pandas requests beautifulsoup4

**How to Use EnP**
1. Running the Application
Launch EnP:
Run the main Python script (e.g., EnP.py):
python EnP.py
This will open the EnP GUI.

2. Adding Websites
Manual Entry:
Enter a website URL (e.g., example.com) in the "Add Website Manually" field and click the Add Website button. The app will automatically prepend https:// if needed.

Excel Upload:
Click the Load Excel button to select an Excel file. EnP will automatically detect the column containing website URLs and add the websites to the list (avoiding duplicates).

3. Starting the Crawl
Begin Crawling:
Once websites are loaded or manually added, click the Start Crawling button. The app will start extracting emails and mobile numbers from each website. The log area will show progress messages and an elapsed time counter.

4. Viewing the Results
Results Table:
The main table displays the following for each website:
Website: URL of the crawled site.
Email Count & Emails: The number of emails found and the list of email addresses.
Mobile Numbers: The normalized mobile numbers.
WhatsApp Links: Clickable WhatsApp links for each mobile number. Clicking on a link will open WhatsApp Web/desktop for messaging.

5. Sending Bulk Email
Bulk Email:
Click the Send Bulk Email button. This will open your default email client with a mailto: link pre-populated with all extracted email addresses (comma-separated).

6. Exporting Results
Export to Excel:
Click the Export Results button to save the extracted data (websites, emails, mobile numbers, and WhatsApp links) to an Excel file.

7. Stopping the Crawl
Stop Crawling:
If you wish to halt the process, click the Stop Crawling button. The app will send a stop signal and finish processing the current website.
