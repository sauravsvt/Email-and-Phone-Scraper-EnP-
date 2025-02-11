import sys
import time
import re
import pandas as pd
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QUrl
from PyQt5.QtGui import QDesktopServices

# Ensure you have installed phonenumbers: pip install phonenumbers
import phonenumbers
from phonenumbers import PhoneNumberMatcher, PhoneNumberFormat, number_type, PhoneNumberType

# ------------------------------------------------------------------------------
# Worker Thread for Crawling
# ------------------------------------------------------------------------------
class CrawlerWorker(QtCore.QThread):
    # Signals to communicate with the main GUI thread.
    log_signal = QtCore.pyqtSignal(str)
    website_done_signal = QtCore.pyqtSignal(str, set, set)  # website, emails set, mobiles set
    finished_signal = QtCore.pyqtSignal()

    def __init__(self, websites, max_pages=100, max_depth=0, dynamic_crawl=False,
                 region="Auto", email_limit=None, phone_limit=None, parent=None):
        """
        :param websites: List of starting website URLs.
        :param max_pages: Maximum number of pages to visit per website (0 for unlimited).
        :param max_depth: Maximum depth for crawling links (0 for unlimited).
        :param dynamic_crawl: If True, use dynamic crawling (Playwright) for all pages.
        :param region: Default region code (or "Auto") for phone number extraction.
        :param email_limit: Maximum number of emails to collect (None means unlimited).
        :param phone_limit: Maximum number of phone numbers to collect (None means unlimited).
        """
        super().__init__(parent)
        self.websites = websites
        self.max_pages = max_pages
        self.max_depth = max_depth  # 0 means no depth limit
        self.dynamic_crawl = dynamic_crawl
        self.default_region = region  # e.g. "Auto", "IT", "US", etc.
        self.email_limit = email_limit
        self.phone_limit = phone_limit
        self.stop_requested = False
        self.email_regex = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b')
        self.headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/115.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        }

    def normalize_url(self, url):
        """Remove the fragment part and lowercase the domain."""
        parsed = urlparse(url)
        netloc = parsed.netloc.lower()
        normalized = parsed._replace(netloc=netloc, fragment="")
        return normalized.geturl()

    def detect_region_from_url(self, url):
        """
        Try to detect a region code based on the website's TLD.
        If the TLD is not in our mapping (or ambiguous), fallback to default_region or "US".
        """
        parsed = urlparse(url)
        netloc = parsed.netloc
        parts = netloc.split('.')
        if len(parts) < 2:
            return self.default_region if self.default_region != "Auto" else "US"
        tld = parts[-1].upper()
        mapping = {
            "IT": "IT",
            "UK": "GB",  # .uk sites usually use region code "GB"
            "US": "US",
            "FR": "FR",
            "DE": "DE",
            "IN": "IN",
            "ES": "ES",
            "CN": "CN"
        }
        return mapping.get(tld, self.default_region if self.default_region != "Auto" else "US")

    def extract_phone_numbers(self, text, region, website_url):
        """
        Extract phone numbers using the phonenumbers library.
        Only numbers that are mobile (or fixed_line_or_mobile) are kept.
        :param text: Text to search.
        :param region: The default region to use (if not "Auto").
        :param website_url: URL of the website (for auto-detection if region == "Auto").
        :return: A set of phone numbers formatted in E.164.
        """
        numbers_found = set()
        # If region is "Auto", try to detect based on website TLD.
        if region == "Auto":
            detected_region = self.detect_region_from_url(website_url)
        else:
            detected_region = region

        try:
            for match in PhoneNumberMatcher(text, detected_region):
                phone_num = match.number
                num_type = number_type(phone_num)
                if num_type in [PhoneNumberType.MOBILE, PhoneNumberType.FIXED_LINE_OR_MOBILE]:
                    formatted = phonenumbers.format_number(phone_num, PhoneNumberFormat.E164)
                    numbers_found.add(formatted)
        except Exception as e:
            self.log_signal.emit(f"Error extracting phone numbers: {e}")
        return numbers_found

    def run(self):
        for website in self.websites:
            if self.stop_requested:
                self.log_signal.emit("Crawling stopped by user.")
                break

            website = website.strip()
            if not website.startswith("http"):
                website = "https://" + website

            self.log_signal.emit(f"\nStarting crawl for: {website}")
            emails, mobiles = self.crawl_website(website)
            self.website_done_signal.emit(website, emails, mobiles)
            email_list = ", ".join(emails) if emails else "None"
            mobile_list = ", ".join(mobiles) if mobiles else "None"
            self.log_signal.emit(
                f"Completed {website}: Found {len(emails)} email(s) and {len(mobiles)} mobile number(s). "
                f"Emails: {email_list} | Mobiles: {mobile_list}"
            )
        self.finished_signal.emit()

    def crawl_website(self, base_url):
        visited = set()
        emails_found = set()
        mobiles_found = set()
        queue = [(base_url, 0)]
        base_domain = urlparse(self.normalize_url(base_url)).netloc

        # If dynamic crawling is forced, launch Playwright.
        if self.dynamic_crawl:
            from playwright.sync_api import sync_playwright
            playwright = sync_playwright().start()
            browser = playwright.chromium.launch(headless=True)

        while queue:
            if self.stop_requested:
                break

            current_url, current_depth = queue.pop(0)
            normalized_current = self.normalize_url(current_url)
            if normalized_current in visited:
                continue

            if self.max_pages and len(visited) >= self.max_pages:
                self.log_signal.emit("Reached maximum page limit.")
                break

            visited.add(normalized_current)
            self.log_signal.emit(f"Visiting: {normalized_current}")

            try:
                if self.dynamic_crawl:
                    page = browser.new_page()
                    page.goto(normalized_current, wait_until="networkidle")
                    content = page.content()
                    page.close()
                else:
                    response = requests.get(normalized_current, timeout=10, headers=self.headers)
                    if response.status_code != 200:
                        self.log_signal.emit(
                            f"Failed to retrieve {normalized_current} (Status Code: {response.status_code})"
                        )
                        continue
                    content = response.text

                # Extract emails.
                found_emails = self.email_regex.findall(content)
                emails_found.update(found_emails)

                # Extract mobile numbers.
                found_mobiles = self.extract_phone_numbers(content, self.default_region, normalized_current)
                mobiles_found.update(found_mobiles)

                # Follow internal links.
                soup = BeautifulSoup(content, "html.parser")
                for a in soup.find_all('a', href=True):
                    href = a['href'].strip()
                    if href.startswith("#"):
                        continue
                    absolute_url = urljoin(normalized_current, href)
                    normalized_link = self.normalize_url(absolute_url)
                    link_domain = urlparse(normalized_link).netloc
                    if link_domain == base_domain and normalized_link not in visited:
                        if self.max_depth == 0 or current_depth < self.max_depth:
                            queue.append((normalized_link, current_depth + 1))
                time.sleep(1)  # Polite delay

                # Check if we have reached the desired limits.
                if (self.email_limit is not None or self.phone_limit is not None):
                    email_done = (self.email_limit is None or len(emails_found) >= self.email_limit)
                    phone_done = (self.phone_limit is None or len(mobiles_found) >= self.phone_limit)
                    if email_done and phone_done:
                        self.log_signal.emit("Threshold reached for this website. Stopping crawl for this website.")
                        break

            except Exception as e:
                self.log_signal.emit(f"Error accessing {normalized_current}: {e}")

        # Dynamic fallback: if static crawling yielded insufficient data.
        if not self.dynamic_crawl and (len(emails_found) == 0 or len(mobiles_found) == 0):
            self.log_signal.emit("Static crawling yielded insufficient data; trying dynamic fallback on base URL.")
            try:
                from playwright.sync_api import sync_playwright
                playwright = sync_playwright().start()
                browser_dynamic = playwright.chromium.launch(headless=True)
                page = browser_dynamic.new_page()
                page.goto(base_url, wait_until="networkidle")
                dynamic_content = page.content()
                page.close()
                browser_dynamic.close()
                playwright.stop()

                dynamic_emails = set(self.email_regex.findall(dynamic_content))
                dynamic_mobiles = set(self.extract_phone_numbers(dynamic_content, self.default_region, base_url))
                if dynamic_emails:
                    emails_found.update(dynamic_emails)
                if dynamic_mobiles:
                    mobiles_found.update(dynamic_mobiles)
                self.log_signal.emit(
                    f"Dynamic fallback added {len(dynamic_emails)} email(s) and {len(dynamic_mobiles)} mobile number(s)."
                )
            except Exception as e:
                self.log_signal.emit(f"Dynamic fallback failed: {e}")

        if self.dynamic_crawl:
            browser.close()
            playwright.stop()

        return emails_found, mobiles_found

    def stop(self):
        self.stop_requested = True

# ------------------------------------------------------------------------------
# Main GUI Application
# ------------------------------------------------------------------------------
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("EnP")
        self.setWindowIcon(QtGui.QIcon("EnP.ico"))
        self.resize(1100, 800)
        self.websites = []  # List to hold website URLs.
        self.results = {}   # Dictionary: {website: (emails set, mobiles set)}
        self.worker = None
        self.start_time = None

        self.initUI()

    def initUI(self):
        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QtWidgets.QVBoxLayout(central_widget)

        # --- Top Controls ---
        control_layout = QtWidgets.QHBoxLayout()

        self.load_button = QtWidgets.QPushButton("Load Excel")
        self.load_button.clicked.connect(self.load_excel)
        control_layout.addWidget(self.load_button)

        self.start_button = QtWidgets.QPushButton("Start Crawling")
        self.start_button.setEnabled(False)
        self.start_button.clicked.connect(self.start_crawling)
        control_layout.addWidget(self.start_button)

        self.stop_button = QtWidgets.QPushButton("Stop Crawling")
        self.stop_button.setEnabled(False)
        self.stop_button.clicked.connect(self.stop_crawling)
        control_layout.addWidget(self.stop_button)

        self.export_button = QtWidgets.QPushButton("Export Results")
        self.export_button.setEnabled(False)
        self.export_button.clicked.connect(self.export_results)
        control_layout.addWidget(self.export_button)

        self.bulk_email_button = QtWidgets.QPushButton("Send Bulk Email")
        self.bulk_email_button.setEnabled(False)
        self.bulk_email_button.clicked.connect(self.send_bulk_email)
        control_layout.addWidget(self.bulk_email_button)

        # Checkbox to force dynamic crawling.
        self.dynamic_checkbox = QtWidgets.QCheckBox("Force Dynamic Crawling (JS)")
        control_layout.addWidget(self.dynamic_checkbox)

        # Region selection for phone numbers.
        self.region_label = QtWidgets.QLabel("Phone Region:")
        control_layout.addWidget(self.region_label)
        self.region_combo = QtWidgets.QComboBox()
        self.region_combo.addItems(["Auto", "IT", "US", "GB", "FR", "DE", "IN"])
        self.region_combo.setToolTip("Select default region for phone number extraction. 'Auto' will try to detect based on website TLD.")
        control_layout.addWidget(self.region_combo)

        # Email and Phone limits.
        self.email_limit_label = QtWidgets.QLabel("Max Emails:")
        control_layout.addWidget(self.email_limit_label)
        self.email_limit_edit = QtWidgets.QLineEdit()
        self.email_limit_edit.setPlaceholderText("All")
        self.email_limit_edit.setFixedWidth(50)
        control_layout.addWidget(self.email_limit_edit)

        self.phone_limit_label = QtWidgets.QLabel("Max Phones:")
        control_layout.addWidget(self.phone_limit_label)
        self.phone_limit_edit = QtWidgets.QLineEdit()
        self.phone_limit_edit.setPlaceholderText("All")
        self.phone_limit_edit.setFixedWidth(50)
        control_layout.addWidget(self.phone_limit_edit)

        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setVisible(False)
        control_layout.addWidget(self.progress_bar)

        self.timer_label = QtWidgets.QLabel("Elapsed Time: 00:00:00")
        control_layout.addWidget(self.timer_label)

        main_layout.addLayout(control_layout)

        # --- Manual Website Input and Removal ---
        manual_layout = QtWidgets.QHBoxLayout()
        manual_label = QtWidgets.QLabel("Add Website Manually:")
        manual_layout.addWidget(manual_label)
        self.website_input = QtWidgets.QLineEdit()
        self.website_input.setPlaceholderText("Enter website URL (e.g., example.com)")
        manual_layout.addWidget(self.website_input)
        self.add_button = QtWidgets.QPushButton("Add Website")
        self.add_button.clicked.connect(self.add_website_manually)
        manual_layout.addWidget(self.add_button)
        self.remove_button = QtWidgets.QPushButton("Remove Selected Website")
        self.remove_button.clicked.connect(self.remove_selected_website)
        manual_layout.addWidget(self.remove_button)
        main_layout.addLayout(manual_layout)

        # --- Table to Display Website Results ---
        self.table = QtWidgets.QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Website", "Email Count", "Emails", "Mobile Numbers", "WhatsApp Links"])
        self.table.horizontalHeader().setStretchLastSection(True)
        main_layout.addWidget(self.table)

        # --- Log Text Area ---
        self.log_text = QtWidgets.QTextEdit()
        self.log_text.setReadOnly(True)
        main_layout.addWidget(self.log_text)

        # --- Credits ---
        self.credit_label = QtWidgets.QLabel("Concept and development by Saurav Shriwastav")
        self.credit_label.setAlignment(QtCore.Qt.AlignCenter)
        self.credit_label.setStyleSheet("color: blue; font-weight: bold;")
        main_layout.addWidget(self.credit_label)

        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.update_timer)

    def add_website_manually(self):
        url = self.website_input.text().strip()
        if not url:
            QtWidgets.QMessageBox.warning(self, "Input Error", "Please enter a website URL.")
            return
        if not url.startswith("http"):
            url = "https://" + url
        parsed = urlparse(url)
        if not parsed.netloc:
            QtWidgets.QMessageBox.warning(self, "Input Error", "The entered URL is not valid.")
            return
        if url in self.websites:
            QtWidgets.QMessageBox.information(self, "Duplicate", "This website is already added.")
            self.website_input.clear()
            return
        self.websites.append(url)
        self.log(f"Manually added website: {url}")
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(url))
        self.table.setItem(row, 1, QtWidgets.QTableWidgetItem("0"))
        self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(""))
        self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(""))
        self.table.setItem(row, 4, QtWidgets.QTableWidgetItem(""))
        self.website_input.clear()
        if self.websites:
            self.start_button.setEnabled(True)

    def remove_selected_website(self):
        selected_rows = set()
        for item in self.table.selectedItems():
            selected_rows.add(item.row())
        if not selected_rows:
            QtWidgets.QMessageBox.information(self, "Remove Website", "Please select a website to remove.")
            return
        for row in sorted(selected_rows, reverse=True):
            website_item = self.table.item(row, 0)
            if website_item:
                website = website_item.text().strip()
                if website in self.websites:
                    self.websites.remove(website)
                if website in self.results:
                    del self.results[website]
            self.table.removeRow(row)
        self.log("Selected website(s) removed.")

    def load_excel(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            try:
                df = pd.read_excel(file_path)

                def is_url(s):
                    if not isinstance(s, str):
                        return False
                    s = s.strip()
                    if not s:
                        return False
                    url_pattern = re.compile(r'^(https?://)?(www\.)?[\w-]+\.[\w.-]+')
                    return bool(url_pattern.match(s))

                detected_col = None
                max_ratio = 0.0
                for col in df.columns:
                    non_null = df[col].dropna().astype(str)
                    if len(non_null) == 0:
                        continue
                    url_count = sum(1 for x in non_null if is_url(x))
                    ratio = url_count / len(non_null)
                    if ratio > max_ratio:
                        max_ratio = ratio
                        detected_col = col

                if detected_col is None or max_ratio == 0:
                    QtWidgets.QMessageBox.critical(
                        self, "Error", "No column containing website URLs was detected in the Excel file."
                    )
                    return

                new_websites = df[detected_col].dropna().astype(str).tolist()
                self.log(f"Loaded {len(new_websites)} website(s) from column '{detected_col}'.")
                for website in new_websites:
                    website = website.strip()
                    if not website.startswith("http"):
                        website = "https://" + website
                    if website not in self.websites:
                        self.websites.append(website)
                        row = self.table.rowCount()
                        self.table.insertRow(row)
                        self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(website))
                        self.table.setItem(row, 1, QtWidgets.QTableWidgetItem("0"))
                        self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(""))
                        self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(""))
                        self.table.setItem(row, 4, QtWidgets.QTableWidgetItem(""))
                self.start_button.setEnabled(True)
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load Excel file: {e}")

    def start_crawling(self):
        if not self.websites:
            QtWidgets.QMessageBox.warning(self, "Warning", "No websites loaded.")
            return

        # Only crawl new websites (skip already crawled ones).
        new_sites = [site for site in self.websites if site not in self.results]
        if not new_sites:
            QtWidgets.QMessageBox.information(self, "Information", "All websites have already been crawled.")
            return

        self.log("Starting crawling process for new websites...")
        # Preserve previous results.
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.export_button.setEnabled(False)
        self.bulk_email_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.timer.start(1000)
        self.start_time = time.time()

        # Get the default region from the combo box.
        region = self.region_combo.currentText()

        # Retrieve the user-specified email and phone limits.
        email_limit_text = self.email_limit_edit.text().strip()
        if email_limit_text.isdigit():
            email_limit = int(email_limit_text)
        else:
            email_limit = None

        phone_limit_text = self.phone_limit_edit.text().strip()
        if phone_limit_text.isdigit():
            phone_limit = int(phone_limit_text)
        else:
            phone_limit = None

        self.worker = CrawlerWorker(
            new_sites,
            max_pages=100,
            max_depth=0,
            dynamic_crawl=self.dynamic_checkbox.isChecked(),
            region=region,
            email_limit=email_limit,
            phone_limit=phone_limit
        )
        self.worker.log_signal.connect(self.log)
        self.worker.website_done_signal.connect(self.update_table)
        self.worker.finished_signal.connect(self.crawling_finished)
        self.worker.start()

    def stop_crawling(self):
        if self.worker:
            self.worker.stop()
            self.log("Stop signal sent. Waiting for crawling to halt...")
            self.stop_button.setEnabled(False)

    def crawling_finished(self):
        self.log("Crawling finished.")
        self.progress_bar.setVisible(False)
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.export_button.setEnabled(True)
        self.bulk_email_button.setEnabled(True)
        self.timer.stop()

    def get_whatsapp_link(self, phone):
        """Create a WhatsApp link from a normalized phone number."""
        if phone.startswith("+"):
            return f"https://wa.me/{phone[1:]}"
        return f"https://wa.me/{phone}"

    def update_table(self, website, emails, mobiles):
        self.results[website] = (emails, mobiles)
        wa_links = [self.get_whatsapp_link(m) for m in mobiles]
        for row in range(self.table.rowCount()):
            if self.table.item(row, 0).text().strip() == website:
                email_count = len(emails)
                self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(email_count)))
                self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(", ".join(emails)))
                self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(", ".join(mobiles)))
                # Create a QLabel with clickable WhatsApp links.
                wa_html = ""
                for m in sorted(mobiles):
                    url = self.get_whatsapp_link(m)
                    wa_html += f'<a href="{url}">{m}</a><br>'
                wa_label = QtWidgets.QLabel(wa_html)
                wa_label.setOpenExternalLinks(True)
                wa_label.setTextFormat(QtCore.Qt.RichText)
                wa_label.setWordWrap(True)
                self.table.setCellWidget(row, 4, wa_label)
                break

    def export_results(self):
        if not self.results:
            QtWidgets.QMessageBox.warning(self, "Warning", "No results to export.")
            return
        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save Results", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            try:
                export_data = []
                for website, (emails, mobiles) in self.results.items():
                    wa_links = [self.get_whatsapp_link(m) for m in mobiles]
                    export_data.append({
                        "Website": website,
                        "Emails": ", ".join(emails),
                        "Email Count": len(emails),
                        "Mobile Numbers": ", ".join(mobiles),
                        "WhatsApp Links": ", ".join(wa_links),
                        "Mobile Count": len(mobiles)
                    })
                df = pd.DataFrame(export_data)
                df.to_excel(file_path, index=False)
                QtWidgets.QMessageBox.information(self, "Success", f"Results exported successfully to {file_path}")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to export results: {e}")

    def send_bulk_email(self):
        all_emails = set()
        for emails, _ in self.results.values():
            all_emails.update(emails)
        if not all_emails:
            QtWidgets.QMessageBox.information(self, "No Emails", "No email addresses available to send.")
            return
        mailto_link = "mailto:" + ",".join(all_emails)
        QDesktopServices.openUrl(QUrl(mailto_link))

    def log(self, message):
        current_time = time.strftime("%H:%M:%S")
        self.log_text.append(f"[{current_time}] {message}")

    def update_timer(self):
        if self.start_time:
            elapsed = int(time.time() - self.start_time)
            hours, remainder = divmod(elapsed, 3600)
            minutes, seconds = divmod(remainder, 60)
            self.timer_label.setText(f"Elapsed Time: {hours:02}:{minutes:02}:{seconds:02}")

# ------------------------------------------------------------------------------
# Main entry point with Splash Screen
# ------------------------------------------------------------------------------
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    # Set the application icon (for taskbar, etc.)
    app.setWindowIcon(QtGui.QIcon("EnP.ico"))

    # Create and show the splash screen.
    splash_pix = QtGui.QPixmap("EnP.ico")
    splash = QtWidgets.QSplashScreen(splash_pix, QtCore.Qt.WindowStaysOnTopHint)
    splash.setWindowFlag(QtCore.Qt.FramelessWindowHint)
    splash.showMessage("Developed by Saurav Shriwastav", QtCore.Qt.AlignBottom | QtCore.Qt.AlignCenter, QtCore.Qt.white)
    splash.show()
    app.processEvents()

    # (Optional) simulate loading delay if needed.
    time.sleep(2)

    # Create and show the main window.
    window = MainWindow()
    window.show()

    # Close the splash screen once the main window is ready.
    splash.finish(window)

    sys.exit(app.exec_())
