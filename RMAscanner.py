#!venv/bin/python3.9
import threading
import time
import os
import sys
import pandas as pd
import gspread
import gspread_dataframe as gd

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from subprocess import CREATE_NO_WINDOW
from gooey import Gooey, GooeyParser
from oauth2client.service_account import ServiceAccountCredentials

class RMA_scanner:
    def open_google_sheets(self):
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(os.path.join(sys._MEIPASS, 'client_secret.json'), scope)
        client = gspread.authorize(creds)

        self.spread_sheet = client.open("próba efektywnosci - Automatyczne liczenie MRT")

        self.spread_sheet.worksheet("1").clear()
        self.spread_sheet.worksheet("2").clear()
        self.spread_sheet.worksheet("3").clear()

    def scanning_admins(self, website, table_id, sheet, admin_pages, other_markets):
        l = 1
        start = time.time()

        datalist = []
        while 2 <= len(other_markets):
            market = other_markets[:2]
            other_markets = other_markets.replace(market, "")
            l += 1

            print("progress: {}/{}".format(l-1, int(len(other_markets)/2 + l-1)))

            if market.lower() == 'pl':
                url = 'https://www.showroom.pl/admin/tools_rma_request_list.php?method=listByStatusAction&status='
            elif market.lower() == 'shwrm':
                url = 'https://www.showroom.pl/admin/tools_rma_request_list.php?method=listByStatusAction&status='
            elif market.lower() == 'uk':
                url = f'https://www.miinto.co.uk/admin/tools_rma_request_list.php?method=listByStatusAction&status='
            elif market.lower() == 'cn':
                url = f'https://china.miinto.net/admin/tools_rma_request_list.php?method=listByStatusAction&status='
            else:
                url = f'https://www.miinto.{market.lower()}/admin/tools_rma_request_list.php?method=listByStatusAction&status='

            # logowanie i przejście do shops
            website.get(url + str(table_id))
            website.implicitly_wait(4)

          # Logowanie adminy
            if website.find_elements(By.XPATH,"/html/body/div[2]/div/div/form/fieldset/div[4]/div/input[2]"):
                try:
                    element = WebDriverWait(website, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="password"]')))
                except:
                    website.quit()
                username = website.find_element(By.XPATH,'//*[@id="username"]').send_keys("XXX")
                password = website.find_element(By.XPATH,'//*[@id="password"]').send_keys("XXX")
                website.find_element(By.XPATH,'/html/body/div[2]/div/div/form/fieldset/div[4]/div/input[2]').click()

                website.implicitly_wait(4)
                if market.lower() == 'cn':
                    website.find_element(By.XPATH,'/html/body/div[2]/div/div/form/fieldset/table/tbody/tr[1]/td[4]/button').click()
                    website.implicitly_wait(4)
                website.get(url + str(table_id))
                time.sleep(2)

            ## Kopiowanie tabelek z adminów
            for s in range(2, int(admin_pages) + 2):
                try:
                    table = pd.read_html(website.page_source)[0]
                    table['Market'] = market
                except:
                    break
                # Uzupełnianie listy
                datalist.append(table)

                ##Kolejna strona
                try:
                    website.get(f"https://www.miinto.{market.lower()}/admin/tools_rma_request_list.php?method=listByStatusAction&status={str(table_id)}&page={str(s)}")
                    website.implicitly_wait(4)
                except:
                    break

        ## Wrzucanie tabelki do google sheet
        google_sheet = self.spread_sheet.worksheet(sheet)
        df = pd.concat(datalist)
        gd.set_with_dataframe(google_sheet, df, include_column_header=True, row=len(google_sheet.col_values(2)) + 1,
                              col=1, include_index=False, resize=False, allow_formulas=True, string_escaping="default")

        temp = time.time() - start
        hours = temp // 3600
        temp = temp - 3600 * hours
        minutes = temp // 60
        seconds = temp - 60 * minutes
        all_time = '%d:%d:%d' % (hours, minutes, seconds)
        website.quit()

        print("\n \n", "RMA scanner has finished all the tasks in", all_time, "min")

class GUI(RMA_scanner):
    chromedriver_options = None

    @Gooey(
        program_name="RMA scanner",
        program_description="Get RMA by one click!",
        terminal_font_color='black',
        terminal_panel_color='white',
        show_restart_button=False,
        progress_regex=r"^progress: (\d+)/(\d+)$",
        progress_expr="x[0] / x[1] * 100",
        disable_progress_bar_animation=False,
        default_size=(810, 530),
    )
    def handle(self):
        parser = GooeyParser()
        parser.add_argument("admin_pages", metavar="How many pages?", type=int)
        parser.add_argument("other_markets", metavar="Which markets?", widget="TextField", default="SEDKNONLBEPLITESDEFRFIUKCH")
        args = parser.parse_args()

        self.other_markets = args.other_markets
        self.admin_pages = args.admin_pages
        self.open_google_sheets()
        self.setup_chromedriver_options()

    def setup_chromedriver_options(self):
        self.chromedriver_options = webdriver.ChromeOptions()
        self.chromedriver_options.add_experimental_option(
            "excludeSwitches", ["enable-logging", "enable-automation"]
        )
        self.chromedriver_options.add_argument("--disable-extensions")
        self.chromedriver_options.add_argument("--headless")
        self.chromedriver_options.add_argument("--window-size=1920,1080")
        self.chromedriver_options.add_argument("--disable-gpu")
        prefs = {}
        prefs["profile.default_content_settings.popups"] = 0
        prefs["download.default_directory"] = os.getcwd()
        self.chromedriver_options.add_experimental_option("prefs", prefs)

        self.chrome_service1 = ChromeService(ChromeDriverManager().install())
        self.chrome_service1.creationflags = CREATE_NO_WINDOW

        self.chrome_service2 = ChromeService(ChromeDriverManager().install())
        self.chrome_service2.creationflags = CREATE_NO_WINDOW

        self.chrome_service3 = ChromeService(ChromeDriverManager().install())
        self.chrome_service3.creationflags = CREATE_NO_WINDOW
        
        self.threads()

    def threads(self):
        self.open_google_sheets()

        os.environ["WDM_LOG_LEVEL"] = "0"
        website1 = webdriver.Chrome(options=self.chromedriver_options, service=self.chrome_service1)
        website2 = webdriver.Chrome(options=self.chromedriver_options, service=self.chrome_service2)
        website3 = webdriver.Chrome(options=self.chromedriver_options, service=self.chrome_service3)

        table_L, table_A, table_R = 0, 1, 2
        sheet_L, sheet_A, sheet_R = "1", "2", "3"

        thread1 = threading.Thread(target=self.scanning_admins, args=(website1, table_L, sheet_L, self.admin_pages,self.other_markets))
        thread2 = threading.Thread(target=self.scanning_admins, args=(website2, table_A, sheet_A, self.admin_pages,self.other_markets))
        thread3 = threading.Thread(target=self.scanning_admins, args=(website3, table_R, sheet_R, self.admin_pages,self.other_markets))

        thread1.start()
        thread2.start()
        thread3.start()

        thread1.join()
        thread2.join()
        thread3.join()

if __name__ == "__main__":
    gui = GUI()
    gui.handle()
