from importlib.resources import path
import time
import os
import shutil
from datetime import datetime
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager # ALWAYS USE THIS TO INSTALL CHROME DRIVER
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import PySimpleGUI as sg

###################################### Main Scrape Functions ######################################

def aal_report_scrape(username, password, folder, window):
    credentials = f'{username}:{password}@'

    # Set download options and driver
    options = Options()
    download_options = {
      "download.prompt_for_download": False,
      "plugins.always_open_pdf_externally": True,
      "download.open_pdf_in_system_reader": False,
      "profile.default_content_settings.popups": 0,
    }
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_experimental_option("prefs", download_options)
    # options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    print("\n")
    folder = folder.replace("/", "\\")
    # params = {'behavior': 'allow', 'downloadPath': f"{folder}\\Books\\AAL\\EOM SFG Reports"} # Use this when running in headless mode
    # driver.execute_cdp_cmd('Page.setDownloadBehavior', params) # Use this when running in headless mode

    ############################## AAL Section ##############################

    # Track AAL process start time
    start_time = time.time()

    # Log in using AAL ID
    aal_login = f'https://{credentials}ssl.drgnetwork.com/client/aal/app/live/partnerportal?'
    driver.get(aal_login)
    time.sleep(2)

    # Navigate to current month then the first report
    driver.get("https://ssl.drgnetwork.com/client/aal/app/live/eomreports?org=AAL")
    time.sleep(3)
    driver.find_element(By.XPATH, "*//button[contains(text(), 'Submit')]").click()

    # Extract current period
    current_period = driver.find_element(By.XPATH, "*//div/select/option[1][@value]")
    current_period = current_period.get_attribute("value")
    print(f"Current Period: {current_period}", "\n")
    time.sleep(3)

    # Loop through all PDF versions of reports and append needed reports to list
    print("Gathering AAL PDF reports...", "\n")
    aal_download_list = []

    needed_pdf_reports = [
      "AR1030.pdf", "AR1031.pdf", "AR1106.pdf", "AR2027.pdf", "AR2050.pdf", "CAT005ALL.pdf", "CAT005.pdf", "CAT014.pdf", "CAT1401.pdf", "CAT1401C.pdf", "CAT1402.pdf", "CAT142.pdf", "CAT159.pdf", "CN3006R.pdf", "CN3008R.pdf", "CN3022R.pdf", "CN3026R.pdf", "CN3038R.pdf", "DM3030.pdf", "PRD049.pdf", "PRD056.pdf", "PRD377R.pdf", "SUB242.pdf", "SUB321.pdf", "SUB417.pdf", "SUB418.pdf", "SUB577.pdf", "SUB580.pdf", "WH1021.pdf", "WH1025.pdf", "WH1032.pdf"
    ]
    aal_pdf_reports_on_site = driver.find_elements(By.XPATH, "*//tr/td/a")
    for report in aal_pdf_reports_on_site:
      hyperlink = report.get_attribute("href")
      if hyperlink.endswith(tuple(needed_pdf_reports)) == True:
        aal_download_list.append(hyperlink)
    time.sleep(2)

    # Get Download Directory of current EOM
    current_download_dir = f"https://ssl.drgnetwork.com/client/AAL/files/monthend/{current_period}/csv/" # Change for each company
    driver.get(current_download_dir)

    # Loop through all CSV versions of reports and append needed reports to list
    print("Gathering AAL CSV reports...", "\n")
    needed_csv_reports = [
      "CAT005ALL.csv", "CAT005CSH.csv", "CAT005_F_", "CAT014ALL.csv", "CAT1401.csv", "CAT1401C.csv", "CAT1410.csv", "CAT701.csv", "CN3006R",
      "CN3008R", "CN3038R", "DW5010RMS", "LP1001R", "PRD049Aging", "PRD056", "PRD580R", "REFLIST", "SUB242", "SUB418", "SUB491", "TAC3007_DI"
    ]
    time.sleep(1)
    aal_csv_reports_on_site = driver.find_elements(By.XPATH, "*//li/a")
    for report in aal_csv_reports_on_site:
      hyperlink = report.get_attribute("href")
      report_name = str(report.text)
      if report_name.startswith(tuple(needed_csv_reports)) == True:
        aal_download_list.append(hyperlink)
    time.sleep(1)

    # Dowload all files
    for download in aal_download_list:
      driver.get(download)
      print(f"Downloaded {download}")
      time.sleep(.8)

    # Move and rename files that were downloaded in the past 15 minutes
    SECONDS_IN_15_MINS = 60 * 15
    source = f"{os.path.expanduser('~')}\Downloads"
    destination = f"{folder}\\Books\\AAL\\EOM SFG Reports" # CHANGE FOR EACH COMPANY

    now = time.time()
    before = now - SECONDS_IN_15_MINS
    now_datetime = datetime.now()
    month_day_year = now_datetime.strftime("%m%d%Y") # MMDDYYYY format

    def last_mod_time(filename):
        return os.path.getmtime(filename)

    for filename in os.listdir(source):
      src_fname = os.path.join(source, filename)
      if last_mod_time(src_fname) > before and filename.startswith(tuple(needed_csv_reports)) == True or filename.endswith(tuple(needed_pdf_reports)) == True:
        dst_fname = os.path.join(destination, f"{filename[:-4]}_{month_day_year}{filename[-4:]}")
        shutil.move(src_fname, dst_fname)
        print(f"Moved {src_fname} to {dst_fname}")

    run_time = round(((time.time() - start_time)/60), 2)
    print("\n", f"Process Finished for AAL --- {run_time} Minutes ---", "\n")

    driver.quit()

    window.write_event_value('-THREAD DONE-', '')


def nsl_report_scrape(username, password, folder, window):
    credentials = f'{username}:{password}@'

    # Set download options and driver
    options = Options()
    download_options = {
      "download.prompt_for_download": False,
      "plugins.always_open_pdf_externally": True,
      "download.open_pdf_in_system_reader": False,
      "profile.default_content_settings.popups": 0,
    }
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_experimental_option("prefs", download_options)
    # options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    print("\n")
    folder = folder.replace("/", "\\")
    # params = {'behavior': 'allow', 'downloadPath': f"{folder}\\Books\\NSL\\EOM SFG Reports"} # Use this when running in headless mode
    # driver.execute_cdp_cmd('Page.setDownloadBehavior', params) # Use this when running in headless mode

    ############################## NSL Section ##############################

    # Track NSL process start time
    start_time = time.time()

    # Log in using NSL ID
    nsl_login = f'https://{credentials}ssl.drgnetwork.com/client/nsl/app/live/partnerportal?'
    driver.get(nsl_login)
    time.sleep(2)

    # Navigate to current month then the first report
    driver.get("https://ssl.drgnetwork.com/client/nsl/app/live/eomreports?org=NSL")
    time.sleep(3)
    driver.find_element(By.XPATH, "*//button[contains(text(), 'Submit')]").click()

    # Extract current period (this only needs done once for whole script, this can be deleted for all section except AAL)
    current_period = driver.find_element(By.XPATH, "*//div/select/option[1][@value]")
    current_period = current_period.get_attribute("value")
    print(f"Current Period: {current_period}", "\n")
    time.sleep(3)

    # Loop through all PDF versions of reports and append needed reports to list
    print("Gathering NSL PDF reports...", "\n")
    nsl_download_list = []

    needed_pdf_reports = [
      "AR1030.pdf", "AR1031.pdf", "AR1106.pdf", "AR2050.pdf", "CAT014.pdf", "CAT159.pdf", "CN3008R.pdf", "CN3022R.pdf", "CN3026R.pdf", "CN3038R.pdf",
      "DM3030.pdf", "PRD049.pdf", "PRD056.pdf", "SUB321.pdf", "SUB417.pdf", "SUB418.pdf", "SUB577.pdf", "SUB580.pdf", "WH1021.pdf", "WH1025.pdf", "WH1032.pdf"
    ]
    nsl_pdf_reports_on_site = driver.find_elements(By.XPATH, "*//tr/td/a")
    for report in nsl_pdf_reports_on_site:
      hyperlink = report.get_attribute("href")
      if hyperlink.endswith(tuple(needed_pdf_reports)) == True:
        nsl_download_list.append(hyperlink)
    time.sleep(2)

    # Get Download Directory of current EOM
    current_download_dir = f"https://ssl.drgnetwork.com/client/NSL/files/monthend/{current_period}/csv/" # Change for each company
    driver.get(current_download_dir)

    # Loop through all CSV versions of reports and append needed reports to list
    print("Gathering NSL CSV reports...", "\n")
    needed_csv_reports = [
      "CAT005CSH.csv", "PRD049Aging", "SUB418"
    ]
    time.sleep(1)
    nsl_csv_reports_on_site = driver.find_elements(By.XPATH, "*//li/a")
    for report in nsl_csv_reports_on_site:
      hyperlink = report.get_attribute("href")
      report_name = str(report.text)
      if report_name.startswith(tuple(needed_csv_reports)) == True:
        nsl_download_list.append(hyperlink)
    time.sleep(1)

    # Dowload all files
    for download in nsl_download_list:
      driver.get(download)
      print(f"Downloaded {download}")
      time.sleep(.8)

    # Move and rename files that were downloaded in the past 15 minutes
    SECONDS_IN_15_MINS = 60 * 15
    source = f"{os.path.expanduser('~')}\Downloads"
    destination = f"{folder}\\Books\\NSL\\EOM SFG Reports" # CHANGE FOR EACH COMPANY

    now = time.time()
    before = now - SECONDS_IN_15_MINS
    now_datetime = datetime.now()
    month_day_year = now_datetime.strftime("%m%d%Y") # MMDDYYYY format

    def last_mod_time(filename):
        return os.path.getmtime(filename)

    for filename in os.listdir(source):
      src_fname = os.path.join(source, filename)
      if last_mod_time(src_fname) > before and filename.startswith(tuple(needed_csv_reports)) == True or filename.endswith(tuple(needed_pdf_reports)) == True:
        dst_fname = os.path.join(destination, f"{filename[:-4]}_{month_day_year}{filename[-4:]}")
        shutil.move(src_fname, dst_fname)
        print(f"Moved {src_fname} to {dst_fname}")

    run_time = round(((time.time() - start_time)/60), 2)
    print("\n", f"Process Finished for NSL --- {run_time} Minutes ---", "\n")

    driver.quit()

    window.write_event_value('-THREAD DONE-', '')


def hwb_report_scrape(username, password, folder, window):
    credentials = f'{username}:{password}@'

    # Set download options and driver
    options = Options()
    download_options = {
      "download.prompt_for_download": False,
      "plugins.always_open_pdf_externally": True,
      "download.open_pdf_in_system_reader": False,
      "profile.default_content_settings.popups": 0,
    }
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_experimental_option("prefs", download_options)
    # options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    print("\n")
    folder = folder.replace("/", "\\")
    # params = {'behavior': 'allow', 'downloadPath': f"{folder}\\Books\\AAL\\EOM SFG Reports"} # Use this when running in headless mode
    # driver.execute_cdp_cmd('Page.setDownloadBehavior', params) # Use this when running in headless mode

    ############################## HWB Section ##############################

    # Track HWB process start time
    start_time = time.time()

    # Log in using HWB ID
    hwb_login = f'https://{credentials}ssl.drgnetwork.com/client/hwb/app/live/partnerportal?'
    driver.get(hwb_login)
    time.sleep(2)

    # Navigate to current month then the first report
    driver.get("https://ssl.drgnetwork.com/client/hwb/app/live/eomreports?org=HWB")
    time.sleep(5)
    driver.find_element(By.XPATH, "*//button[contains(text(), 'Submit')]").click()

    # Extract current period
    current_period = driver.find_element(By.XPATH, "*//div/select/option[1][@value]")
    current_period = current_period.get_attribute("value")
    print(f"Current Period: {current_period}", "\n")
    time.sleep(3)

    # Loop through all PDF versions of reports and append needed reports to list
    print("Gathering HWB PDF reports...", "\n")
    hwb_download_list = []

    needed_pdf_reports = [
      "AR1030.pdf", "AR1031.pdf", "AR1106.pdf", "AR2050.pdf", "CAT005ALL.pdf", "CAT005.pdf", "CAT014.pdf", "CAT1401.pdf", "CAT1401C.pdf", "CAT142.pdf", "CAT159.pdf", "CN3008R.pdf", "CN3022R.pdf", "CN3026R.pdf", "CN3038R.pdf", "DM3030.pdf", "PRD049.pdf", "PRD056.pdf", "SUB321.pdf", "SUB417.pdf", "SUB418.pdf", "SUB577.pdf", "SUB580.pdf", "WH1021.pdf", "WH1025.pdf", "WH1032.pdf"
    ]
    hwb_pdf_reports_on_site = driver.find_elements(By.XPATH, "*//tr/td/a")
    for report in hwb_pdf_reports_on_site:
      hyperlink = report.get_attribute("href")
      if hyperlink.endswith(tuple(needed_pdf_reports)) == True:
        hwb_download_list.append(hyperlink)
    time.sleep(2)

    # Get Download Directory of current EOM
    current_download_dir = f"https://ssl.drgnetwork.com/client/HWB/files/monthend/{current_period}/csv/" # Change for each company
    driver.get(current_download_dir)

    # Loop through all CSV versions of reports and append needed reports to list
    print("Gathering HWB CSV reports...", "\n")
    needed_csv_reports = [
      "CAT005ALL.csv", "CAT005CSH.csv", "CAT005_F_", "CAT1401.csv", "CAT1401C.csv", "PRD049Aging", "SUB418"
    ]
    time.sleep(1)
    hwb_csv_reports_on_site = driver.find_elements(By.XPATH, "*//li/a")
    for report in hwb_csv_reports_on_site:
      hyperlink = report.get_attribute("href")
      report_name = str(report.text)
      if report_name.startswith(tuple(needed_csv_reports)) == True:
        hwb_download_list.append(hyperlink)
    time.sleep(1)

    # Dowload all files
    for download in hwb_download_list:
      driver.get(download)
      print(f"Downloaded {download}")
      time.sleep(.8)

    # Move and rename files that were downloaded in the past 15 minutes
    SECONDS_IN_15_MINS = 60 * 15
    source = f"{os.path.expanduser('~')}\Downloads"
    destination = f"{folder}\\Books\\HWB\\EOM SFG Reports" # CHANGE FOR EACH COMPANY

    now = time.time()
    before = now - SECONDS_IN_15_MINS
    now_datetime = datetime.now()
    month_day_year = now_datetime.strftime("%m%d%Y") # MMDDYYYY format

    def last_mod_time(filename):
        return os.path.getmtime(filename)

    for filename in os.listdir(source):
      src_fname = os.path.join(source, filename)
      if last_mod_time(src_fname) > before and filename.startswith(tuple(needed_csv_reports)) == True or filename.endswith(tuple(needed_pdf_reports)) == True:
        dst_fname = os.path.join(destination, f"{filename[:-4]}_{month_day_year}{filename[-4:]}")
        shutil.move(src_fname, dst_fname)
        print(f"Moved {src_fname} to {dst_fname}")

    run_time = round(((time.time() - start_time)/60), 2)
    print("\n", f"Process Finished for HWB --- {run_time} Minutes ---", "\n")

    driver.quit()

    window.write_event_value('-THREAD DONE-', '')


def csl_report_scrape(username, password, folder, window):
    credentials = f'{username}:{password}@'

    # Set download options and driver
    options = Options()
    download_options = {
      "download.prompt_for_download": False,
      "plugins.always_open_pdf_externally": True,
      "download.open_pdf_in_system_reader": False,
      "profile.default_content_settings.popups": 0,
    }
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_experimental_option("prefs", download_options)
    # options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    print("\n")
    folder = folder.replace("/", "\\")
    # params = {'behavior': 'allow', 'downloadPath': f"{folder}\\Books\\CSL\\EOM SFG Reports"} # Use this when running in headless mode
    # driver.execute_cdp_cmd('Page.setDownloadBehavior', params) # Use this when running in headless mode

    ############################## AAL Section ##############################

    # Track AAL process start time
    start_time = time.time()

    # Log in using AAL ID
    csl_login = f'https://{credentials}ssl.drgnetwork.com/client/csl/app/live/partnerportal?'
    driver.get(csl_login)
    time.sleep(2)

    # Navigate to current month then the first report
    driver.get("https://ssl.drgnetwork.com/client/csl/app/live/eomreports?org=CSL")
    time.sleep(3)
    driver.find_element(By.XPATH, "*//button[contains(text(), 'Submit')]").click()

    # Extract current period
    current_period = driver.find_element(By.XPATH, "*//div/select/option[1][@value]")
    current_period = current_period.get_attribute("value")
    print(f"Current Period: {current_period}", "\n")
    time.sleep(3)

    # Loop through all PDF versions of reports and append needed reports to list
    print("Gathering CSL PDF reports...", "\n")
    csl_download_list = []

    needed_pdf_reports = [
      "AR1030.pdf", "AR1031.pdf", "AR1106.pdf", "AR2027.pdf", "AR2050.pdf", "CAT005ALL.pdf", "CAT005.pdf", "CAT1401.pdf", 
      "CAT1401C.pdf", "CAT142.pdf", "CAT159.pdf", "DM3030.pdf", "PRD049.pdf", "PRD056.pdf", "SUB242.pdf", "SUB321.pdf", "SUB417.pdf", "SUB418.pdf", "SUB577.pdf", "SUB580.pdf",
      "WH1021.pdf", "WH1025.pdf", "WH1032.pdf"
    ]
    csl_pdf_reports_on_site = driver.find_elements(By.XPATH, "*//tr/td/a")
    for report in csl_pdf_reports_on_site:
      hyperlink = report.get_attribute("href")
      if hyperlink.endswith(tuple(needed_pdf_reports)) == True:
        csl_download_list.append(hyperlink)
    time.sleep(2)

    # Get Download Directory of current EOM
    current_download_dir = f"https://ssl.drgnetwork.com/client/CSL/files/monthend/{current_period}/csv/" # Change for each company
    driver.get(current_download_dir)

    # Loop through all CSV versions of reports and append needed reports to list
    print("Gathering CSL CSV reports...", "\n")
    needed_csv_reports = [
      "CAT005ALL.csv", "CAT005CSH.csv", "CAT005_F_", "CAT1401.csv", "CAT1401C.csv", "CAT1410.csv", "CAT701.csv", "PRD049Aging", "PRD056", "REFLIST", "SUB242", "SUB418", "SUB491_CS", "SUB491_FS", "SUB491_S2"
    ]
    time.sleep(1)
    csl_csv_reports_on_site = driver.find_elements(By.XPATH, "*//li/a")
    for report in csl_csv_reports_on_site:
      hyperlink = report.get_attribute("href")
      report_name = str(report.text)
      if report_name.startswith(tuple(needed_csv_reports)) == True:
        csl_download_list.append(hyperlink)
    time.sleep(1)

    # Dowload all files
    for download in csl_download_list:
      driver.get(download)
      print(f"Downloaded {download}")
      time.sleep(.8)

    # Move and rename files that were downloaded in the past 15 minutes
    SECONDS_IN_15_MINS = 60 * 15
    source = f"{os.path.expanduser('~')}\Downloads"
    destination = f"{folder}\\Books\\CSL\\EOM SFG Reports" # CHANGE FOR EACH COMPANY

    now = time.time()
    before = now - SECONDS_IN_15_MINS
    now_datetime = datetime.now()
    month_day_year = now_datetime.strftime("%m%d%Y") # MMDDYYYY format

    def last_mod_time(filename):
        return os.path.getmtime(filename)

    for filename in os.listdir(source):
      src_fname = os.path.join(source, filename)
      if last_mod_time(src_fname) > before and filename.startswith(tuple(needed_csv_reports)) == True or filename.endswith(tuple(needed_pdf_reports)) == True:
        dst_fname = os.path.join(destination, f"{filename[:-4]}_{month_day_year}{filename[-4:]}")
        shutil.move(src_fname, dst_fname)
        print(f"Moved {src_fname} to {dst_fname}")

    run_time = round(((time.time() - start_time)/60), 2)
    print("\n", f"Process Finished for CSL --- {run_time} Minutes ---", "\n")

    driver.quit()

    window.write_event_value('-THREAD DONE-', '')


def csl_report_scrape(username, password, folder, window):
    credentials = f'{username}:{password}@'

    # Set download options and driver
    options = Options()
    download_options = {
      "download.prompt_for_download": False,
      "plugins.always_open_pdf_externally": True,
      "download.open_pdf_in_system_reader": False,
      "profile.default_content_settings.popups": 0,
    }
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_experimental_option("prefs", download_options)
    # options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    print("\n")
    folder = folder.replace("/", "\\")
    # params = {'behavior': 'allow', 'downloadPath': f"{folder}\\Books\\CSL\\EOM SFG Reports"} # Use this when running in headless mode
    # driver.execute_cdp_cmd('Page.setDownloadBehavior', params) # Use this when running in headless mode

    ############################## AAL Section ##############################

    # Track AAL process start time
    start_time = time.time()

    # Log in using AAL ID
    csl_login = f'https://{credentials}ssl.drgnetwork.com/client/csl/app/live/partnerportal?'
    driver.get(csl_login)
    time.sleep(2)

    # Navigate to current month then the first report
    driver.get("https://ssl.drgnetwork.com/client/csl/app/live/eomreports?org=CSL")
    time.sleep(3)
    driver.find_element(By.XPATH, "*//button[contains(text(), 'Submit')]").click()

    # Extract current period
    current_period = driver.find_element(By.XPATH, "*//div/select/option[1][@value]")
    current_period = current_period.get_attribute("value")
    print(f"Current Period: {current_period}", "\n")
    time.sleep(3)

    # Loop through all PDF versions of reports and append needed reports to list
    print("Gathering CSL PDF reports...", "\n")
    csl_download_list = []

    needed_pdf_reports = [
      "AR1030.pdf", "AR1031.pdf", "AR1106.pdf", "AR2027.pdf", "AR2050.pdf", "CAT005ALL.pdf", "CAT005.pdf", "CAT1401.pdf", "CAT1401C.pdf", "CAT142.pdf", "CAT159.pdf", "DM3030.pdf", "PRD049.pdf", "PRD056.pdf", "SUB242.pdf", "SUB321.pdf", "SUB417.pdf", "SUB418.pdf", "SUB577.pdf", "SUB580.pdf", "WH1021.pdf", "WH1025.pdf", "WH1032.pdf"
    ]
    csl_pdf_reports_on_site = driver.find_elements(By.XPATH, "*//tr/td/a")
    for report in csl_pdf_reports_on_site:
      hyperlink = report.get_attribute("href")
      if hyperlink.endswith(tuple(needed_pdf_reports)) == True:
        csl_download_list.append(hyperlink)
    time.sleep(2)

    # Get Download Directory of current EOM
    current_download_dir = f"https://ssl.drgnetwork.com/client/CSL/files/monthend/{current_period}/csv/" # Change for each company
    driver.get(current_download_dir)

    # Loop through all CSV versions of reports and append needed reports to list
    print("Gathering CSL CSV reports...", "\n")
    needed_csv_reports = [
      "CAT005ALL.csv", "CAT005CSH.csv", "CAT005_F_", "CAT1401.csv", "CAT1401C.csv", "CAT1410.csv", "CAT701.csv", "PRD049Aging", "PRD056", "REFLIST", "SUB242", "SUB418", "SUB491_CS", "SUB491_FS", "SUB491_S2"
    ]
    time.sleep(1)
    csl_csv_reports_on_site = driver.find_elements(By.XPATH, "*//li/a")
    for report in csl_csv_reports_on_site:
      hyperlink = report.get_attribute("href")
      report_name = str(report.text)
      if report_name.startswith(tuple(needed_csv_reports)) == True:
        csl_download_list.append(hyperlink)
    time.sleep(1)

    # Dowload all files
    for download in csl_download_list:
      driver.get(download)
      print(f"Downloaded {download}")
      time.sleep(.8)

    # Move and rename files that were downloaded in the past 15 minutes
    SECONDS_IN_15_MINS = 60 * 15
    source = f"{os.path.expanduser('~')}\Downloads"
    destination = f"{folder}\\Books\\CSL\\EOM SFG Reports" # CHANGE FOR EACH COMPANY

    now = time.time()
    before = now - SECONDS_IN_15_MINS
    now_datetime = datetime.now()
    month_day_year = now_datetime.strftime("%m%d%Y") # MMDDYYYY format

    def last_mod_time(filename):
        return os.path.getmtime(filename)

    for filename in os.listdir(source):
      src_fname = os.path.join(source, filename)
      if last_mod_time(src_fname) > before and filename.startswith(tuple(needed_csv_reports)) == True or filename.endswith(tuple(needed_pdf_reports)) == True:
        dst_fname = os.path.join(destination, f"{filename[:-4]}_{month_day_year}{filename[-4:]}")
        shutil.move(src_fname, dst_fname)
        print(f"Moved {src_fname} to {dst_fname}")

    run_time = round(((time.time() - start_time)/60), 2)
    print("\n", f"Process Finished for CSL --- {run_time} Minutes ---", "\n")

    driver.quit()

    window.write_event_value('-THREAD DONE-', '')


def clo_report_scrape(username, password, folder, window):
    credentials = f'{username}:{password}@'

    # Set download options and driver
    options = Options()
    download_options = {
      "download.prompt_for_download": False,
      "plugins.always_open_pdf_externally": True,
      "download.open_pdf_in_system_reader": False,
      "profile.default_content_settings.popups": 0,
    }
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_experimental_option("prefs", download_options)
    # options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    print("\n")
    folder = folder.replace("/", "\\")
    # params = {'behavior': 'allow', 'downloadPath': f"{folder}\\Books\\CLO\\EOM SFG Reports"} # Use this when running in headless mode
    # driver.execute_cdp_cmd('Page.setDownloadBehavior', params) # Use this when running in headless mode

    ############################## CLO Section ##############################

    # Track AAL process start time
    start_time = time.time()

    # Log in using AAL ID
    clo_login = f'https://{credentials}ssl.drgnetwork.com/client/clo/app/live/partnerportal?'
    driver.get(clo_login)
    time.sleep(2)

    # Navigate to current month then the first report
    driver.get("https://ssl.drgnetwork.com/client/clo/app/live/eomreports?org=CLO")
    time.sleep(3)
    driver.find_element(By.XPATH, "*//button[contains(text(), 'Submit')]").click()

    # Extract current period
    current_period = driver.find_element(By.XPATH, "*//div/select/option[1][@value]")
    current_period = current_period.get_attribute("value")
    print(f"Current Period: {current_period}", "\n")
    time.sleep(3)

    # Loop through all PDF versions of reports and append needed reports to list
    print("Gathering CLO PDF reports...", "\n")
    clo_download_list = []

    needed_pdf_reports = [
      "CAT159.pdf", "PRD377R.pdf"
    ]
    clo_pdf_reports_on_site = driver.find_elements(By.XPATH, "*//tr/td/a")
    for report in clo_pdf_reports_on_site:
      hyperlink = report.get_attribute("href")
      if hyperlink.endswith(tuple(needed_pdf_reports)) == True:
        clo_download_list.append(hyperlink)
    time.sleep(2)

    # Get Download Directory of current EOM
    current_download_dir = f"https://ssl.drgnetwork.com/client/CLO/files/monthend/{current_period}/csv/" # Change for each company
    driver.get(current_download_dir)

    # Dowload all files
    for download in clo_download_list:
      driver.get(download)
      print(f"Downloaded {download}")
      time.sleep(.8)

    # Move and rename files that were downloaded in the past 15 minutes
    SECONDS_IN_15_MINS = 60 * 15
    source = f"{os.path.expanduser('~')}\Downloads"
    destination = f"{folder}\\Books\\CLO\\EOM SFG Reports" # CHANGE FOR EACH COMPANY

    now = time.time()
    before = now - SECONDS_IN_15_MINS
    now_datetime = datetime.now()
    month_day_year = now_datetime.strftime("%m%d%Y") # MMDDYYYY format

    def last_mod_time(filename):
        return os.path.getmtime(filename)

    for filename in os.listdir(source):
      src_fname = os.path.join(source, filename)
      if last_mod_time(src_fname) > before and filename.endswith(tuple(needed_pdf_reports)) == True:
        dst_fname = os.path.join(destination, f"{filename[:-4]}_{month_day_year}{filename[-4:]}")
        shutil.move(src_fname, dst_fname)
        print(f"Moved {src_fname} to {dst_fname}")

    run_time = round(((time.time() - start_time)/60), 2)
    print("\n", f"Process Finished for CLO --- {run_time} Minutes ---", "\n")

    driver.quit()

    window.write_event_value('-THREAD DONE-', '')


###################################### GUI Definition ######################################

def the_gui():
  sg.theme("LightBlue3")
  layout = [
      [sg.Push(), sg.Text("SFG Partner Site Username:"), sg.Input(key="-USER-"), sg.Push()],
      [sg.Push(), sg.Text("SFG Partner Site Password:"), sg.Input(key="-PASSWORD-"), sg.Push()],
      [sg.Push(), sg.Text('Browse for EOP Workpapers folder:'), sg.Input(key="-FOLDER-"), sg.FolderBrowse(), sg.Push()],
      [sg.Output(size=(170,30), key="-OUTPUT-")],
      [sg.Push(),
        sg.Button("Pull AAL Reports"),
        sg.Button("Pull NSL Reports"),
        sg.Button("Pull HWB Reports"),
        sg.Button("Pull CSL Reports"),
        sg.Button("Pull CLO Reports"),
        sg.Button("Exit"),
          sg.Push()]
  ]

  window = sg.Window("SFG Report Scraper", layout, resizable=True) # "use_custom_titlebar=True" causes there to be no icon on Windows

  while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
      break
    elif event == ('Pull AAL Reports'):
      window['-OUTPUT-'].update('Accessing SFG Partner Site (Chrome popup and chromedriver.exe popup windows can be minimized)...',"\n")
      window.perform_long_operation(lambda: aal_report_scrape(
              values['-USER-'],
              values['-PASSWORD-'],
              values['-FOLDER-'],
              window
                ), '-OPERATION DONE-')
    elif event == ('Pull NSL Reports'):
      window['-OUTPUT-'].update('Accessing SFG Partner Site (Chrome popup and chromedriver.exe popup windows can be minimized)...', "\n")
      window.perform_long_operation(lambda: nsl_report_scrape(
              values['-USER-'],
              values['-PASSWORD-'],
              values['-FOLDER-'],
              window
                ), '-OPERATION DONE-')
    elif event == ('Pull HWB Reports'):
      window['-OUTPUT-'].update('Accessing SFG Partner Site (Chrome popup and chromedriver.exe popup windows can be minimized)...')
      window.perform_long_operation(lambda: hwb_report_scrape(
              values['-USER-'],
              values['-PASSWORD-'],
              values['-FOLDER-'],
              window
                ), '-OPERATION DONE-')
    elif event == ('Pull CSL Reports'):
      window['-OUTPUT-'].update('Accessing SFG Partner Site (Chrome popup and chromedriver.exe popup windows can be minimized)...')
      window.perform_long_operation(lambda: csl_report_scrape(
              values['-USER-'],
              values['-PASSWORD-'],
              values['-FOLDER-'],
              window
                ), '-OPERATION DONE-')
    elif event == ('Pull CLO Reports'):
      window['-OUTPUT-'].update('Accessing SFG Partner Site (Chrome popup and chromedriver.exe popup windows can be minimized)...')
      window.perform_long_operation(lambda: clo_report_scrape(
              values['-USER-'],
              values['-PASSWORD-'],
              values['-FOLDER-'],
              window
                ), '-OPERATION DONE-')

  window.close()

if __name__ == '__main__':
    the_gui()
    print('Exiting Program')
