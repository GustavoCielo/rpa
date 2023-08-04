import time
import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from PyPDF2 import PdfReader
import re
import pandas as pd


"""initial settings"""
def main():
    url = 'https://dejt.jt.jus.br/dejt/f/n/diariocon'
    option = Options()
    # option.add_argument('--headless') # here shows the browser opening, can set to true to run in background
    driver = webdriver.Chrome(options=option)

    """finding current date to get last weeks result"""
    today = datetime.date.today()
    weekday = today.weekday()
    start_delta = datetime.timedelta(days=weekday, weeks=1)
    start_of_week = today - start_delta
    week_dates = []
    for day in range(7):
        week_dates.append(start_of_week + datetime.timedelta(days=day))
    week_start = week_dates[0]
    week_end = week_dates[4]

    """opening browser"""
    driver.get(url)

    """ensuring load time"""
    time.sleep(1)

    """Gets last week results"""
    first_date = driver.find_element(By.ID, "corpo:formulario:dataIni")
    ActionChains(driver)\
        .move_to_element(first_date)\
        .pause(1)\
        .click_and_hold()\
        .pause(1)\
        .send_keys("13/02/2023")\
        .perform()
        # .send_keys(week_start.strftime('%d/%m/%Y'))\

    last_date =  driver.find_element(By.ID, "corpo:formulario:dataFim")
    ActionChains(driver)\
        .move_to_element(last_date)\
        .pause(1)\
        .click_and_hold()\
        .pause(1)\
        .send_keys("17/02/2023")\
        .perform()
        # .send_keys(week_end.strftime('%d/%m/%Y'))\
    driver.find_element(By.ID, "corpo:formulario:botaoAcaoPesquisar").click()
    """Waits for load then gets all elements with class tag"""
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CLASS_NAME, 'bt.af_commandButton')))
    downloads = driver.find_elements(By.CLASS_NAME, 'bt.af_commandButton')

    """Downloads all current page files"""
    for i in downloads:
        i.click()
        time.sleep(3)

    time.sleep(30)

    driver.quit()

    """
    Open pdf files, get all informations needed and extract text
    Have to do a for loop to open all files in sequence and read them as binary
    Must use fullpath
    """
    # name must be dynamic
    regex = re.compile("[0-9]{7}[-]?[0-9]{2}[.]?[0-9]{4}[.]?[0-9]{1}[.]?[0-9]{2}[.]?[0-9]{4}")
    pdf_to_read = 'C:/Users/User/Downloads/Diario_3666__17_2_2023.pdf'
    for i in range(0, 30):
        with open(pdf_to_read, 'rb') as f:
            # starts a reader to be able to manipulate pdf
            pdf = PdfReader(f)
            print(pdf.metadata)
            # number of pages to iterate over
            number_of_pages = len(pdf.pages)

            sheet_name = (pdf.metadata.modification_date).strftime('%d/%m/%Y').replace('/', '-')

            results = set()
            duplicates = []
            for n in range(0, number_of_pages):
                page = pdf.pages[n]
                page_content = page.extract_text()
                # creates a regex for process numbers
                res_search = re.findall(regex, page_content)
                for value in res_search:
                    # add only existing results and not duplicated
                    if res_search and value not in results:
                        results = set(value)
                    # adds duplicated results to another list
                    if res_search and value in results and value not in duplicates:
                        duplicates.append(value)

            df = pd.DataFrame(results)
            df_dup = pd.DataFrame(duplicates)

            # Write searched files in file with name based on creation date
            if len(df) > 0:
                if i >= 1:
                    with pd.ExcelWriter(f'C:/tmp/TST-{sheet_name}.xlsx', mode='a', if_sheet_exists='overlay') as writer:
                        df.to_excel(writer, index=False, header=False)
                else:
                    with pd.ExcelWriter(f'C:/tmp/TST-{sheet_name}.xlsx') as writer:
                        df.to_excel(writer, index=False, header=False)
            # write duplicates in another file
            if len(df_dup) > 0:
                if i >= 1:
                    with pd.ExcelWriter(f'C:/tmp/duplicata-TST-{sheet_name}.xlsx', mode='a', if_sheet_exists='overlay') as writer:
                        df_dup.to_excel(writer, index=False, header=False)
                else:
                    with pd.ExcelWriter(f'C:/tmp/duplicata-TST-{sheet_name}.xlsx') as writer:
                        df_dup.to_excel(writer, index=False, header=False)
            
            # name must be dynamic
            pdf_to_read = r"C:\Users\User\Downloads\Diario_3666__17_2_2023 ({}).pdf".format(i+1)

    print('RPA done')

if __name__ == '__main__':
    main()