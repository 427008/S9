from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time


chromeOptions = webdriver.ChromeOptions()
chromeOptions.add_argument('no-sandbox')
# chromeOptions.add_argument('headless')
path = r'chrome\chromedriver.exe'
driver = webdriver.Chrome(executable_path=path, options=chromeOptions)
base_url = r'http://s9bbt:57772/csp/sys/exp/_CSP.UI.Portal.'
url = base_url + r'GlobalList.cls?$NAMESPACE=BBTWORK#'

cache_globals = ['B1', 'B2', 'B11', 'B11S1', 'B59', 'BE1', 'BE5', 'BE7', 'BE7S1', 'BE70', 'BE70S1', 'BE23']
for global_name in cache_globals:
    driver.get(url)
    wait = WebDriverWait(driver, 3)
    lines_input = wait.until(EC.presence_of_element_located((By.ID, 'control_27')))
    lines_input.send_keys(Keys.BACKSPACE*5)
    lines_input.send_keys('1')
    lines_input.send_keys(Keys.ENTER)
    time.sleep(0.5)

    global_input = wait.until(EC.presence_of_element_located((By.ID, 'input_26')))
    global_input.clear()
    global_input.send_keys(Keys.BACKSPACE)
    time.sleep(0.5)
    global_input.send_keys(global_name)
    time.sleep(0.5)
    global_input.send_keys(Keys.ENTER+Keys.ESCAPE)

    global_label = wait.until(EC.text_to_be_present_in_element((By.CSS_SELECTOR, '#tr_0_52 td.tpStr'), global_name))
    global_check = wait.until(EC.element_to_be_clickable((By.ID, 'cb_0_52')))
    global_check.click()
    time.sleep(0.5)

    default_handle = driver.current_window_handle
    button_export = wait.until(EC.element_to_be_clickable((By.ID, 'command_btnExport')))
    button_export.click()

    wait.until(EC.new_window_is_opened)
    new_handle = driver.window_handles[-1]
    driver.switch_to.window(new_handle)
    window = driver.switch_to.active_element
    driver.close()

    driver.switch_to.window(default_handle)
    url_p1 = base_url + r'Dialog.ExportOutput.cls?FILETYPE=Global&FILENAME=D%3A%5CUKReports%5C'
    url_p2 = r'.gof&CHARSET=CP1251&NAMESPACE=BBTWORK&EXPORTALL=0&EXPORTFORMAT=GO&RUNBACKGROUND=0&OutputFormat=5&RecordFormat=S'
    driver.get(url_p1 + global_name + url_p2)

    wait = WebDriverWait(driver, 300)
    result_label = wait.until(EC.text_to_be_present_in_element((By.CSS_SELECTOR, 'html body pre'), 'Завершено'))

driver.quit()
