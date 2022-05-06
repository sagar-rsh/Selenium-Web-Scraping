import os
from webbrowser import Chrome
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd

# Options to bypass bot detection
options = webdriver.ChromeOptions()
options.add_argument("start-maximized")
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)


# Import/Run Chrome Driver
script_dir = os.path.dirname(os.path.realpath(__file__))
executable_path = os.path.join(script_dir, "chromedriver.exe")
browser = webdriver.Chrome(options=options,
                           executable_path=executable_path)
browser.execute_script(
    "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
browser.execute_cdp_cmd('Network.setUserAgentOverride', {"userAgent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                                                         'AppleWebKit/537.36 (KHTML, like Gecko) '
                                                         'Chrome/85.0.4183.102 Safari/537.36'})


# time.sleep(2)

# Accepts URL and client_name as the parameter
# and returns the client address
def get_client_data(site_url, client_name):
    # Send GET request to site
    browser.get(site_url)

    # Search for input tag
    search_client = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="content"]/div/div[1]/div/div/div[4]/div[1]/div/div/div[1]/input')))
    # Send the data to the input tag and return/enter/search
    search_client.send_keys(client_name + Keys.RETURN)

    # browser.quit()
    # Select/Click the first result
    browser.execute_script("window.scrollTo(100,document.body.scrollHeight);")
    client_page = WebDriverWait(browser, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="page"]/div[3]/div[1]/div/div/div[4]/div[2]/div[3]/div[2]/div[2]/table/tbody/tr[1]/td[1]/a[1]'))).click()

    # browser.execute_script("arguments[0].scrollIntoView(true);", client_page)

    # Scrape/Grab the address data
    client_addr = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="company_profile_snapshot"]/div[2]/div[2]/span/span')))

    # Return the address
    return browser.current_url, client_addr.text.replace(' See other locations', '')


def write_csv(client_list, client_urls, client_addrs):
    # Create a client_data dictionary to store each clients name, url and the address
    client_data = {'Client': client_list,
                   "DnB Url": client_urls, 'DnB Address': client_addrs}

    # Create a dataframe of size n(no. of clients) * m (includes name, url, address)
    df = pd.DataFrame(client_data)

    # Write to csv
    df.to_csv('Client_Data.csv')


def main():
    client_list = ['Artemys Inc', 'BioComposites Ltd', 'Biofactura',
                   'Chimagen Biosciences Ltd']
    client_urls = []
    client_addrs = []
    site_url = 'https://www.dnb.com/business-directory.html#CompanyProfilesPageNumber=1&ContactProfilesPageNumber=1&DAndBMarketplacePageNumber=1&IndustryPageNumber=1&SiteContentPageNumber=1&tab=Company%20Profiles'

    # Loop through all the clients
    for client in client_list:
        # print(client)
        # Call get_client_data function and get the url and the address
        client_url, client_addr = get_client_data(site_url, client)
        # store the received data into the corresponding list
        client_urls.append(client_url)
        client_addrs.append(client_addr)

    # print(client_addrs)
    write_csv(client_list, client_urls, client_addrs)

    browser.quit()


if __name__ == '__main__':
    main()
