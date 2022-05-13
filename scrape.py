import os
import time
from webbrowser import Chrome
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException
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


# Accepts URL and client_name as the parameter
# and returns the client address
def get_dnb_data(site_url, client_name):
    # Send GET request to site
    browser.get(site_url)

    # Search for input tag
    search_client = WebDriverWait(browser, 20).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="page"]/div[2]/div/div[1]/div/div/div[4]/div[1]/div/div/div[1]/input')))

    # Send the data to the input tag and return/enter/search
    search_client.send_keys(client_name + Keys.RETURN)

    # Select/Click the first result
    # browser.execute_script("window.scrollTo(100,document.body.scrollHeight);")
    try:
        client_page = WebDriverWait(browser, 20).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="page"]/div[3]/div[1]/div/div/div[4]/div[2]/div[3]/div[2]/div[2]/table/tbody/tr[1]/td[1]/a[1]'))).click()

        # Scrape/Grab the client name
        client_name = WebDriverWait(browser, 20).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="page"]/div[3]/div/div/div[4]/div/div[1]/div/div[2]/div/div[1]/div[1]/span/span')))

        # Scrape/Grab the address data
        client_addr = WebDriverWait(browser, 20).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="company_profile_snapshot"]/div[2]/div[2]/span/span')))

        # Return the name, url and the address
        return client_name.text, browser.current_url, client_addr.text.replace(' See other locations', '')
    except (NoSuchElementException, TimeoutException):
        return client_name, None, None


def get_clientSite_data(client_name, client_search_addr):
    browser.get('https://google.com')
    search_term = client_name + ' ' + client_search_addr + ' address'

    # Search for input tag
    search_client = WebDriverWait(browser, 20).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input')))

    # Send the data to the input tag and return/enter/search
    search_client.send_keys(search_term + Keys.RETURN)

    xpaths = ['//*[@id="rso"]/div[1]/div/div/div/div/div[1]/div/div[1]', '//*[@id="kp-wp-tab-overview"]/div[1]/div/div/div/div/div/div[4]/div/div/div/span[2]',
              '//*[@id="kp-wp-tab-overview"]/div[1]/div/div/div/div/div/div[6]/div/div/div/span[2]']

    site_addr = ''
    for xpath in xpaths:
        try:
            site_addr = WebDriverWait(browser, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, xpath))).text
        except (NoSuchElementException, TimeoutException):
            pass

    search_term = search_term.replace(
        client_search_addr + ' address', 'contact page')

    search_client = WebDriverWait(browser, 20).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="tsf"]/div[1]/div[1]/div[2]/div/div[2]/input')))

    # Send the data to the input tag and return/enter/search
    search_client.clear()
    search_client.send_keys(search_term + Keys.RETURN)

    browser.find_element_by_tag_name('h3').click()
    time.sleep(2)

    return browser.current_url, site_addr


def write_csv(client_list, dnb_urls, dnb_addrs, client_urls, client_addrs):
    # Create a client_data dictionary to store each clients name, url and the address
    client_data = {'Client': client_list,
                   "DnB Url": dnb_urls, 'DnB Address': dnb_addrs, 'Client Url': client_urls, 'Client Address': client_addrs}

    # Create a dataframe of size n(no. of clients) * m (includes name, url, address)
    df = pd.DataFrame(client_data)

    # Write to csv
    df.to_csv('Client_Data.csv')


def main():
    client_search = ['Artemys Inc', 'BioComposites Ltd',
                     'Biofactura', 'Chimagen Biosciences Ltd', 'Chinook Therapeutics Us', 'CytomX Therapeutics Inc', 'Baxalta US Inc', 'Janssen Pharmaceutica NV', 'Emmes Biopharma Global s.r.o']
    client_search_addr = ['CA', 'ST5 5NL', '21701', '610000', 'Washington 98109-5311',
                          'California 94080-1840', 'Massachusetts 02421-2101', '2340 Belgium', '11000']
    client_list = []
    dnb_urls = []
    dnb_addrs = []
    client_urls = []
    client_addrs = []

    dnb_url = 'https://www.dnb.com/business-directory.html#CompanyProfilesPageNumber=1&ContactProfilesPageNumber=1&DAndBMarketplacePageNumber=1&IndustryPageNumber=1&SiteContentPageNumber=1&tab=Company%20Profiles'

    # Loop through all the clients
    for client_idx in range(len(client_search)):
        # Call get_client_data function and get the url and the address
        client_name, client_url, client_addr = get_dnb_data(
            dnb_url, client_search[client_idx])

        # store the received data into the corresponding list
        client_list.append(client_name)
        dnb_urls.append(client_url)
        dnb_addrs.append(client_addr)

        client_url, client_addr = get_clientSite_data(
            client_name, client_search_addr[client_idx])

        client_urls.append(client_url)
        client_addrs.append(client_addr)

    write_csv(client_list, dnb_urls, dnb_addrs, client_urls, client_addrs)
    # print(get_clientSite_data('BIOCOMPOSITES (UK) LIMITED'))
    browser.quit()


if __name__ == '__main__':
    main()
