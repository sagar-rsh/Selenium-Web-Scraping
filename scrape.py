import os
import time
from webbrowser import Chrome
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import pandas as pd
import pypostalwin
from collections import ChainMap
from rapidfuzz import fuzz
from rapidfuzz.distance import Levenshtein
import pycountry

# Options to bypass bot detection
options = webdriver.ChromeOptions()
options.add_argument("start-maximized")
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# Change the pageLoadStrategy to eager to make page interactive/ Don't wait for whole page to load
caps = DesiredCapabilities().CHROME
caps["pageLoadStrategy"] = "eager"


# Import/Run Chrome Driver
script_dir = os.path.dirname(os.path.realpath(__file__))
executable_path = os.path.join(script_dir, "chromedriver.exe")
browser = webdriver.Chrome(desired_capabilities=caps, options=options,
                           executable_path=executable_path)
browser.execute_script(
    "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
browser.execute_cdp_cmd('Network.setUserAgentOverride', {"userAgent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                                                         'AppleWebKit/537.36 (KHTML, like Gecko) '
                                                         'Chrome/85.0.4183.102 Safari/537.36'})
# browser.set_page_load_timeout(25)


# Accepts URL and client_name as the parameter
# and returns the client address
def get_dnb_data(site_url, client_name):
    # Send GET request to site
    browser.get(site_url)

    # Search for input tag
    search_client = WebDriverWait(browser, 20).until(
        EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div/div[4]/div[1]/div/div/div[1]/input')))

    # Send the data to the input tag and return/enter/search
    search_client.send_keys(client_name + Keys.RETURN)

    # Select/Click the first result
    try:
        client_page = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="page"]/div[3]/div[1]/div/div/div[4]/div[2]/div[3]/div[2]/div[2]/table/tbody/tr[1]/td[1]/a[1]'))).click()

        # Scrape/Grab the client name
        client_name = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[2]/div[3]/div/div/div[5]/div/div/div/div[2]/div/div[1]/div[1]/span/span'))).get_attribute("textContent")

        # Scrape/Grab the address data
        client_addr = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="company_profile_snapshot"]/div[2]/div[2]/span/span'))).text

        # Return the name, url and the address
        return client_name, browser.current_url, client_addr.replace(' See other locations', '')
    except (NoSuchElementException, TimeoutException):
        return client_name, None, None


def get_clientSite_data(client_name, client_city):
    browser.get('https://google.com')
    search_term = client_name + ' ' + client_city + ' address'

    # Search for input tag
    search_client = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input')))

    # Send the data to the input tag and return/enter/search
    search_client.send_keys(search_term + Keys.RETURN)

    xpaths = ['//*[@id="rso"]/div[1]/div/div/div/div/div[1]/div/div[1]', '//*[@id="kp-wp-tab-overview"]/div[1]/div/div/div/div/div/div[4]/div/div/div/span[2]',
              '//*[@id="kp-wp-tab-overview"]/div[1]/div/div/div/div/div/div[6]/div/div/div/span[2]']

    site_addr = ''
    for xpath in xpaths:
        try:
            site_addr = WebDriverWait(browser, 5).until(
                EC.presence_of_element_located(
                    (By.XPATH, xpath))).get_attribute("textContent")
            break
        except (NoSuchElementException, TimeoutException):
            pass

    search_term = search_term.replace(
        client_city + ' address', 'contact page')

    search_client = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="tsf"]/div[1]/div[1]/div[2]/div/div[2]/input')))

    search_client.clear()
    # Send the data to the input tag and return/enter/search
    search_client.send_keys(search_term + Keys.RETURN)

    browser.find_element_by_tag_name('h3').click()

    # time.sleep(2)

    return browser.current_url, site_addr


def get_sec_data(site_url, client_name):
    # Send GET request to site
    browser.get(site_url)

    try:
        WebDriverWait(browser, 1).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="acsMainInvite"]/div/a[1]'))).click()
    except:
        pass

    # Search for input tag
    search_client = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="company"]')))

    time.sleep(2)
    # Send the data to the input tag and return/enter/search
    search_client.send_keys(client_name + Keys.RETURN)

    try:
        WebDriverWait(browser, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="entityInformationHeader"]'))).click()

        time.sleep(1)

        client_addr = WebDriverWait(browser, 5).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, '#businessAddress'))).get_attribute("textContent")

    except (NoSuchElementException, TimeoutException):
        return None, None

    return browser.current_url, client_addr


def desired_addr_format(addr):
    addr = dict(ChainMap(*addr))
    country = pycountry.countries.get(name=addr.get('country', ''))
    addr_str = ''
    addr_list_str = []
    addr_dict = {
        'Address_Line 2': f"{addr.get('unit', '')}".strip(),
        'Address Line 3': f"{addr.get('house_number', '')} {addr.get('house', '')} {addr.get('road', '')}".strip(),
        'City': f"{addr.get('city', '')}".strip(),
        'State': f"{addr.get('state', '')}".strip(),
        'PostalCode': f"{addr.get('postcode', '')}".strip(),
        'CountryCode_2': country.alpha_2 if country else '',
        'CountryCode_3': country.alpha_3 if country else ''
    }

    for key, val in addr_dict.items():
        addr_str += f"{key}: {val} \n"
    # print(addr_dict)
    # print(addr_str)
    return addr_str


def compare_addr(parsed_sample_addr, parsed_dnb_addr, parsed_client_addr, parsed_sec_addr):
    dnb_levenshtein_factor = 0
    client_levenshtein_factor = 0
    sec_levenshtein_factor = 0

    # Encapsulate many dictionaries to one dictionary
    sample_addr_dict = dict(ChainMap(*parsed_sample_addr))
    dnb_addr_dict = dict(ChainMap(*parsed_dnb_addr))
    client_addr_dict = dict(ChainMap(*parsed_client_addr))
    sec_addr_dict = dict(ChainMap(*parsed_sec_addr))

    for addr_key in sample_addr_dict:
        if addr_key == 'postcode':
            sample_addr_dict[addr_key] = sample_addr_dict[addr_key].split(
                '-')[0] if addr_key in sample_addr_dict else ''
            dnb_addr_dict[addr_key] = dnb_addr_dict[addr_key].split(
                '-')[0] if addr_key in dnb_addr_dict else ''
            client_addr_dict[addr_key] = client_addr_dict[addr_key].split(
                '-')[0] if addr_key in client_addr_dict else ''
            sec_addr_dict[addr_key] = sec_addr_dict[addr_key].split(
                '-')[0] if addr_key in sec_addr_dict else ''

        try:
            levenshtein_percent = Levenshtein.normalized_similarity(
                sample_addr_dict[addr_key], dnb_addr_dict[addr_key])
            dnb_levenshtein_factor += (levenshtein_percent /
                                       len(parsed_sample_addr))

            levenshtein_percent = Levenshtein.normalized_similarity(
                sample_addr_dict[addr_key], client_addr_dict[addr_key])
            client_levenshtein_factor += (levenshtein_percent /
                                          len(parsed_sample_addr))

            levenshtein_percent = Levenshtein.normalized_similarity(
                sample_addr_dict[addr_key], sec_addr_dict[addr_key])
            sec_levenshtein_factor += (levenshtein_percent /
                                       len(parsed_sample_addr))
        except KeyError:
            pass

    similarity_factor_dict = {dnb_levenshtein_factor: parsed_dnb_addr,
                              client_levenshtein_factor: parsed_client_addr, sec_levenshtein_factor: parsed_sec_addr}

    return similarity_factor_dict


def parse_addr(client_search_addr, dnb_addrs, client_addrs, sec_addrs):
    parser = pypostalwin.AddressParser()
    addr_list_str = []

    for addr_idx in range(len(client_search_addr)):
        print('Before Comparison')

        print('Sample')
        parsed_sample_addr = parser.runParser(client_search_addr[addr_idx])
        print(parsed_sample_addr, '\n')

        print('Fetched from DNB')
        parsed_dnb_addr = parser.runParser(
            dnb_addrs[addr_idx]) if dnb_addrs[addr_idx] else ''
        print(parsed_dnb_addr, '\n')

        print('Fetched from Google')
        parsed_client_addr = parser.runParser(
            client_addrs[addr_idx]) if client_addrs[addr_idx] else ''
        print(parsed_client_addr, '\n')

        print('Fetched from sec.gov')
        parsed_sec_addr = parser.runParser(
            sec_addrs[addr_idx]) if sec_addrs[addr_idx] else ''
        print(parsed_sec_addr, '\n')

        print('After Comparison')
        similarity_factor_dict = compare_addr(parsed_sample_addr, parsed_dnb_addr,
                                              parsed_client_addr, parsed_sec_addr)
        print('\n \n')

        addr_str = desired_addr_format(
            similarity_factor_dict[max(similarity_factor_dict)])

        addr_list_str.append(addr_str)

    return addr_list_str


def write_csv(client_list, dnb_urls, dnb_addrs, client_urls, client_addrs, sec_urls, sec_addrs, addr_list_str):
    # Create a client_data dictionary to store each clients name, url and the address
    client_data = {'Client': client_list,
                   'DnB URL': dnb_urls,
                   'DnB Address': dnb_addrs,
                   'Client URL': client_urls,
                   'Client Address': client_addrs,
                   'SEC URL': sec_urls,
                   'SEC Address': sec_addrs,
                   'Address Lines': addr_list_str}

    # Create a dataframe of size n(no. of clients) * m (includes name, url, address)
    df = pd.DataFrame(client_data)

    # Write to csv
    df.to_csv('Client_Data.csv')


def main():
    client_search_list = ['Artemys Inc', 'BioComposites Ltd',
                          'Biofactura', 'Chimagen Biosciences Ltd', 'Chinook Therapeutics US Inc', 'CytomX Therapeutics Inc', 'Baxalta US Inc', 'Janssen Pharmaceutica NV']
    # , 'Emmes Biopharma Global s.r.o'
    client_search_addr = ['1933 Davis St Suite 244, San Leandro, CA 94577, United States', 'Keele Science Park, Keele University, Keele,ST5 5NL,United Kingdom',
                          '8435 Progress Dr, Frederick, Maryland 21701,United States', 'No 5 Keyuan S Rd, Bldg 1 Fl 9, Chengdu, 610000,China', 'Ste 100,1600 Fairview Ave E,Seattle, Washington 98109-5311,United States', '151 Oyster Point Blvd,Ste 400,South San Francisco, California 94080-1840,United States', '650 Kendall DrCambridge, Massachusetts 02421-2101,United States', 'Turnhoutseweg 30,Beerse, 2340,Belgium']
    # , 'V Jame 699/1,Prague 1, 11000,Czech Republic'

    client_city = ['CA', 'ST5 5NL', '21701', '610000', 'Washington',
                   'California 94080-1840', 'Massachusetts 02421-2101', '2340 Belgium', '11000']
    client_list = []
    dnb_urls = []
    dnb_addrs = []
    client_urls = []
    client_addrs = []
    sec_urls = []
    sec_addrs = []

    dnb_url = 'https://www.dnb.com/business-directory.html#CompanyProfilesPageNumber=1&ContactProfilesPageNumber=1&DAndBMarketplacePageNumber=1&IndustryPageNumber=1&SiteContentPageNumber=1&tab=Company%20Profiles'

    sec_url = 'https://www.sec.gov/edgar/searchedgar/companysearch.html'

    for client_idx in range(len(client_search_list)):
        client_search_list[client_idx] = client_search_list[client_idx].replace(
            ' US ', ' U.S. ')

    # Loop through all the clients
    for client_idx in range(len(client_search_list)):
        # Call get_client_data function and get the url and the address
        client_name, client_url, client_addr = get_dnb_data(
            dnb_url, client_search_list[client_idx])

        # store the received data into the corresponding list
        client_list.append(client_name)
        dnb_urls.append(client_url)
        dnb_addrs.append(client_addr)

        client_url, client_addr = get_clientSite_data(
            client_search_list[client_idx], client_city[client_idx])

        client_urls.append(client_url)
        client_addrs.append(client_addr)

        client_url, client_addr = get_sec_data(
            sec_url, client_search_list[client_idx])

        sec_urls.append(client_url)
        sec_addrs.append(client_addr)

    # print(get_clientSite_data('BIOCOMPOSITES (UK) LIMITED'))

    browser.quit()

    addr_list_str = parse_addr(
        client_search_addr, dnb_addrs, client_addrs, sec_addrs)
    write_csv(client_list, dnb_urls, dnb_addrs, client_urls,
              client_addrs, sec_urls, sec_addrs, addr_list_str)


if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
