import os
import time
from webbrowser import Chrome
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
import pandas as pd
import pypostalwin
from collections import ChainMap
from rapidfuzz import fuzz
from rapidfuzz.distance import Levenshtein
import pycountry
import unicodedata

# Options to bypass bot detection
options = webdriver.ChromeOptions()
options.add_argument("start-maximized")
# options.add_argument("--headless")
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_argument('--disable-site-isolation-trials')
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
# Change the pageLoadStrategy to eager to make page interactive/ Don't wait for whole page to load
caps = DesiredCapabilities().CHROME
caps["pageLoadStrategy"] = "eager"


# Import/Run Chrome Driver
# script_dir = os.path.dirname(os.path.realpath(__file__))
# executable_path = os.path.join(script_dir, "chromedriver.exe")
# browser = webdriver.Chrome(desired_capabilities=caps, options=options,
#                            executable_path=executable_path)

service = ChromeDriverManager().install()
with open(service, "rb") as input_file:
    content = input_file.read()
    content = content.replace(b"cdc_", b"dog_")

with open(service, "wb") as output_file:
    output_file.write(content)

browser = webdriver.Chrome(
    service=Service(service), desired_capabilities=caps, options=options)

browser.execute_script(
    "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
browser.execute_cdp_cmd('Network.setUserAgentOverride', {
                        "userAgent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36'})


# browser.set_page_load_timeout(25)

# Initialize Address Parser
parser = pypostalwin.AddressParser()

# Removes special character from address before passing the address to the parser


def remove_accents(input_str):
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return u"".join([c for c in nfkd_form if not unicodedata.combining(c)])


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
                (By.XPATH, "(//*[@class='z20f9dcf9d844d2bf_tableCompanyNameLink _BH_iAJa6u2y3ordNzHR'])[1]"))).click()

        # Scrape/Grab the client name
        client_name = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@class="company-profile-header-title"]'))).get_attribute("textContent")

        # Scrape/Grab the address data
        client_addr = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@name="company_address"]/span'))).get_attribute("textContent")

        # Return the name, url and the address
        return client_name, browser.current_url, client_addr.replace(' See other locations', '').strip()
    except (NoSuchElementException, TimeoutException):
        return client_name, None, None


def get_clientSite_data(client_name, city_postcode):
    # Send GET request to google/ Go to google.com
    browser.get('https://google.com')

    # Set the search term to client name followed by city and postcode + address string to the input tag
    search_term = client_name + ' ' + city_postcode + ' address'

    # Search for input tag
    search_client = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input')))

    # Send the data to the input tag and return/enter/search
    search_client.send_keys(search_term + Keys.RETURN)

    site_addr = ''

    try:
        site_addr = WebDriverWait(browser, 5).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="rso"]/div[1]/div/div/div/div/div[1]/div/div[1]'))).get_attribute("textContent")
    except (NoSuchElementException, TimeoutException):
        try:
            WebDriverWait(browser, 5).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '(//*[@class="OSrXXb"])[1]'))).click()

            site_addr = WebDriverWait(browser, 5).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[@class='LrzXr']"))).get_attribute("textContent")
            # '/html/body/div[6]/div/div[9]/div[2]/div/div[2]/async-local-kp/div/div/div[1]/div/div/block-component/div/div[1]/div/div/div/div[1]/div/div/div[4]/div/div/span[2]'))).get_attribute("textContent")

            browser.back()
        except:
            pass

    # Replace the address string with contact page string/ Change the search term
    search_term = search_term.replace(
        city_postcode + ' address', 'contact page')

    # Search for the input tag
    search_client = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="tsf"]/div[1]/div[1]/div[2]/div/div[2]/input')))

    # Clear the input text area
    search_client.clear()

    # Send the data to the input tag and return/enter/search
    search_client.send_keys(search_term + Keys.RETURN)

    # Go to the first site
    try:
        browser.find_element_by_tag_name('h3').click()
    except:
        return None, site_addr

    return browser.current_url, site_addr


def get_sec_data(site_url, client_name):
    # Send GET request to site
    browser.get(site_url)

    # Close the AD Popup
    try:
        WebDriverWait(browser, 1).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="acsMainInvite"]/div/a[1]'))).click()
    except:
        pass

    # Search for input tag
    try:
        search_client = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="company"]')))

        time.sleep(1)

        # Send the data to the input tag and return/enter/search
        search_client.send_keys(client_name)
        time.sleep(1)
        search_client.send_keys(Keys.ENTER)
    except:
        return None, None

    try:
        # Expand the dropdown
        WebDriverWait(browser, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="entityInformationHeader"]'))).click()

        time.sleep(1)

        client_addr = WebDriverWait(browser, 5).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, '#businessAddress'))).get_attribute("textContent")

    except (NoSuchElementException, TimeoutException):
        try:
            client_addr = WebDriverWait(browser, 5).until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[4]/div[1]/div[2]'))).get_attribute("textContent")

            client_addr = client_addr.replace('Business Address', '').strip()

            return browser.current_url, client_addr
        except (NoSuchElementException, TimeoutException):
            return None, None

    return browser.current_url, client_addr


def desired_addr_format(addr):
    # Converts list of n dictionaries to list of dictionary
    addr = dict(ChainMap(*addr))

    # Get the 2-digit and 3-digit country code
    country = pycountry.countries.get(name=addr.get('country', ''))
    addr_str = ''

    # Store the segregated address in a dictionary
    addr_dict = {
        'Address_Line 2': f"{addr.get('unit', '')}".strip(),
        'Address Line 3': f"{addr.get('house_number', '')} {addr.get('house', '')} {addr.get('road', '')}".strip(),
        'City': f"{addr.get('city', '')}".strip(),
        'State': f"{addr.get('state', '')}".strip(),
        'PostalCode': f"{addr.get('postcode', '')}".strip(),
        'CountryCode_2': country.alpha_2 if country else '',
        'CountryCode_3': country.alpha_3 if country else ''
    }

    # Convert dictionary to string to store the data in one cell
    for key, val in addr_dict.items():
        addr_str += f"{key.upper()}: {val.upper()} \n"

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
        # Compare only the first 5 digits of the postcode for accurate comparison
        if addr_key == 'postcode':
            sample_addr_dict[addr_key] = sample_addr_dict[addr_key].split(
                '-')[0] if addr_key in sample_addr_dict else ''
            dnb_addr_dict[addr_key] = dnb_addr_dict[addr_key].split(
                '-')[0] if addr_key in dnb_addr_dict else ''
            client_addr_dict[addr_key] = client_addr_dict[addr_key].split(
                '-')[0] if addr_key in client_addr_dict else ''
            sec_addr_dict[addr_key] = sec_addr_dict[addr_key].split(
                '-')[0] if addr_key in sec_addr_dict else ''

        # Levenshtein distance algo to compare two strings
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

    # store the similarity percent along with the address in a dictionary
    similarity_factor_dict = {dnb_levenshtein_factor: parsed_dnb_addr,
                              client_levenshtein_factor: parsed_client_addr, sec_levenshtein_factor: parsed_sec_addr}

    return similarity_factor_dict


def parse_addr(client_search_addr, dnb_addrs, client_addrs, sec_addrs):
    addr_list_str = []

    # Send the unparsed address to the parser to segregate the address into house, state, city, postalcode, country, etc
    for addr_idx in range(len(client_search_addr)):
        print('Before Comparison')

        print('Sample')
        parsed_sample_addr = parser.runParser(client_search_addr[addr_idx])
        print(parsed_sample_addr, '\n')

        print('Fetched from DNB')
        # print(dnb_addrs[addr_idx])
        if dnb_addrs[addr_idx]:
            # try:
            #     parsed_dnb_addr = parser.runParser(dnb_addrs[addr_idx])
            # except:
            # Pass the address to the remove accent function to remove special characters
            addr = remove_accents(dnb_addrs[addr_idx]).encode(
                "ascii", "ignore").decode()
            parsed_dnb_addr = parser.runParser(addr)
        else:
            parsed_dnb_addr = {}
        print(parsed_dnb_addr, '\n')

        print('Fetched from Google')
        if client_addrs[addr_idx]:
            # try:
            #     parsed_client_addr = parser.runParser(client_addrs[addr_idx])
            # except:
            addr = remove_accents(
                client_addrs[addr_idx]).encode("ascii", "ignore").decode()

            parsed_client_addr = parser.runParser(addr)
        else:
            parsed_client_addr = {}
        print(parsed_client_addr, '\n')

        print('Fetched from sec.gov')
        if sec_addrs[addr_idx]:
            # try:
            #     parsed_sec_addr = parser.runParser(sec_addrs[addr_idx])
            # except:
            addr = remove_accents(sec_addrs[addr_idx]).encode(
                "ascii", "ignore").decode()
            parsed_sec_addr = parser.runParser(addr)
        else:
            parsed_sec_addr = {}
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
    df.to_csv('Client_Data_3.csv')


def main():
    # Read Input/Sample data from excel sheet into dataframe
    # client_df = pd.read_excel('Client.xlsx')
    client_data_sheet = input('Enter excel file name (with extension): ')
    client_df = pd.read_excel(client_data_sheet)

    # Fetch the data from columns
    client_search_list = client_df['Client'].tolist()
    client_search_addr = client_df['Address'].tolist()
    city_postcode = []

    # get the city and postcode for seaching the address on google
    for addr_idx in range(len(client_search_addr)):
        client_search_addr[addr_idx] = client_search_addr[addr_idx].replace(
            '\n', ', ')
        temp = dict(ChainMap(*parser.runParser(client_search_addr[addr_idx])))
        city_postcode.append(temp.get('city', '') +
                             ' ' + temp.get('postcode', ''))

    client_list = []
    dnb_urls = []
    dnb_addrs = []
    client_urls = []
    client_addrs = []
    sec_urls = []
    sec_addrs = []

    dnb_url = 'https://www.dnb.com/business-directory.html'

    sec_url = 'https://www.sec.gov/edgar/searchedgar/companysearch.html'

    # Improves searching result
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

        # Get data from google
        client_url, client_addr = get_clientSite_data(
            client_search_list[client_idx], city_postcode[client_idx])

        client_urls.append(client_url)
        client_addrs.append(client_addr)

        # Get data from sec.gov
        client_url, client_addr = get_sec_data(
            sec_url, client_search_list[client_idx])

        sec_urls.append(client_url)
        sec_addrs.append(client_addr)

    browser.quit()

    # Send the fetched address into the parser function
    addr_list_str = parse_addr(
        client_search_addr, dnb_addrs, client_addrs, sec_addrs)

    # Save data to csv
    write_csv(client_list, dnb_urls, dnb_addrs, client_urls,
              client_addrs, sec_urls, sec_addrs, addr_list_str)


if __name__ == '__main__':
    start_time = time.time()
    main()
    parser.terminateParser()
    print("--- %s seconds ---" % (time.time() - start_time))
