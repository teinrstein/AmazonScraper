from typing import Dict, List, Optional
import requests
from bs4 import BeautifulSoup
import pandas as pd
# from pandas import ExcelWriter
import time
import re
import xlsxwriter


BASE_URL = "https://www.amazon.co.uk"
amzn_IDs = {
    "A3P5ROKL5A1OLE": "GB",
    "ATVPDKIKX0DER": "US",
    "A2KVF7QXNCLV8H": "US",
    "A3DWYIK6Y9EEQB": "CA",
    "A3JWKAKR8XB7XF": "DE",
    "A30DC7701CXIBH": "DE",
    "A1X6FK5RDHNB96": "FR",
    "AN1VRQENFRJN5": "JP"
}

HEADERS = {
    "Host": "www.amazon.co.uk",
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_5) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/12.1.1 Safari/605.1.15",
     # "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:91.0) Gecko/20100101 Firefox/91.0",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-GB,en;q=0.5",
    # "Accept-Encoding": "gzip, deflate, br",
    # "Referer": "https://www.amazon.co.uk/s?k=aaa&page=1",
    # "DNT": "1",
    "Connection": "keep-alive",
    "Cookie": 'session-id=262-1842978-9544710; session-id-time=2082787201l; i18n-prefs=GBP; csm-hit=tb:s-HJDT1ZFW5DTCJC1QX8BD|1629919883593&t:1629919885753&adb:adblk_no; ubid-acbuk=259-7037632-9005606; session-token="yaFVW1jB/hz1KcfO9w5MZS8LtJM2I+aB/6OyWYIY7XU7eTjjivKhxbBt72YEzp0+/6l8Agh5NOC9SaP/KYhJiexVaXmDPPktvhz4WdBglU6hurOak2/gFoUHIZ4L8c0FoJcdf8QOOJ9ZedkNNXo+2MlHZgM5jMBiTlozd42vje+brOkYkQpf5UrA5iY+QXratrqdZSjVttuKbpkubV6ycERTs8ibbTrFky2rOCeRKu4e8EGMEKjV+PrgofGW6Kwh9nhh5vrmG/A="',
    # "Sec-Fetch-Dest": "document",
    # "Sec-Fetch-Mode": "navigate",
    # "Sec-Fetch-Site": "same-origin",
    # "Sec-Fetch-User": "?1",
    # "Sec-GPC": "1",
    # "Cache-Control": "max-age=0"
}


def search_url(term: str, page: int = 1) -> str:
    return f"{BASE_URL}/s?k={term}&page={page}"


def get_request(url: str) -> Optional[bytes]:
#    print("request.get: ", url, " ", end = " ")
    try:
        r = requests.get(url, headers=HEADERS)
    except requests.exceptions.ConnectionError:
        print("bounced")
        time.sleep(60)
        # print("request.get: ", url, " ", end=" ")
        try:
            r = requests.get(url, headers=HEADERS)
        except requests.exceptions.ConnectionError:
            print("bounced AGAIN")
            time.sleep(180)
            # print("request.get: ", url, " ", end=" ")
            r = requests.get(url, headers=HEADERS)
    print (r.status_code, end =" ")
    time.sleep(5)
    if r.status_code != 200:
        if r.status_code > 500:
            print("Probably rate-restricted, code ", r.status_code)
        else:
            print("Got error ", r.status_code)
        return None
    else:
        return r.content


def clean_soup(soup: BeautifulSoup) -> BeautifulSoup:
    widget = soup.find(class_="widgetId=loom-desktop-top-slot_hsa-id")
    if widget is not None:
        widget.decompose()
    return soup


def get_page_items(term: str, page: int):
    r = get_request(search_url(term, page))
    if r is None:
        return []
    soup = clean_soup(BeautifulSoup(r, "lxml"))
    return [x for x in soup.find_all("div", class_="s-result-item") if
            "AdHolder" not in x["class"] and x["data-asin"] != '']

def get_merchant_ID(soup: BeautifulSoup):
    m_ID = None
    buybox = soup.find("div", id="buybox")
    if buybox is None:
        print("++no buybox on page, aborting++", end=" ")
        return m_ID

    buying_choices = buybox.find_all("input", id="merchantID")

    if len(buying_choices) == 0 or buying_choices[0]["value"] == '':
        print("++got buybox, but no merchantID++", end=" ")
        return m_ID
    m_ID = buying_choices[0]["value"]
    return m_ID

def get_fbx(soup: BeautifulSoup):

    buybox = soup.find("div", id="buybox")

    if buybox is None:
#        print("++FBX no buybox on page, aborting++", end=" ")
        return None

    sbfb = {
        'Dispatches from': None,
        'Sold by': None
    }
#    if buybox.find("div", id="merchant-info") is not None:
#        merchant_info = buybox.find("div", id="merchant-info")
#        s = merchant_info.find_all("div", class_="a-row")

    if buybox.find("div", id = "sfsb_accordion_head") is not None:
        merchant_info = buybox.find("div", id = "sfsb_accordion_head")
        sb_arr = [row.get_text().replace('\n','') for row in merchant_info.find_all("div", class_="a-row")]
        sbfb['Dispatches from'] = sb_arr[0].split('Dispatches from:')[1]
        sbfb['Sold by'] = sb_arr[1].split('Sold by:')[1]

    elif buybox.find("div", id = "tabular_feature_div") is not None:
        merchant_info = buybox.find("div", id = "tabular_feature_div")
        for t_row in merchant_info.find_all("tr"):
            t_row_str = t_row.get_text().replace('\n','')
            if t_row_str.startswith('Dispatches from'):
                sbfb['Dispatches from'] = t_row_str.split('Dispatches from')[1]
            elif t_row_str.startswith('Sold by'):
                sbfb['Sold by'] = t_row_str.split('Sold by')[1]
            else:
                print("Unrecognised string in buybox: ", t_row_str)
    else:
        print("Unrecognised format in buybox = ", buybox)

    if sbfb['Sold by'].startswith("Amazon"):
        return 'AMZ'
    elif sbfb['Dispatches from'].startswith("Amazon"):
        return 'FBA'
    else:
        return 'FBM'

def get_num_products(in_soup: BeautifulSoup):
    num_products = None
    linkbox = in_soup.find("div", id="storefront-link")
    if linkbox is not None:
        url = BASE_URL + linkbox.find("a", href = True)['href']

    r = get_request(url)
    soup =  BeautifulSoup(r, "lxml")
    widget = soup.find("span", cel_widget_id="UPPER-RESULT_INFO_BAR-0")
    results_str = widget.find("div", class_="sg-col-inner").get_text().replace('\n','')
    num_pr_str = re.search(r"(\d+) result", results_str)
    if num_pr_str:
        num_products = num_pr_str.group(1)
    return num_products

def get_details(url: str):
    p_details = {
        'P_Reviews_No': None,
        'P_Reviews_Rank': None,
        'First_Available': None,
        's_href': None,
        'merchant_ID': None,
        'Seller': None,
        'No. Products': None,
        'S_Star_Rating': None,
        'S_Reviews_No': None,
        'S_Positive_Reviews': None,
        'Company': None,
        'Registration_Number': None,
        'Country': None
    }

    s_details = {
        's_href': None,
        'merchant_ID': None,
        'Seller': None,
        'No. Products': None,
        'S_Star_Rating': None,
        'S_Reviews_No': None,
        'S_Positive_Reviews': None,
        'Company': None,
        'Registration_Number': None,
        'Country': None
    }
    merchant_ID = None
    soup = None
    # ################ OPEN PRODUCT PAGE #####################
    r_seller_product = get_request(url)
    if r_seller_product is None:
        print("++error calling seller product++")
        return []

    soup = BeautifulSoup(r_seller_product, "lxml")

    if get_merchant_ID(soup) is None:
        url = url + "&psc=1"
        r_seller_product = get_request(url) # another request, now with &psc = 1
        if r_seller_product is None:        # still no seller linked to product, abort mission
            print("++ ERROR calling seller product, terminate get_details() ++")
            return []
        soup = BeautifulSoup(r_seller_product, "lxml")

    merchant_ID = get_merchant_ID(soup)


    # p_details["Title"] = soup.find("h1", id="title").get_text().strip()
    # price_box = soup.find("div", id="price")
    # if price_box is not None:
    #     if price_box.find("span", id ="priceblock_ourprice") is not None:
    #         price_str = soup.find("div", id="price").find("span", id ="priceblock_ourprice").get_text().replace('\n','').replace('£','')
    #         p_details["Price"] = float(price_str)
    #     elif price_box.find("span", id="priceblock_saleprice") is not None:
    #         price_str = soup.find("div", id="price").find("span", id="priceblock_saleprice").get_text().replace('\n','').replace('£', '')
    #         p_details["Price"] = float(price_str)
    # else:
    #     if soup.find("span", id="price_inside_buybox") is not None:
    #         price_string = soup.find("span", id="price_inside_buybox").get_text().strip()
    #         p_details["Price"] = float(price_string[1:].replace(',', ''))

    product_addinfo = soup.find("div", id="productDetails_feature_div")
    if product_addinfo is not None:
        product_addinfo_table = soup.find("div", id="productDetails_feature_div").find("table", id="productDetails_detailBullets_sections1")
        if product_addinfo_table is not None:
            for tr in product_addinfo_table.find_all("tr"):
                row_header = tr.find("th").get_text().strip()
                # if row_header == 'Date First Available':
                #     p_details["First_Available"]=tr.find("td").get_text().strip()
                if row_header == 'Customer Reviews':
                    p_details["P_Reviews_No"] = tr.find("span", id="acrCustomerReviewText").get_text().split('ratings', 1)[0].strip()
                    p_details["P_Reviews_Rank"] = tr.find("span", id="acrPopover")["title"].split("out", 1)[0].strip()
                # elif row_header == 'Best Sellers Rank':
                #     rankings = tr.find("td").get_text().strip().split('\n', 5)
                #     details["BSR"] = rankings[0].split('in', 1)[0].strip()

    if merchant_ID == '' or merchant_ID is None:  # merchant_ID not recovered despite &psc=1
        print("++merchant_ID not recovered++")
        return p_details, s_details

    if merchant_ID in amzn_IDs.keys():  # it is one of AMZN's units in diff. jurisdictions
        # details["fbx"] = "AMZ"
        p_details["Seller"] = "Amazon"
        p_details["Company"] = "Amazon" + " " + amzn_IDs[merchant_ID]
        p_details["Country"] = amzn_IDs[merchant_ID]
        return p_details, s_details

    # details["fbx"] = get_fbx(soup)

    if dfp["merchant_ID"].str.contains(merchant_ID).sum() > 0:  # Merchant already captured in previous runs
        dfrow = dfp.loc[dfp['merchant_ID'] == merchant_ID].head(n=1).to_dict('r')[0]
        for key, value in p_details.items():
            if key in ['s_href', 'merchant_ID', 'Seller', 'No. Products', 'S_Star_Rating', 'S_Reviews_No', 'S_Positive_Reviews', 'Company', 'Registration_Number', 'Country']:
                p_details[key] = dfrow[key]
        return p_details, s_details

    p_details["merchant_ID"] = merchant_ID
    s_details["merchant_ID"] = merchant_ID

    # ################ OPEN SELLER PAGE #####################
    soup = None
    seller_url = BASE_URL + "/sp?&seller=" + merchant_ID
    p_details["s_href"] = seller_url
    s_details["s_href"] = seller_url
    seller_page = get_request(seller_url)
    if seller_page is None:
        print("++ error calling seller url, terminating get_details) ++")
        return []

    soup = BeautifulSoup(seller_page, "lxml")
    seller_name_h1 = soup.find("h1", id="sellerName")
    if seller_name_h1 is None:
        print("++cant see seller name++")
        return p_details, s_details

    p_details["Seller"] = seller_name_h1.get_text().strip()
    s_details["Seller"] = seller_name_h1.get_text().strip()

    num_products = get_num_products(soup)
    p_details["No. Products"] = num_products
    s_details["No. Products"] = num_products

    seller_feedback_box = soup.find("div", id="seller-feedback-summary")
    if seller_feedback_box is not None:
        stars = seller_feedback_box.find("i", class_="a-icon a-icon-star a-star-5 feedback-detail-stars")
        if stars is not None:
            p_details["S_Star_Rating"] = stars.get_text().split(' out of ', 1)[0]
            s_details["S_Star_Rating"] = stars.get_text().split(' out of ', 1)[0]
        else:
            stars = seller_feedback_box.find("i", class_="a-icon a-icon-star a-star-4-5 feedback-detail-stars")
            if stars is not None:
                p_details["S_Star_Rating"] = stars.get_text().split(' out of ', 1)[0]
                s_details["S_Star_Rating"] = stars.get_text().split(' out of ', 1)[0]
        seller_rev_soup = seller_feedback_box.find("a", class_="a-link-normal feedback-detail-description")
        if seller_rev_soup is not None:
            seller_reviews = seller_rev_soup.get_text().split('% positive ',1)
            p_details["S_Positive_Reviews"] = seller_reviews[0]
            s_details["S_Positive_Reviews"] = seller_reviews[0]
            p_details["S_Reviews_No"] = seller_reviews[1].split(" (",1)[1].replace(' ratings)','')
            s_details["S_Reviews_No"] = seller_reviews[1].split(" (", 1)[1].replace(' ratings)', '')

    biz_details = soup.find("ul", class_="a-unordered-list a-nostyle a-vertical")
    if not biz_details:
        print("++cant see company details++", end=" ")
        return p_details, s_details

    all_rows =  [x.get_text("|") for x in biz_details.find_all("li")]
    if all_rows[0].startswith('Business Name:'):
        p_details["Company"] = all_rows[0][len('Business Name:|'):]
        s_details["Company"] = all_rows[0][len('Business Name:|'):]
        if len(all_rows)>=2:
            for i in range(1, len(all_rows)-1):
                if all_rows[i].startswith('Trade Register Number:'):
                    p_details["Registration_Number"] = all_rows[i][len('Trade Register Number:|'):]
                    s_details["Registration_Number"] = all_rows[i][len('Trade Register Number:|'):]
                elif all_rows[i].startswith('Business Address:'):
                    p_details["Country"] = all_rows[-1]
                    s_details["Country"] = all_rows[-1]
    else:
        print("++incomplete company record++", end=" ")
        p_details["Company"] = all_rows[0]
        s_details["Company"] = all_rows[0]   # most likely just VAT number of some company, but still

    return p_details, s_details


# ++++++++++++++++++++++++++ MAIN START  +++++++++++++++++++++++++
# counter = 1
# page = 1
# page_items = get_page_items(search_term, page)

Columns = [
    'Product Name',
    'Brand',
    'Price',
    'Mo. Sales',
    'D. Sales',
    'Mo. Revenue',
    'Date First Available',
    'Net',
    'Rating Number',
    'Rating',
    'Rank',
    'Seller type',
    'Sellers',
    'Category',
    'Tier',
    'ASIN',
    'Link',
    'P_Reviews_No',
    'P_Reviews_Rank',
    's_href',
    'merchant_ID',
    'Seller',
    'No. Products',
    'S_Star_Rating',
    'S_Reviews_No',
    'S_Positive_Reviews',
    'Company',
    'Registration_Number',
    'Country'
]
dfp = pd.DataFrame(columns=Columns)

Columns = [
    's_href',
    'merchant_ID',
    'Seller',
    'No. Products',
    'S_Star_Rating',
    'S_Reviews_No',
    'S_Positive_Reviews',
    'Company',
    'Registration_Number',
    'Country'
]
filename = "Search Term of health for cats.xlsx"
dfs = pd.DataFrame(columns=Columns)
dfi = pd.read_excel("xlsx inputs/" + filename)
dfi.drop('#', axis = 1, inplace = True)

counter = 1
for index, row in dfi.iterrows():
    print(counter, ": ", end = " ")

    p_dict = {}
    for name in dfi.columns:
        p_dict.update({name: row[name]})

    p_details, s_details = get_details(row["Link"])
    p_dict.update(p_details)

    s_dict = {}
    s_dict.update(s_details)

    print(p_dict["ASIN"],
          p_dict["Seller type"],
          p_dict["Price"],
          p_dict["Rank"],
          p_dict["First_Available"],
          " :::: ",
          s_dict["merchant_ID"],
          s_dict["Seller"],
          s_dict["No. Products"],
          s_dict["S_Reviews_No"],
          s_dict["S_Positive_Reviews"],
          s_dict["Company"],
          s_dict["Registration_Number"],
          s_dict["Country"])
    p_row = pd.DataFrame(p_dict, index=[0])
    s_row = pd.DataFrame(s_dict, index=[0])

    dfp = dfp.append(p_row, ignore_index=True, sort = False)
    if any(value for value in s_details.values()):
        dfs = dfs.append(s_row, ignore_index=True, sort=False)
    counter += 1

writer = pd.ExcelWriter("C:/Users/Euler Graf/Documents/AmazonScraper/xslx outputs/res_" + filename, engine = 'xlsxwriter')

dfs.to_excel(writer, sheet_name = 'rSellers')
dfp.to_excel(writer, sheet_name = 'rProducts')
writer.save()

print("the end folks")
