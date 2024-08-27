import requests
from bs4 import BeautifulSoup
import gspread
import time
import json
import re



base_url = 'https://www.kohls.com'
headers = {
    'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'
}

# For the google sheet
gc = gspread.service_account(filename='credentials.json')

#Gets the URL and Google Sheet Link
page_url = input("Enter the URL: ")
sheet_key = input("Enter the Google Sheet Link: ")
sh = gc.open_by_key(sheet_key[39:83])

worksheet = sh.sheet1
page_number = 1
more_pages = True
product_number_total=0

cells=0
items=0
information=[None, None]
while more_pages:
    try:
        r = requests.get(page_url, headers=headers)
        r.raise_for_status()  # raises an exception if the request was not successful
        # print(r.content)
    except requests.exceptions.RequestException as e:
        print(e)
        print('An error occurred while making the request')
    else:
        # request was successful
        soup = BeautifulSoup(r.content, 'lxml')

    # Makes product list equal to all divs with class products_grid
    product_list = soup.find_all('li', class_='products_grid')
    for product in product_list:
        product_number = product_list.index(product)+1+ product_number_total
        print("product #",product_number,": Item #", items,": Cell #", cells)
        # Extract the name of the product
        name = product.find('div', class_='prod_nameBlock').text
        name = name.strip()

        # Extract the link to the product page
        link = product.find('div', class_='prod_img_block').find('a')['href']  
        link=base_url+link
    
        try:
            UPCresponse = requests.get(link, headers=headers)
            UPCresponse.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(e)
            print('An error occurred while making the request')
        else:
            # request was successful
            # find the productV2JsonData variable in the JavaScript code
            pattern = r'var productV2JsonData = ({.*?});'
            match = re.search(pattern, UPCresponse.text, re.DOTALL)

            #Check if the item was found
            if match:
                upc_price=[]
                # extract the JSON data from the matched group
                json_data = match.group(1)

                # parse the JSON data
                product_data = json.loads(json_data)
                
                #pulls upc and price from json data
                for product in product_data['SKUS']:
                    #upcs+=1
                    upc = product['UPC']['ID']
                    price = product['price']['lowestApplicablePrice']#['minPrice']
                    availability = product['availability']
                    upc = str(upc)
                    price = str(price)
                    #print(price)
                    if availability != "Out of Stock":
                        #information.append({'upc':upc, 'price':price})
                        #information.append(upc)
                        #information.append(price)
                        #information.append(upc_price)
                        #user = upc, price
                        information.append(upc)
                        information.append(price)
                        if len(information) >=  166:
                            worksheet.insert_row(information, 2)
                            cells+=168
                            items+=166
                            time.sleep(1)
                            information=[None, None]
            else:
                print('Product data not found in JavaScript code')

    # Check if there is a next page
    pagination_block = soup.find('div', class_='prodsPagination_block fr')
    if pagination_block:
        # Find the link to the next page
        print('There is a next page')
        page_number += 1
        
        #next_page_link = pagination_block.find('a', class_='ce-pgntn nextArw fr', data_oprtn="next")
        #next_page_link = pagination_block.find('a', class_=['ce-pgntn', 'nextArw', 'fr'], data_oprtn="next")
        link_element = soup.find('link', rel='next')
        if link_element == None:
            worksheet.insert_row(information, 2)
            more_pages = False
            break
        else:
            next_page_link = link_element['href']
            print(next_page_link)
            #cells+=1
            #worksheet.insert_row(next_page_link, 2)
            if next_page_link:
                # Update the page number and URL
                page_number += 1
                product_number_total = product_number
                page_url = base_url + next_page_link
                # Make a request to the next page
                r = requests.get(page_url, headers=headers)
                soup = BeautifulSoup(r.content, 'lxml')
            else:
                # There is no next page
                more_pages = False
    else:
        # There is no pagination block, so there must be only one page
        more_pages = False

print("product #",product_number,": Item #", items,": Cell #", cells)
print("POGGERS!")
