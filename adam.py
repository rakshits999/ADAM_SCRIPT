import time
import re
import pandas as pd
from selenium import webdriver
from bs4 import BeautifulSoup
import openpyxl
import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException

input_wb = openpyxl.load_workbook(filename='thirdtask.xlsx')
input_ws = input_wb.active

data = pd.DataFrame(columns=["URL","Country", "Colour","Length","Quantity","Shipping Method","Delivery Time", "Shipping Price","Price", "Total Price"])

driver = webdriver.Chrome("/home/vqcodes-bill/Desktop/client/chromedriver")
driver.maximize_window()

for i, row in enumerate(input_ws.iter_rows(min_row=1, values_only=True)):
    url = row[0]
    print(url) 
    driver.get(url)
       
    country_names = ["Germany","Italy","Poland","United Kingdom","France","Norway","Denmark", "Sweden","Finland","Spain"]
    
    for country_name in country_names:
        try:
            wait = WebDriverWait(driver, 3)
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            movee = driver.find_element(By.CLASS_NAME, "ng-switcher")
            actions = ActionChains(driver)
            actions.move_to_element(movee).perform()

            element = driver.find_element(By.XPATH, "//div[contains(@class, 'nav-global')]//div[contains(@class, 'ng-switcher')]")
            element.click()
            time.sleep(2)

            element5 = driver.find_element(By.CLASS_NAME, value="address-select-trigger")
            element5.click()
            time.sleep(2)

            element = driver.find_element(By.CLASS_NAME, value="filter-input")
            element.send_keys(country_name)
            time.sleep(2)

            element = driver.find_element(By.CLASS_NAME, value="address-select-content")
            element.click()
            time.sleep(3)

            element = driver.find_element(By.CLASS_NAME, "language-selector")
            element.click()
            time.sleep(2)

            element = driver.find_element(By.CLASS_NAME, "search-currency")
            element.send_keys("English")
            time.sleep(2)

            element = driver.find_element(By.XPATH, "//ul[@class='notranslate']")
            element.click()
            time.sleep(2)


            element5 = driver.find_element(By.CLASS_NAME, value="go-contiune-btn")
            element5.click()
            time.sleep(2)
            
            try:
                product_delivery_to = driver.find_element(By.CLASS_NAME, 'product-delivery-to')
            except:
                product_delivery_to = driver.find_element(By.CLASS_NAME, 'delivery--to--EA0FvsN')

            countryy = product_delivery_to.text.strip()
            print(countryy)
            driver.refresh()
        
            sku_title_text = ""

            sku_title = soup.find('span', class_='sku-title-value')
            if sku_title:
                sku_title_text = sku_title.text.strip()
                print(sku_title_text)
            else:           
                sku_item_wrap = driver.find_element(By.CLASS_NAME, 'sku-item--wrap--PyDVB9w')
                sku_images = sku_item_wrap.find_elements(By.CLASS_NAME, 'sku-item--image--mXsHo3h')
                
                for sku_image in sku_images:
                    try:
                        img_tag = sku_image.find_element(By.TAG_NAME, 'img')
                        alt_text = img_tag.get_attribute('alt')
                        print(alt_text)
                        sku_image.click()
                        time.sleep(2)

                     
                        child_divs = driver.find_elements(By.CLASS_NAME, value='sku-item--text--s0fbnzX')
                        if child_divs:
                            for child_div in child_divs:
                                child_div.click()
                                time.sleep(2)
                                length=child_div.text.strip()
                                print(length)

                                product_price = driver.find_element(By.CLASS_NAME, 'product-price-current')
                                price = product_price.text.strip()
                                print(price) 

                                input_element = driver.find_element_by_class_name('quantity--info--Lv_Aw6e')
                                quantity= input_element.text.strip()
                                print(quantity)


                                dynamic_shipping_div = driver.find_element(By.CLASS_NAME, 'dynamic-shipping')
                                dynamic_shipping_div.click()
                                time.sleep(2) 

                                try:
                                    span_element = driver.find_element(By.CLASS_NAME, 'comet-icon-chevrondown')
                                    span_element.click()
                                    time.sleep(2)
                                except:
                                    print()  

                                parent_div = driver.find_element(By.CLASS_NAME, 'comet-modal-body')
                                divs = parent_div.find_elements(By.CLASS_NAME, 'logistics--item--lywJWwT')
                                for div in divs:
                                    div_text = div.text.strip()
                                    # print(div_text)
                                    text= div_text

                                    shipping_price_match = re.search(r'Shipping:\s*(.*)', text)
                                    if shipping_price_match:
                                        shipping_price = shipping_price_match.group(1)
                                    else:
                                        shipping_price = '0'

                                    print('Shipping Price:', shipping_price)

                                    today = datetime.date.today()
                                    
                                    delivery_range_match = re.search(r'Estimated delivery: (\d+-\d+ days )', text)
                                    specific_date_match = re.search(r'Estimated delivery on (\w+ \d+)', text)

                                    delivery_range = delivery_range_match.group(1) if delivery_range_match else None
                                    specific_date = specific_date_match.group(1) if specific_date_match else None
                                    print(specific_date)

                                    if specific_date:
                                        month, day = specific_date.split()
                                        month = datetime.datetime.strptime(month, "%b").month

                                        formatted_date = datetime.date(datetime.date.today().year, month, int(day))
                                        estimated_date = datetime.date(datetime.date.today().year, month, int(day))
                                        days_difference = (estimated_date - today).days
                                        estimated_delivery = f"{days_difference} days"
                                    else:
                                        estimated_delivery = delivery_range

                                    print('Estimated delivery=', estimated_delivery)

                                    shipping_method_match = re.search(r'(From\s.+?)\n', text)
                                    if shipping_method_match:
                                        shipping_method = shipping_method_match.group(1)
                                        print('Shipping Method:', shipping_method)
                                    
                                    price = price.replace(',', '.')
                                    price_match = re.search(r'(\d+(?:\.\d+)?)', price)
                                    price_value = float(price_match.group(1))
                                    print(price_value)

                                    prices = shipping_price.replace(',', '.')

                                    price_match = re.search(r'(\d+(?:\.\d+)?)', prices)
                                    price_value1 = float(price_match.group(1))
                                    print(price_value1)
                                    total_price = 2 * (price_value + price_value1)
                                    print("total price:",total_price)
                                
                                    data = data.append({"URL": url,"Country": countryy, "Colour": alt_text,"Length":length,"Quantity":quantity,"Shipping Method": shipping_method ,"Delivery Time":estimated_delivery,"Shipping Price":shipping_price,"Price":price,"Total Price":total_price}, ignore_index=True)

                                    data.to_excel("output.xlsx", index=False)

                                button = driver.find_element(By.CLASS_NAME, "comet-modal-close")
                                button.click()
                        else:
                            product_price = driver.find_element(By.CLASS_NAME, 'product-price-current')
                            price = product_price.text.strip()
                            print(price) 
                            print("hello")

                            input_element = driver.find_element_by_class_name('quantity--info--Lv_Aw6e')
                            quantity= input_element.text.strip()
                            print(quantity)


                            dynamic_shipping_div = driver.find_element(By.CLASS_NAME, 'dynamic-shipping')
                            dynamic_shipping_div.click()
                            time.sleep(2) 

                            try:
                                span_element = driver.find_element(By.CLASS_NAME, 'comet-icon-chevrondown')
                                span_element.click()
                                time.sleep(2)
                            except:
                                print()  

                            parent_div = driver.find_element(By.CLASS_NAME, 'comet-modal-body')
                            divs = parent_div.find_elements(By.CSS_SELECTOR, '._3yzHC, .logistics--item--lywJWwT')
                            print(len(divs))
                            for div in divs:
                                div_text = div.text.strip()
                                print(div_text)
                                text= div_text
                                

                                shipping_price_match = re.search(r'Shipping:\s*(.*)', text)
                                if shipping_price_match:
                                    shipping_price = shipping_price_match.group(1)
                                else:
                                    shipping_price = '0'

                                print('Shipping Price:', shipping_price)

                                today = datetime.date.today()
                                
                                delivery_range_match = re.search(r'Estimated delivery: (\d+-\d+ days )', text)
                                specific_date_match = re.search(r'Estimated delivery on (\w+ \d+)', text)

                                delivery_range = delivery_range_match.group(1) if delivery_range_match else None
                                specific_date = specific_date_match.group(1) if specific_date_match else None
                                print(specific_date)

                                if specific_date:
                                    month, day = specific_date.split()
                                    month = datetime.datetime.strptime(month, "%b").month

                                    formatted_date = datetime.date(datetime.date.today().year, month, int(day))
                                    estimated_date = datetime.date(datetime.date.today().year, month, int(day))
                                    days_difference = (estimated_date - today).days
                                    estimated_delivery = f"{days_difference} days"
                                else:
                                    estimated_delivery = delivery_range

                                print('Estimated delivery=', estimated_delivery)

                                shipping_method_match = re.search(r'(From\s.+?)\n', text)
                                if shipping_method_match:
                                    shipping_method = shipping_method_match.group(1)
                                    print('Shipping Method:', shipping_method)
                                
                                price = price.replace(',', '.')
                                price_match = re.search(r'(\d+(?:\.\d+)?)', price)
                                price_value = float(price_match.group(1))
                                print(price_value)

                                prices = shipping_price.replace(',', '.')

                                price_match = re.search(r'(\d+(?:\.\d+)?)', prices)
                                price_value1 = float(price_match.group(1))
                                print(price_value1)
                                total_price = 2 * (price_value + price_value1)
                                print("total price:",total_price)
                            
                                data = data.append({"URL": url,"Country": countryy, "Colour": alt_text,"Length":"Not found","Quantity":quantity,"Shipping Method": shipping_method ,"Delivery Time":estimated_delivery,"Shipping Price":shipping_price,"Price":price,"Total Price":total_price}, ignore_index=True)

                                data.to_excel("output.xlsx", index=False)
                                

                            button = driver.find_element(By.CLASS_NAME, "comet-modal-close")
                            button.click()
                                

                            
                           

                    except:
                        continue

        except:
            print(f"The product is not delivered in {country_name}. Skipping...")
            data = data.append({"URL": url, "Country": country_name, "Colour": "Not found","Length":"Not found",
                            "Shipping Method": "Not found", "Delivery Time": "Not found",
                            "Shipping Price": "Not found", "Price": "Not found", "Total Price": "Not found"},
                           ignore_index=True)
        
            data.to_excel("output.xlsx", index=False)
            continue
            
driver.quit()
