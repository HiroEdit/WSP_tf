import bs4
import requests
import xlsxwriter
import time

product_manufacturer = []
product_model = []
product_price = []
product_link = []
product_stock = []
product_reviews = []
product_stars = []
product_rebate = []
product_image = []
product_sku = []
product_offer = []
product_availability = []

# start of main program function
def get_url(url):
    # sending http request to our url
    gpu_data = requests.get(url)

    # documentation: https://docs.python.org/3/tutorial/inputoutput.html#reading-and-writing-files
    with open("gpu_file", "wb") as f:
        f.write(gpu_data.content)

    # telling beautiful soup that we want to parse html
    s_obj = bs4.BeautifulSoup(gpu_data.content, 'html.parser')

    # product wrappers. Each wrapper contains all the info about the product
    products = s_obj.findAll("li", {"class": "product_wrapper"})

    # however many products on the page
    length = len(products)

    # gets manufacturer of all products page
    def get_manufacturer():
        for i in range(length):
            brand = products[i].findAll("a")[1]
            product_manufacturer.append(brand["data-brand"])
            # print(brand["data-brand"])

    # gets model name of all products on page
    def get_model():
        for i in range(length):
            model = products[i].findAll("a")[1]
            product_model.append(model["data-name"])
            # print(model["data-name"])

    # gets price of all products on page
    def get_price():
        for i in range(length):
            price = products[i].findAll("a")[1]
            product_price.append(price["data-price"])
            # print("$" + price["data-price"])

    # gets link to product
    def get_link():
        for i in range(length):
            link = products[i].findAll("a")[1]
            product_link.append("https://www.microcenter.com" + link["href"])
            # print("microcenter.com" + link["href"])

    # gets reviews of all products
    def get_reviews():
        for i in range(length):
            review = products[i].span.text
            product_reviews.append(review.strip().replace("(", "").replace(")", ""))
            # print(review)

    # gets products rebate price
    def get_rebate():
        for i in range(length):
            rebate = products[i].find("span", {"class": "price"}).text

            if len(rebate) == 0:
                product_rebate.append("0.00")
            else:
                product_rebate.append(rebate.replace("$", "").replace(",", ""))

    # gets image link of product
    def get_image():
        for i in range(length):
            find_image = products[i].find("img")
            image = find_image.get("src")
            product_image.append(image)
            # print(image)

    # get product SKU number
    def get_sku():
        for i in range(length):
            sku = products[i].p.text
            product_sku.append(sku)

    # gets offer if there is one
    def get_offer():
        for i in range(length):
            offer = products[i].find("div", {"class": "highlight clear"}).text

            if len(offer) == 0:
                product_offer.append("------")

            else:
                product_offer.append(offer)
                # print(offer)

    # shows whether product can be purchased online or in store
    def get_availability():
        for i in range(length):
            avail = products[i].find("p", {"class": "limit"})

            if avail is not None:
                avail = products[i].find("p", {"class": "limit"}).text  # In-store only
                product_availability.append(avail)
                # print(avail)

            else:
                unavail = products[i].find("p", {"class": "limitNoSale"})

                if unavail is not None:
                    unavail = products[i].find("p", {"class": "limitNoSale"}).text  # unavailable online
                    product_availability.append(unavail)
                    # print(unavail)

    # get the star rating of the product
    def get_stars():
        for i in range(length):
            # using ratingstars as start point because if a product doesn't have a star rating, the find in our ELSE statement will fail
            # all products have a ratingstars attribute
            stars = products[i].find("div", {"class": "ratingstars"})
            check_for_0 = stars.div.span.text  # should always be 0 reviews for products without rating

            if check_for_0 == "0 Reviews":
                product_stars.append("No star rating")

            else:
                find_stars = products[i].findAll("img")[2]
                stars = find_stars.get("alt")
                product_stars.append(stars)

    # gets stock of all products ************FIX***********
    # def get_stock():
    #
    #     for i in range(length):
    #
    #         stock = products[i].strong.text
    #         product_stock.append(stock)
    #         print(stock)

    get_manufacturer()
    get_model()
    get_price()
    get_rebate()
    get_reviews()
    get_stars()
    get_link()
    get_image()
    get_sku()
    get_offer()
    get_availability()


get_url(
    "https://www.microcenter.com/search/search_results.aspx?N=4294966937&NTK=all&NR=&sku_list=&page=1&myStore=false")
time.sleep(7)
get_url(
    "https://www.microcenter.com/search/search_results.aspx?N=4294966937&NTK=all&NR=&sku_list=&page=2&myStore=false")
time.sleep(7)
get_url(
    "https://www.microcenter.com/search/search_results.aspx?N=4294966937&NTK=all&NR=&sku_list=&page=3&myStore=false")
time.sleep(7)
get_url(
    "https://www.microcenter.com/search/search_results.aspx?N=4294966937&NTK=all&NR=&sku_list=&page=4&myStore=false")

# converting the price and rebate lists to contain floats because we want to apple to currency format in excel
for i in range(len(product_price)):
    product_price[i] = float(product_price[i])

for i in range(len(product_rebate)):
    product_rebate[i] = float(product_rebate[i])

# searching our lists for the index with the greatest length to use that for column formatting
manufacturer_w = max(product_manufacturer, key=len)
model_w = max(product_model, key=len)
reviews_w = max(product_reviews, key=len)
link_w = max(product_link, key=len)
image_w = max(product_image, key=len)
offer_w = max(product_offer, key=len)
availability_w = max(product_availability, key=len)

################################################## Start Excel ##################################################
# https://xlsxwriter.readthedocs.io/tutorial02.html

workbook = xlsxwriter.Workbook("GPU_Products.xlsx")
worksheet1 = workbook.add_worksheet()

# adding some formatting to the column titles and prices
title = workbook.add_format({"font_size": 12, "font_color": "red", "bold": True})
money = workbook.add_format({"num_format": "$#,##0.00"})

# setting width of the columns to fit our information
worksheet1.set_column("A:A", len(manufacturer_w))
worksheet1.set_column("B:B", len(model_w) - 15)
worksheet1.set_column("C:C", 10)  # price
worksheet1.set_column("D:D", 12)  # rebate price
worksheet1.set_column("E:E", len(reviews_w))
worksheet1.set_column("F:F", 11)  # star rating
worksheet1.set_column("G:G", 15)  # SKU number
worksheet1.set_column("H:H", len(offer_w) - 10)
worksheet1.set_column("I:I", len(availability_w))
worksheet1.set_column("J:J", len(link_w) - 15)
worksheet1.set_column("K:K", len(image_w) - 10)

# naming our column titles and applying our custom format
worksheet1.write("A1", "Manufacturer", title)
worksheet1.write("B1", "Model", title)
worksheet1.write("C1", "Price", title)
worksheet1.write("D1", "Rebate Price", title)
worksheet1.write("E1", "Reviews", title)
worksheet1.write("F1", "Stars", title)
worksheet1.write("G1", "Sku #", title)
worksheet1.write("H1", "Offer", title)
worksheet1.write("I1", "Availability", title)
worksheet1.write("J1", "Product Link", title)
worksheet1.write("K1", "Image Link", title)

# write data from our lists
for i in range(len(product_manufacturer)):
    worksheet1.write(i + 1, 0, product_manufacturer[i])  # A
    worksheet1.write(i + 1, 1, product_model[i])  # B
    worksheet1.write(i + 1, 2, product_price[i], money)  # C
    worksheet1.write(i + 1, 3, product_rebate[i], money)  # D
    worksheet1.write(i + 1, 4, product_reviews[i])  # E
    worksheet1.write(i + 1, 5, product_stars[i])  # F
    worksheet1.write(i + 1, 6, product_sku[i])  # G
    worksheet1.write(i + 1, 7, product_offer[i])  # H
    worksheet1.write(i + 1, 8, product_availability[i])  # I
    worksheet1.write(i + 1, 9, product_link[i])  # J
    worksheet1.write(i + 1, 10, product_image[i])  # K

workbook.close()
