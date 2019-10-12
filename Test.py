#modules used in this project
import bs4, requests, webbrowser

#string of the url
url = "https://www.microcenter.com/category/4294966937/video-cards"

#open url in browser
#webbrowser.open(url)

#requesting data
gpu_data = requests.get(url)


# documentation: https://docs.python.org/3/tutorial/inputoutput.html#reading-and-writing-files
with open("gpu_file", "wb") as f:
    f.write(gpu_data.content)


#telling beautiful soup that we want to parse html
s_obj = bs4.BeautifulSoup(gpu_data.content, 'html.parser')

#find list tag
products = s_obj.findAll("li", {"class": "product_wrapper"})

length = len(products)


#some list varibles to store the information in
brand_model = []
brand_brand = []
product_price = []
product_link = []

#gets brand of all products page
def get_brand():

    for i in range(length):
        brand = products[i]("a")[1]
        brand_brand.append(brand["data-brand"])
        print(brand["data-brand"])

    #print(brand_list)


#gets model of all products on page
def get_model():

    for i in range(length):
        model = products[i]("a")[1]
        brand_model.append(model["data-name"])
        print(model["data-name"])

    #print(brand_name)


#gets price of all products on page
def get_price():

    for i in range(length):
        price = products[i]("a")[1]
        product_price.append(price["data-price"])
        print("$" + price["data-price"])

    #print(product_price)

#gets link to product
def get_link():

    for i in range(length):
        link = products[i]("a")[1]
        product_link.append("microcenter.com" + link["href"])
        print("microcenter.com" + link["href"])

    #print(product_link)



get_brand()
get_model()
get_price()
get_link()
