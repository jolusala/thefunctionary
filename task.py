# Robot to enter weekly sales data into the RobotSpareBin Industries Intranet.
from RPA.Browser.Selenium import Selenium
import re
import time
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files


browser = Selenium()
http = HTTP()
excelFile = Files()

phrase = "USA"
#phrase = input("Insert a phrase to search: ")

print("initializing robot: " + phrase)


def open_website():
    """
This function opens a web browser and navigates to the New York Times website.
It uses the open_available_browser() method from the browser library to open the website.

"""
    browser.open_available_browser("https://www.nytimes.com/")


def find_search_button():
    """
# This function finds and clicks the search button on a webpage.
# It uses the wait_and_click_button method from the browser object,
# which takes an xpath argument to locate the element.

"""

    browser.wait_and_click_button(
        'css:div:nth-child(3) > div.NYTAppHideMasthead.css-1r6wvpq.e1m0pzr40 > header > section.css-9kr9i3.e1m0pzr42 > div.css-qo6pn.ea180rp0 > div.css-10488qs > button')


def insert_phrase_to_search(phrase: str) -> None:
    """

This function takes in a string phrase as an argument
and inserts it into the search bar of a browser.
It then clicks the search button to initiate the search.
The function does not return anything.

"""
    browser.input_text(
        'css:div:nth-child(3) > div.NYTAppHideMasthead.css-1r6wvpq.e1m0pzr40 > header > section.css-9kr9i3.e1m0pzr42 > div.css-qo6pn.ea180rp0 > div.css-10488qs > div > form > div > input', phrase)
    browser.wait_and_click_button(
        'css:div:nth-child(3) > div.NYTAppHideMasthead.css-1r6wvpq.e1m0pzr40 > header > section.css-9kr9i3.e1m0pzr42 > div.css-qo6pn.ea180rp0 > div.css-10488qs > div > form > button')


def store_screenshot():
    # browser.screenshot(
    #   filename="./robot-python/output/screenshot.png")

    browser.screenshot(
        filename="./screenshot.png")
    print('screenshot taken')


def get_category() -> list:
    # create a function to get the category from the user
    categoriesString = input("Insert categories separated by comma(,):")
    categories = categoriesString.split(",")

    return categories


def click_category_button():
    """
# This function waits for and then clicks the button with the xpath
# 'xpath://*[@id="site-content"]/div/div[1]/div[2]/div/div/div[2]/div/div/button'
# using the browser object.

"""
    browser.wait_and_click_button(
        'xpath://*[@id="site-content"]/div/div[1]/div[2]/div/div/div[2]/div/div/button')


def set_category(userCategories):
    """

# This function sets a category for a browser.
# It takes in an argument of 'categories' and checks if it
# is equal to any of the categories listed.
# If it is, it will select the corresponding checkbox on the browser.
# If not, it will print an error message.


    """
    try:
        listaCategorias = []
        listaCategoriasWebelements = browser.get_webelements(
            "css:#site-content > div > div.css-1npexfx > div.css-1az02az > div > div > div:nth-child(2) > div > div > div > ul > li")

        longi = len(listaCategoriasWebelements)

        for i in range(longi):
            categoryWeb = browser.get_webelement(
                "css:#site-content > div > div.css-1npexfx > div.css-1az02az > div > div > div:nth-child(2) > div > div > div > ul > li:nth-child(" + str(i+1) + ") > label").text.split(",")[0]
            palabra = re.findall(
                "^[a-zA-Z]+.+[^0-9]", categoryWeb)
            listaCategorias.append(palabra[0])

        for i in userCategories:
            if i in listaCategorias:
                browser.select_checkbox(
                    'css:#site-content > div > div.css-1npexfx > div.css-1az02az > div > div > div:nth-child(2) > div > div > div > ul > li:nth-child(' + str(listaCategorias.index(i)+1) + ') > label > input')

    except:
        print("Error al seleccionar la categoria")


def get_data_range(fechas: str) -> list:
    # create a function to get the date range from the user and convert to NYtimes format
    fechas = []
    pass


def click_date_range_button():
    """

This function clicks the date range button and selects 
the 6th option from the dropdown menu. 
It uses the browser driver to locate the elements using xpath and clicks them.


    """
    browser.wait_and_click_button(
        'xpath://*[@id="site-content"]/div/div[1]/div[2]/div/div/div[1]/div/div/button')
    browser.wait_and_click_button(
        'xpath://*[@id="site-content"]/div/div[1]/div[2]/div/div/div[1]/div/div/div/ul/li[6]/button')


def set_data_range(star_date, end_date):
    """

# This function sets the data range of a web page by 
# entering the start date and end date provided as parameters. 
# It uses the input_text_when_element_is_visible method to enter the 
# start and end dates into the respective elements with xpaths 
# 'xpath://*[@id="startDate"]' and 'xpath://*[@id="endDate"]'. 
# It then presses the ENTER key on the element with xpath 
# 'xpath://*[@id="endDate"]' to confirm the data range selection.


    """
    browser.input_text_when_element_is_visible(
        'xpath://*[@id="startDate"]', star_date)
    browser.input_text_when_element_is_visible(
        'xpath://*[@id="endDate"]', end_date)
    browser.press_keys('xpath://*[@id="endDate"]', "ENTER")


def click_sort_by_Newest_button():
    """
# This function uses the browser object to select 
# an option from a list by its value. The list is located at 
# the xpath provided and the option selected is "newest". 
# When this function is called, it will click on the "Newest" button.


    """

    browser.select_from_list_by_value(
        'xpath://*[@id="site-content"]/div/div[1]/div[1]/form/div[2]/div/select', "newest")


def obtain_list_news() -> dict:
    """
 This function obtains a list of news from a website and returns them as a dictionary
 . It uses the webdriver library to get the elements from the website and then iterates
 through them, appending each element to a list. For each element, it gets the date, 
 title, paragraph, image URL and image filename and adds them to a dictionary. 
 Finally, it prints out the dictionary and returns it.

    """

    listaWebElements = browser.get_webelements(
        'css:#site-content > div > div:nth-child(2) > div.css-46b038 > ol > li')

    listaVariables = []
    listaNumeros = list(range(1, len(listaWebElements)+1))
    counter = 1

    for i in listaWebElements:
        if "css-1l4w6pd" in i.get_attribute("class"):
            posi = listaWebElements.index(i)

            fecha = browser.get_webelement(
                'css:#site-content > div > div:nth-child(2) > div.css-46b038 > ol > li:nth-child(' + str(posi+1) + ') > div > span').text
            titulo = browser.get_webelement(
                'css:#site-content > div > div:nth-child(2) > div.css-46b038 > ol > li:nth-child(' + str(posi+1) + ') > div > div > div > a > h4').text
            parrafo = browser.get_webelement(
                'css:#site-content > div > div:nth-child(2) > div.css-46b038 > ol > li:nth-child(' + str(posi+1) + ') > div > div > div > a > p.css-16nhkrn').text
            imgURL = (browser.get_webelement(
                'css:#site-content > div > div:nth-child(2) > div.css-46b038 > ol > li:nth-child(' + str(posi+1) + ') > div > div > figure > div > img')).get_attribute("src")
            imagenFileName = "image"+str(counter)+".jpg"
            listaVariables.append(dict(zip(["fecha", "titulo", "parrafo", "imagenUrl", "imagenFileName"],
                                           [fecha, titulo, parrafo, imgURL, imagenFileName])))
            counter += 1

    dictIdNews = dict(zip(listaNumeros, listaVariables))

    # print(dictIdNews)
    return dictIdNews


def count_of_search_phrases(dictIdNews: dict) -> dict:
    """
This function takes in a dictionary of news items, identified by their ID, a
nd counts the number of times a search phrase appears in the title and paragraph
of each news item. It returns a tuple containing the count for titles counts and paragraphs counts. 
"""
    listaConteo = []
    for i in dictIdNews:
        conteoPalabras = 0
        conteoPalabras += dictIdNews[i]["titulo"].count(phrase)
        conteoPalabras += dictIdNews[i]["parrafo"].count(phrase)
        listaConteo.append([i, conteoPalabras])

    dictConteo = dict(listaConteo)
    return dictConteo


def check_contains_money(dictIdNews: dict) -> dict:
    """
This function takes in a dictionary with news IDs as 
keys and values as dictionaries containing the title and 
paragraph of the news article. It uses regular expressions 
to search for money related terms in the titles and 
paragraphs of each article. If it finds a money related 
term, it adds the ID of the article to a list with a 
boolean value of True. If it doesn't find any money related
terms, it adds the ID to the list with a boolean value of False. 
The list is then converted into a dictionary and returned. 
"""

    listaContains = []

    for i in dictIdNews:
        contadorContiene = 0
        titulo = dictIdNews[i]["titulo"]
        parrafo = dictIdNews[i]["parrafo"]

        if re.search(r"^[$][0-9]*.[0-9]*[.][0-9]*", titulo):
            contadorContiene += 1

        if re.search(r"^[$][0-9]*.[0-9]*[.][0-9]*", parrafo):
            contadorContiene += 1

        if re.search(r"[0-9]+.[dollars]+", titulo):
            contadorContiene += 1

        if re.search(r"[0-9]+.[dollars]+", parrafo):
            contadorContiene += 1

        if re.search(r"[0-9]+.[USD]+", titulo):
            contadorContiene += 1

        if re.search(r"[0-9]+.[USD]+", parrafo):
            contadorContiene += 1

        if contadorContiene > 0:
            listaContains.append([i, True])

        if contadorContiene <= 0:
            listaContains.append([i, False])

    dictCheckContains = dict(listaContains)
    return dictCheckContains
    # print(dictCheckContains)


def download_images(dictIdNews: dict):
    """
This function takes in a dictionary of ids and news as an argument. 
It then iterates through the dictionary and stores the image URL and 
filename from each item in the dictionary. It then uses the http library
to download the image from the URL and save it to a folder called "./robot-python/output/images/" 
with the filename specified in the dictionary.
"""

    for i in dictIdNews:
        url = dictIdNews[i]["imagenUrl"]
        # print(url)
        filename = dictIdNews[i]["imagenFileName"]
        http.download(
            url, target_file="./robot-python/output/images/"+filename)


def write_in_Excel(dictIdNews: dict, conteo: dict, checkContains: dict) -> None:
    """

# This function creates an Excel file named "info.xlsx" in the directory
# "./robot-python/output/". It takes three dictionaries as parameters: 
# dictIdNews, conteo, and checkContains. The function then sets the headers 
# of the Excel file to "Title", "Date", "Description", "Picture File Name", 
# "Count of Search", and "True/False". It then iterates through the dictIdNews
# dictionary and sets each value in its corresponding cell in the Excel file. 
# The same is done for conteo and checkContains. Finally, it prints a message and saves the workbook.


    """

    excelFile.create_workbook(
        path="./robot-python/output/info.xlsx", fmt="xlsx")

    listHeaders = ["Title", "Date", "Description",
                   "Picture File Name", "Count of Search", "True/False"]

    for i in range(1, len(listHeaders)+1):
        excelFile.set_cell_value(1, i, listHeaders[i-1])

    for i in range(1, len(dictIdNews)+1):
        excelFile.set_cell_value(i+1, 1, dictIdNews[i]["titulo"])
        excelFile.set_cell_value(i+1, 2, dictIdNews[i]["fecha"])
        excelFile.set_cell_value(i+1, 3, dictIdNews[i]["parrafo"])
        excelFile.set_cell_value(i+1, 4, dictIdNews[i]["imagenFileName"])
        excelFile.set_cell_value(i+1, 5, conteo[i])
        excelFile.set_cell_value(i+1, 6, checkContains[i])

    print('creating excel file')

    excelFile.save_workbook()


def main():
    """

# This code defines a function called main() which performs a series of operations
# . The function starts by attempting to open a website, then finding and clicking
# the search button and inserting a phrase to search. It then clicks the category 
# button and sets the categories, clicks the date range button and sets the date range, 
# clicks the sort by newest button, obtains a list of news items, extracts variables 
# from that list, counts the number of search phrases in each item, checks if each item
# contains money, downloads images associated with each item, writes all of this information
# into an Excel file, and stores a screenshot. Finally it closes the browser.


    """

    try:
        #cat = get_category()
        open_website()

        find_search_button()
        insert_phrase_to_search(phrase)

        click_category_button()

        set_category(["Arts", "Books", "Business", "New York",
                     "Opinion", "Sports", "U.S.", "Week In Review", "World"])

        # set_category(cat)

        store_screenshot()

        click_date_range_button()

        # get_data_range()
        set_data_range("01/01/2020", "02/19/2023")

        click_sort_by_Newest_button()

        time.sleep(1)
        obtain_list_news()

        count_of_search_phrases(obtain_list_news())

        check_contains_money(obtain_list_news())

        download_images(obtain_list_news())

        write_in_Excel(obtain_list_news(),
                       count_of_search_phrases(obtain_list_news()),
                       check_contains_money(obtain_list_news()))

    except Exception as e:
        print(e)

    finally:
        browser.close_browser()


if __name__ == "__main__":
    main()
