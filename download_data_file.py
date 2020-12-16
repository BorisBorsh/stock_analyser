import requests
from bs4 import BeautifulSoup


def download_excel_data_file(excel_data_file_path):
    """Downloading excel file of dividend champs"""

    with open('file_url.txt') as url_file:
        urls = url_file.read().splitlines()
        file_url = urls[0]
        file_date_url = urls[1]

    #Print info on updated data file
    file_date_page = requests.get(file_date_url)
    soup = BeautifulSoup(file_date_page.content, 'html.parser')
    file_update_date = soup.find_all("em")[2].get_text()
    print("Data file " + file_update_date)

    file = open(excel_data_file_path, "wb")
    print("Downloading file. Please wait")
    url_to_file = requests.get(file_url)
    file.write(url_to_file.content)
    file.close()
    print("File downloaded")

