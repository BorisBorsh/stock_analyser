import requests


def download_excel_data_file(excel_data_file_path):
    """Downloading exel file of dividend champs"""

    with open('file_url.txt') as url_file:
        url = url_file.read()

    file = open(excel_data_file_path, "wb")
    print("Downloading file. Please wait")
    url_to_file = requests.get(url)
    file.write(url_to_file.content)
    file.close()
    print("File downloaded")

