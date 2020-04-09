import requests


def load_excel_file():
    """Downloading exel file of dividend champs"""

    with open('file_url.txt') as url_file:
        url = url_file.read()

    file = open(r"C:\\Champions.xlsx", "wb")
    print("Downloading file. Please wait.")
    url_to_file = requests.get(url)
    file.write(url_to_file.content)
    file.close()
    print("File downloaded.")


if __name__ == '__main__':
    load_excel_file()
