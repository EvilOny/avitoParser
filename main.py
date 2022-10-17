from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import xlsxwriter

class ExcelWriter:
    def __init__(self, file_name: str = "demo"):
        self.file_name = file_name

    def write_item_list(self, item_list):
        workbook = xlsxwriter.Workbook(f'{self.file_name}.xlsx')
        worksheet = workbook.add_worksheet(f'{product}')
        top_shift = 1
        left_shift = 1

        worksheet.write(top_shift - 1, left_shift, "Название")
        worksheet.write(top_shift - 1, left_shift + 1, "Цена")
        worksheet.write(top_shift - 1, left_shift + 2, "Блок")
        worksheet.write(top_shift - 1, left_shift + 3, "Ссылка")

        for i in range(len(item_list)):
            worksheet.write(i + top_shift, left_shift, item_list[i].title)
            worksheet.write(i + top_shift, left_shift + 1, item_list[i].price)
            worksheet.write(i + top_shift, left_shift + 2, item_list[i].block)
            worksheet.write(i + top_shift, left_shift + 3, item_list[i].link)

        workbook.close()
        
class Item:
    def __init__(self, title: str, price: int, block: str, link: str):
        self.title = title
        self.price = price
        self.link = link
        self.block = block

    def __str__(self):
        return f'{self.title} \nцена: {self.price} \nblock: {self.block} \nlink: {self.link}'

print("Введите название товара:")
product = input()

URL = "https://www.avito.ru/severodvinsk?q=" + product
BASE_URL = "https://www.avito.ru"

PAUSE_DURATION_SECONDS = 1

def main():
    driver.get(URL)
    sleep(PAUSE_DURATION_SECONDS)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    item_list = []
    excel_writer = ExcelWriter(f'{product}')
    
    allItem = soup.findAll('div',  {'data-marker':'item'})
    for item in allItem:
        if item.find('div', {'data-marker':'badge-title-11'}):
            block = "Рыночная цена"
        else:
            block = " "       
        price = item.find('meta', {'itemprop':'price'})['content']
        title = item.find('h3', {'itemprop':'name'}).text
        link = item.find('a')['href']
        if ("/severodvinsk") in link:
            item_list.append(Item(title, price, block, BASE_URL + link))
    
    excel_writer.write_item_list(item_list)

if __name__ == '__main__':
    try:
        service = ChromeService(executable_path=ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        main()
    except Exception as e:
        print(e)
    finally:
        driver.quit()