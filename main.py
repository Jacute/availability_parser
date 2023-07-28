from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from datetime import datetime

from config.aviable import *
from config.config import *

import os
import sys
import traceback
import argparse


class Parser:
    def __init__(self):
        self.result = []
        parser = argparse.ArgumentParser(description='Process some integers.')
        parser.add_argument('--headless', action='store_true', help='headless')
        args = parser.parse_args()
        if args.headless:
            self.driver = self.get_driver(True)
        else:
            self.driver = self.get_driver(False)

    def get_driver(self, headless):
        try:
            options = webdriver.ChromeOptions()
            if headless:
                options.add_argument('--headless')
                options.add_argument('--disable-gpu')

            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)

            options.add_argument(
                "user-agent=Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:84.0) Gecko/20100101 Firefox/84.0")

            # options.add_argument('--disable-dev-shm-usage')
            # options.add_argument('--no-sandbox')

            driver = webdriver.Chrome(
                service=Service('chromedriver.exe'),
                options=options
            )
            driver.set_window_size(1920, 1080)
            driver.implicitly_wait(30)

            self.wait = WebDriverWait(driver, 30)

            return driver
        except Exception as e:
            print('Неудачная настройка браузера!')
            print(traceback.format_exc())
            print(input('Нажмите ENTER, чтобы закрыть эту программу'))
            sys.exit()

    def parse(self, articles):
        result = {}
        parsed_articles = []
        for i in articles:
            print(f'{articles.index(i) + 1} of {len(articles)}')

            parts_of_article = i.split('_')
            if len(parts_of_article) == 3:
                prefix, article, size = parts_of_article
            elif len(parts_of_article) > 3:
                prefix, article, size = parts_of_article[0], parts_of_article[1], '_'.join(parts_of_article[2])
            elif len(parts_of_article) == 2:
                prefix, article, size = parts_of_article[0], parts_of_article[1], None
            else:
                print(i)
                raise Exception('Article has less than 3 parts')
            if prefix != 'H&M' or i in parsed_articles:
                continue
            parsed_articles.append(i)
            url = f'https://www2.hm.com/pl_pl/productpage.{article}.html'
            self.driver.get(url)
            price = self.driver.find_element(By.ID, 'product-price').text.replace(' PLN', '').replace(',', '.')
            if len(price.split('\n')) > 1: price = price.split('\n')[1]  # На случай если есть скидка
            if size:
                sizes = self.driver.find_elements(By.XPATH, '//hm-size-selector/ul/li/label')
                for elem in sizes:
                    new_article = prefix + '_' + article + '_' + elem.text.split('\n')[0]
                    if 'Zostało tylko kilka sztuk!' in elem.text:
                        result[new_article] = [AVIABLE["few_items"], self.get_price(price)]
                    elif elem.get_attribute('aria-disabled') == 'true':
                        result[new_article] = [AVIABLE["no_aviable"], self.get_price(price)]
                    else:
                        result[new_article] = [AVIABLE["aviable"], self.get_price(price)]
            else:  # Для сумок
                new_article = prefix + '_' + article
                btn = self.driver.find_element(By.CLASS_NAME, 'item.button.fluid')
                if 'Dodaj' not in btn.text:
                    result[new_article] = [AVIABLE["no_aviable"], self.get_price(price)]
                else:
                    result[new_article] = [AVIABLE["aviable"], self.get_price(price)]
        return result

    def gPriceDict(self, key):
        return float(PRICE_TABLE[key])

    def get_price(self, pln_price):
        cost_price = ((float(pln_price) / self.gPriceDict("КУРС_USD_ЗЛОТЫ")) * self.gPriceDict("КОЭФ_КОНВЕРТАЦИИ") * self.gPriceDict(
            'КУРС_USD_RUB')) + (self.gPriceDict('ЦЕНА_ДОСТАВКИ_В_КАТЕГОРИИ') * self.gPriceDict('КУРС_БЕЛ.РУБ_РУБ') * self.gPriceDict(
            'КУРС_EUR_БЕЛ.РУБ'))
        final_price = ((cost_price + self.gPriceDict('СРЕД_ЦЕН_ДОСТАВКИ')) * self.gPriceDict('НАЦЕНКА')) / (
                    1 - self.gPriceDict('ПРОЦЕНТЫ_ОЗОН') - self.gPriceDict('ПРОЦЕНТЫ_НАЛОГ') - self.gPriceDict('ПРОЦЕНТЫ_ЭКВАЙРИНГ'))

        if final_price > 20000:
            final_price = (final_price // 1000 + 1) * 1000 - 1
        elif final_price > 10000:
            if final_price % 1000 >= 500:
                final_price = (final_price // 1000) * 1000 + 999
            else:
                final_price = (final_price // 1000) * 1000 + 499
        else:
            final_price = (final_price // 100 + 1) * 100 - 1
        return final_price

    def save(self, result):
        wb = load_workbook(filename=f'templates/{TEMPLATE_NAME}')
        ws = wb['Остатки на складе']

        for i in range(2, ws.max_row + 1):
            try:
                ws.cell(row=i, column=4).value = result[ws['B' + str(i)].value][0]
                ws.cell(row=i, column=6).value = result[ws['B' + str(i)].value][1]
                ws.cell(row=i, column=1).value = STOCK_NAME
            except KeyError:
                pass

        wb.save(SAVE_XLSX_PATH + f"{datetime.now()}.xlsx".replace(':', '.'))

    def get_articles(self):
        articles = []

        wb = load_workbook(filename=f'templates/{TEMPLATE_NAME}')
        ws = wb['Остатки на складе']
        for i in range(2, ws.max_row + 1):
            articles.append(ws['B' + str(i)].value)

        return articles

    def get_urls_sizes(self, articles):
        urls_sizes = {}
        for article in articles:
            prefix, article, size = article.split('_')
            if prefix == 'H&M':
                url = f'https://www2.hm.com/pl_pl/productpage.{article}.html'
                if url in urls_sizes.keys():
                    urls_sizes[url] += [size]
                else:
                    urls_sizes[url] = [size]
        return urls_sizes

    def check_exists_by_xpath(self, elem, xpath):
        try:
            elem.find_element(By.XPATH, xpath)
        except NoSuchElementException:
            return False
        return True

    def delete_duplicates(self, articles):
        result = []
        tmp = []
        for article in articles:
            if article and article.split('_')[1] not in tmp:
                tmp.append(article.split('_')[1])
                result.append(article)
        return result

    def start(self):
        try:
            print('--- START PARSING ---')
            articles = self.delete_duplicates(self.get_articles())
            result = self.parse(articles)
            print('--- END PARSING ---')
        except:
            error = self.driver.current_url + '\n' + traceback.format_exc() + '\n'
            print(error)
            with open('log.log', 'a') as f:
                f.write(error)
        finally:
            self.save(result)
            self.driver.close()
            self.driver.quit()


def main():
    parser = Parser()
    parser.start()


if __name__ == '__main__':
    if 'xlsx' not in os.listdir():
        os.mkdir('xlsx')
    main()
