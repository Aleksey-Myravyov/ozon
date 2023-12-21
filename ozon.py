import requests, os, csv, time, random, json, shutil
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import openpyxl as xl
from openpyxl.styles import Border, Side


# bs Принимаем адрес, возвращаем html текст.
def get_html(url):    
    headers = {    
    # Указываем данные пользователя.
    "Accept": "*/*",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
    }
    responce = requests.get(url, headers=headers)
    print(responce.status_code, url)
    return responce.text





def main():
    # удаляем старые файлы и создаем новую папку.
    if os.path.isfile('links_pages.txt'):
        os.remove('links_pages.txt')
    if os.path.isdir('pages'):
        shutil.rmtree('pages')
    os.mkdir('pages')

    # записываем артикли из файла.
    with open('file.csv', 'r') as file:
        reader = csv.reader(file)
        articles = tuple([str(*row) for row in reader])

    # создаем настройки для selenium
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
    options.add_argument('--ignore-certificate-errors')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    # загружаем из файла cookies
    with open('cookies.json', 'r') as file:
        cookies = json.load(file)
    
    # собираем ссылки страниц
    try:
        driver = webdriver.Chrome(options=options)
        driver.get('https://www.ozon.by/')
        time.sleep(random.randint(1, 4))
        driver.delete_all_cookies()
        for cookie in cookies:
            driver.add_cookie(cookie)
        driver.refresh()
        time.sleep(random.randint(1,4))
        for article in articles:
            try:
                url = f'https://www.ozon.by/search/?text={article}&from_global=true'
                print('Проверь название class в названии товара', url)
                driver.get(f'https://www.ozon.by/search/?text={article}&from_global=true')
                time.sleep(random.randint(1, 4))  
                search = get_html(url)
                bs = BeautifulSoup(search, 'lxml')
                # !!!!!!!!!
                data_page = bs.find('a', class_="ti0 tile-hover-target")
                href = 'https://www.ozon.by' + data_page['href'] + '\n'
                with open(f'links_pages.txt', 'a', encoding='utf-8') as file:
                    file.write(str(href))
            except:
                print(f"{'=' * 45}\nERROR article:{article}\n{'=' * 45}")
                # with open(f'links_pages.txt', 'a', encoding='utf-8') as file:
                #     file.write(str(f'Error_in_article_{article}'))
            finally:
                time.sleep(random.randint(1, 4))        
    except Exception as ex:
        print(f"{'-' * 45}\n{ex}\n{'-' * 45}")
    finally:
        driver.close()
        driver.quit()

    # создаем список с искомой информацией
    list_inform = []
    list_inform.append(["Код товара", "Название товара", "URL страницы с товаром", "URL первой картинки", "Цена базовая", "Цена с учетом скидок без Ozon Карты", "Цена по Ozon Карте", "Продавец", "Рейтинг товара"])
    
    # собираем полные данные страниц в папку.
    with open('links_pages.txt', 'r', encoding='utf-8') as file:
        list_links = [line.strip() for line in file]
    try:
        driver = webdriver.Chrome(options=options)
        driver.get('https://www.ozon.by/')
        driver.delete_all_cookies()
        time.sleep(random.randint(1, 4))
        for cookie in cookies:
            driver.add_cookie(cookie)
        driver.refresh()
        time.sleep(random.randint(1,4))
        for url in list_links:
            
            driver.get(url)
            time.sleep(random.randint(3, 5))
            driver.execute_script("window.scrollBy(0,4000)")
            time.sleep(random.randint(4, 5))

            article_data = []
            url_article = url + '!'

            # 1
            try:
                product_code = str(url).split('=')[-1]
                # url = str(url).split('=')[-1]
                print('///', product_code)
            except:
                product_code = '-'
            finally:
                article_data.append(product_code)
                
            url = get_html(url)
            soup = BeautifulSoup(url, 'lxml')
                
            # 2
            try:
                product_name = soup.find("h1", class_="ol").text.strip()
                print(product_name)
            except:
                product_name = '-'
            finally:
                article_data.append(product_name.rstrip('!'))
            
            # 3
            article_data.append(url_article)
            
            # 4
            try:
                url_picture = soup.find("div", class_="n1j jn2").find("img")['src']
            except:
                url_picture = '-'
            finally:
                article_data.append(url_picture)
            
            # 5
            try:
                base_price = soup.find("span", class_="nl4 nl5 nl3 n4l").text
                base_price = base_price.replace('\u2009', '').replace('₽', '')
            except:
                base_price = '-'
            finally:
                article_data.append(base_price)
            
            # 6
            try:
                discount_price = soup.find("span", class_="ln5 l5n nl9").text
                discount_price = discount_price.replace('\u2009', '').replace('₽', '')
            except:
                discount_price = '-'
            finally:
                article_data.append(discount_price)
            
            # 7
            try:
                ozon_price = soup.find("span", class_="l0n lm9").text
                ozon_price = ozon_price.replace('\u2009', '').replace('₽', '')
            except:
                ozon_price = '-'
            finally:
                article_data.append(ozon_price)
            
            # 8
            try:
                saler = soup.find("a", class_="jq9").text
            except:
                saler = '-'
            finally:
                article_data.append(saler)
            
            # 9
            try:
                raiting = soup.find("div", class_="rs3").text.split(' ')[0]
            except:
                raiting = '-'
            finally:
                article_data.append(raiting)
                
            list_inform.append(article_data)
                    
    except Exception as ex:
        print(f"{'*' * 45}\n{ex}\n{'*' * 45}")
    finally:
        driver.close()
        driver.quit()

    # создаем xlsx файл с необходимыми колонками
    print(list_inform)

if __name__=='__main__':
    # делаем рабочую папку, где лежит файл с кодом.
    os.chdir(os.path.dirname(os.path.realpath(__file__)))
    t = str(datetime.now().date())
    start_time = time.time() 
    main()
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Время выполнения: {elapsed_time // 60} минут.")