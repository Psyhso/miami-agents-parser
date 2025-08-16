from selenium import webdriver

from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.webdriver import WebDriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import math


def get_driver():
    """Initializing the driver"""
    service = EdgeService(r'EdgeDrivers\msedgedriver.exe')

    edge_options = webdriver.EdgeOptions()
    edge_options.add_argument("--headless=new")
    edge_options.add_argument("--disable-gpu")  # Отключение GPU (может улучшить стабильность)
    edge_options.add_argument("--window-size=1920,1080")  # Фиксированный размер окна
    edge_options.add_argument("--log-level=3")  # Минимальный уровень логов

    edge_options.add_argument("--disable-blink-features=AutomationControlled")
    edge_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    edge_options.add_experimental_option("useAutomationExtension", False)
    
    edge_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0")

    driver = webdriver.Edge(service=service, options=edge_options)

    return driver


def safe_scroll_and_load(driver, max_attempts=5):
    """Safe scrolling"""
    attempt = 0
    last_height = driver.execute_script("return document.body.scrollHeight")
    
    while attempt < max_attempts:
        try:
            # Прокрутка до низа страницы
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            
            # Ожидание загрузки нового контента
            time.sleep(2)  # Увеличено время ожидания
            
            # Расчет новой высоты прокрутки
            new_height = driver.execute_script("return document.body.scrollHeight")
            
            if new_height == last_height:
                attempt += 1
                if attempt >= max_attempts:
                    break
                time.sleep(2)  # Дополнительное ожидание
                continue
            else:
                attempt = 0  # Сброс счетчика при успешной загрузке
                last_height = new_height
                
        except Exception as e:
            break
    
    return True


def parser(limit, page=''):
    """Parsing a single page"""
    driver = get_driver()
    wait = WebDriverWait(driver, 10)
    base_url = "https://www.partyslate.com/find-vendors/event-planner/area/miami"
    url_page = base_url if page == '' else f"{base_url}?page={page}"
    try:
        driver.get(url_page)
        safe_scroll_and_load(driver)

        wait.until(
            EC.presence_of_all_elements_located(('xpath', "//article[contains(@class, 'src-components-CompanyDirectoryCard-styles__container__2JUdC src-pages-FindVendors-components-Results-styles__card__MdaJ_')]")))
        all_card = driver.find_elements('xpath', "//article[contains(@class, 'src-components-CompanyDirectoryCard-styles__container__2JUdC src-pages-FindVendors-components-Results-styles__card__MdaJ_')]")

        all_url = []
        for card in all_card[:limit]:
            try:
                url = card.find_element('xpath', f".//a[@class='src-components-CompanyDirectoryCard-components-Photos-styles__carousel-image-wrapper__3F6vi']").get_attribute('href')
                all_url.append(url)
            except Exception as e:
                url = card.find_element('xpath', f".//a[@class='src-components-CompanyDirectoryCard-components-Photos-styles__single__3CyLZ']").get_attribute('href')
                all_url.append(url)
    except Exception as e:
        return(f'Ошибка при парсинге страницы: {e}')
    finally:
        driver.quit()

    return all_url


def one_person(driver: WebDriver, wait: WebDriverWait):
    """Parsing a single person"""
    name = wait.until(
        EC.presence_of_element_located(('xpath', "//h3[contains(@class, 'css-1ham2m0')]")))
    name = name.text

    proffesion = wait.until(
        EC.presence_of_element_located(('xpath', "//span[@class='css-1pxun7d']")))
    proffesion = proffesion.text

    return (name, proffesion)


def get_all_url(limit):
    """Getting a url based on the number of pages"""
    all_url = []
    if limit > 20:
        # Получаем первую страницу
        first_page_urls = parser(limit=20)
        all_url.extend(first_page_urls)
        remaining_limit = limit - len(first_page_urls)  # Учитываем реально полученные URL
        
        print(f"Limit: {limit}")
        print(f"Received from the first page: {len(first_page_urls)}")
        print(f"It remains to receive: {remaining_limit}")
        
        if remaining_limit > 0:
            page_num = 2
            while remaining_limit > 0:
                cards_needed = min(20, remaining_limit)
                print(f"Page {page_num}, need some cards: {cards_needed}")
                
                page_urls = parser(limit=cards_needed, page=str(page_num))
                all_url.extend(page_urls)
                remaining_limit -= len(page_urls)  # Учитываем реально полученные URL
                
                print(f"Received from the page {page_num}: {len(page_urls)}")
                print(f"It remains to receive: {remaining_limit}")
                
                page_num += 1
                
                # Защита от бесконечного цикла
                if len(page_urls) == 0:
                    print("The page is empty, we interrupt")
                    break
                    
                if page_num > 10:  # Максимум 10 страниц
                    print("The page limit has been reached")
                    break
    else:
        all_url = parser(limit)

    return all_url


def card_parser(limit=5):
    """Parsing a card"""
    all_url = get_all_url(limit)
    if not all_url:
        print("Error: Missing cards")
        return
    
    driver = None
    try:
        driver = get_driver()
        wait = WebDriverWait(driver, 10)
        results = []
        count = 0

        for url in all_url:
            count+=1
            try:
                print("-------------------------------------------------------------------------------------")
                print(f"{count}. Card on account")
                print(f"\nI'm working on it: {url}")
                driver.get(url)
                
                # Получение названия компании
                title = wait.until(
                    EC.presence_of_element_located(('xpath', "//h1[contains(@class, 'chakra-heading')]")))
                title_company = title.text
                print(f"Titile: {title_company}")

                # Получение соц сетей
                socials = driver.find_elements('xpath', "//a[contains(@class, 'css-6cxgxb')]")
                all_socials = [s.get_attribute('href') for s in socials] if socials else None
                print(f"All socials: {all_socials}")

                # Получение данных о команде
                team = []
                try:
                    driver.find_element('xpath', "//h2[contains(@class, 'css-1xix1js')]//span[@class='css-1bsgmhw']")
                    team.append(one_person(driver, wait))

                    try:
                        count_person = int(driver.find_element('xpath', "//span[@class='css-dw5ttn']").text[-1])
                        print(count_person)
                    except:
                        count_person = 1
                        print(count_person)

                    if count_person > 1:
                        for i in range(count_person-1):
                            next_person = wait.until(
                                EC.element_to_be_clickable(('xpath', "//button[@aria-label='view next team member']")))
                            driver.execute_script("arguments[0].click();", next_person)  # Более надежный клик
                            team.append(one_person(driver, wait))
                    
                except:
                    team = [(None, None)]
                    print('There are no people')

                print(f"Team: {team}")

                #Получение минимальной цены
                minimal_spend = None  # Значение по умолчанию
                try:
                    price_element = wait.until(
                        EC.presence_of_element_located(('xpath', "//dd[@class='css-1nmdp34']"))
                    )
                    minimal_spend = price_element.text.strip() if price_element.text else None
                except (TimeoutException, NoSuchElementException):
                    pass  # Просто пропускаем, если элемент не найден
                except Exception as e:
                    print(f"Unexpected error when receiving the price: {e}")
                print(f"Minimal spend: {minimal_spend}")
                
                # Открытие информации
                open_button = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'chakra-button') and contains(@class, 'css-15tyt09')]")))
                driver.execute_script("arguments[0].click();", open_button)
                time.sleep(1)
                
                # Получение телефона
                try:
                    phone = wait.until(
                        EC.presence_of_element_located(('xpath', '//div[@class="css-0"]//a[contains(@class, "css-1dvr8y4")]')))
                    phone_number = phone.text
                except (TimeoutException, NoSuchElementException):
                    phone_number = None
                print(f"Phone number: {phone_number}")

                # Получение сайта
                try:
                    website = wait.until(
                        EC.presence_of_element_located(('xpath', '//div[@class="css-8atqhb"]//a[contains(@class, "css-123qr35")]')))
                    website = website.text
                except (TimeoutException, NoSuchElementException):
                    website = None
                print(f"Website: {website}")

                for person in team:
                    results.append({
                        'company': title_company,
                        'website': website,
                        'name': person[0],
                        'Job Title': person[1],
                        'phone': phone_number,
                        'email': None,
                        'minimal spend': minimal_spend,
                        'socials': all_socials,
                        'url': url
                    })
                print("-------------------------------------------------------------------------------------")
                time.sleep(1)  # Пауза между запросами
                
            except Exception as e:
                print(f"Error during processing {url}: {str(e)}")
                continue  # Продолжаем со следующего URL
        
        return results if results else "No data could be retrieved from any of the cards"
        
    except Exception as e:
        return f"Critical error: {str(e)}"
    finally:
        if driver:
            driver.quit()


import pandas as pd


def save_excel(limit=5):
    data = card_parser(limit=limit)
    df = pd.DataFrame(data)
    df.to_excel(
        'output.xlsx',
        index=False,          # Не включать индексы
        na_rep='-'         # Как отображать пропущенные значения
    )

