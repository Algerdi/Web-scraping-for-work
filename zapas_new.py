import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
import time
import datetime
import openpyxl
import pprint

mozilla_driver_path = 'C:/Users/kzlvv/Desktop/monit/geckodriver.exe'
driver = webdriver.Firefox(executable_path=mozilla_driver_path)
driver.get('https://ips3.belgiss.by/indexIps.php')
login = driver.find_element_by_xpath('//*[@id="form_auth_login"]')
password = driver.find_element_by_xpath('//*[@id="form_auth_password"]')
submit = driver.find_element_by_xpath('//*[@id="form_auth_submit"]')
login.send_keys('secret')
password.send_keys('company')
submit.click()
# time.sleep(5)
# perech = driver.find_element_by_xpath('//*[@id="ui-accordion-DivTipkatCheckboxList-header-0"]/span')
# perech.click()
# tkp = driver.find_element_by_xpath('//*[@id="TipkatCheckbox3"]')
# tkp.click()
# n_r = driver.find_element_by_xpath('//*[@id="TipkatCheckbox23"]')
# n_r.click()
# mezh = driver.find_element_by_xpath('//*[@id="TipkatCheckbox81"]')
# mezh.click()
# n_b = driver.find_element_by_xpath('//*[@id="TipkatCheckbox101"]')
# n_b.click()
# perech.click()

time.sleep(3)
driver.execute_script("window.open()")
driver.switch_to.window(driver.window_handles[1])
driver.get("https://standards.cen.eu/dyn/www/f?p=CENWEB:105::RESET::::")
published = driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/form/div/table/tbody/tr[4]/td/div/div[5]/label/input')
published.click()

time.sleep(3)
driver.execute_script("window.open()")
driver.switch_to.window(driver.window_handles[2])
driver.get("https://www.iso.org/ru/advanced-search/x/")
driver.switch_to.window(driver.window_handles[0])

# Функции для работы по сайту

# ТУТ МОЖНО ОБЪЕДИНИТЬ В ОДНУ ФУНКЦИЮ ЛИШЬ ДОБАВЛЯЯ ЦИФРЫ В ССЫЛКИ
def if_iso():
    perech = driver.find_element_by_xpath('//*[@id="ui-accordion-DivTipkatCheckboxList-header-0"]/span')
    perech.click()
    otmen = driver.find_element_by_xpath('//*[@id="ui-accordion-DivTipkatCheckboxList-panel-0"]/a[2]')
    otmen.click()
    isos = driver.find_element_by_xpath('//*[@id="TipkatCheckbox33"]')
    iecs = driver.find_element_by_xpath('//*[@id="TipkatCheckbox34"]')
    isos.click()
    iecs.click()
    perech.click()

def if_en():
    perech = driver.find_element_by_xpath('//*[@id="ui-accordion-DivTipkatCheckboxList-header-0"]/span')
    perech.click()
    otmen = driver.find_element_by_xpath('//*[@id="ui-accordion-DivTipkatCheckboxList-panel-0"]/a[2]')
    otmen.click()
    cens = driver.find_element_by_xpath('//*[@id="TipkatCheckbox35"]')
    cenelec = driver.find_element_by_xpath('//*[@id="TipkatCheckbox36"]')
    cens.click()
    cenelec.click()
    perech.click()


def vvod_ips(doc):
    vvod = driver.find_element_by_xpath('//*[@id="fullseek"]')
    vvod_but = driver.find_element_by_xpath('//*[@id="SearchSimpleForm"]/table/tbody/tr/td[2]/button/span')
    vvod.clear()
    if 'EN' in doc:
        if_en()
    else:
        if_iso()
    vvod.send_keys(doc)
    vvod_but.click()
    # print('вводим в поиск номер документа')
    # print(doc)
    time.sleep(4)
    table = driver.find_element_by_xpath('//*[@id="FlexTnpaSearchSimple"]/tbody')
    try:
        stroka = table.find_element_by_class_name('jqgrow')
        # print('проверяю на наличие в ипс')
        # print(stroka.text)
        info_stand = [x.text.strip() for x in stroka.find_elements_by_tag_name('td')]
        info_stand[7] = info_stand[7].split(' \n')
        print(info_stand)
        return info_stand
    except NoSuchElementException:
        return ['Отменен']
    except StaleElementReferenceException:
        print('ПОШЛА ОШИБКА В ЦИКЛЕ VVOD_IPS')
        vvod_ips(doc)





# def not_active(result):
#     for elem in result[7]:
#         vivod = vvod_ips(elem)

def stage(result, prop, tnpas, tnpa):
    if result is None:
        ex[tnpas][tnpa][4] = 'Заменен с NONETYPE'
        ex[tnpas][tnpa][5].append(prop)
        print('СТУПЕНЬ С NONETYPE на STAGE и BASIS_CODE')
    elif result[-1] == 'Недействующий НД':
        # print('Не действующий')
        basis_code(result[7], tnpas, tnpa)
        print('ступень')
    elif result[-1] == 'Отменен':
        # print('Базовый документ отменен')
        ex[tnpas][tnpa][4] = 'Отменен'
        print('ступень')
    else:
        # print('Действующий')
        ex[tnpas][tnpa][4] = 'Заменен'
        ex[tnpas][tnpa][5].append(prop)
        ex[tnpas][tnpa][6].append(result[5])
        # print('версия проверки без учета анализа изменений')
        which_doc(prop, tnpas, tnpa)
        # print('указываем скорректированную версию')
        # print(ex[tnpas][tnpa])
        print('ступень')


def basis_code(i, tnpas, tnpa):
    # print('ы не прошли на активность документ, поэтому ищем действующие его версии')
    for mom in i:
        result = vvod_ips(mom)
        print('очередной result прошел')
        stage(result, mom, tnpas, tnpa)
        print('stage прошел')


# Проверяем через какой сайт пропускать. CEN или ISO.
def which_doc(i, tnpas, tnpa):
    if 'EN' in i:
        driver.switch_to.window(driver.window_handles[1])
        cen_search = driver.find_element_by_xpath('//*[@id="STAND_REF"]')
        cen_search.clear()
        cen_search.send_keys(i)
        cen_butt = driver.find_element_by_xpath('//*[@id="tformsub1"]')
        cen_butt.click()
        time.sleep(3)
        table_cen = driver.find_element_by_xpath('/html/body/div/div[3]/div[2]/div[3]/div[2]/div/div[2]/table/tbody')
        stroka_cen = table_cen.find_elements_by_css_selector('strong a')
        if len(stroka_cen) > 1:
            for adds in stroka_cen[1:]:
                ex[tnpas][tnpa][7].append(adds.text)
        print('прошел проверку EN')
    else:
        if 'TS' in i:
            numb = i[7:i.find(':')]
        else:
            numb = i[4:i.find(':')]
        if '-' in i:
            chis = numb[:numb.find('-')]
            oboz = numb[len(chis) + 1:]
        else:
            chis = numb
            oboz = ''
        # print('выводим номер, число и часть ИСО')
        # print(numb, chis, oboz)
        driver.switch_to.window(driver.window_handles[2])
        iso_search_chis = driver.find_element_by_xpath('//*[@id="formISONumber"]')
        iso_search_oboz = driver.find_element_by_xpath('//*[@id="formPartNumber"]')
        iso_button = driver.find_element_by_xpath('//*[@id="advancedSearchForm"]/div[2]/button[1]')
        iso_search_chis.clear()
        iso_search_oboz.clear()
        iso_search_chis.send_keys(chis)
        iso_search_oboz.send_keys(oboz)
        iso_button.click()
        time.sleep(3)
        block = driver.find_element_by_xpath('//*[@id="search-results"]/div/ul')
        lis = block.find_elements_by_css_selector('li')
        for li in lis:
            block_doc = li.find_element_by_css_selector('a')
            if f'{i}/' in block_doc.text:
                ex[tnpas][tnpa][7].append(block_doc.text)
        print('прошел проверку ISO')
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(2)



SHEETY_IZ_ENDPOINT = 'https://api.sheety.co/bb88d1f903af5db2b8c236edfde94c48/monitAugust2020/test'

ex = {}

response_iz = requests.get(url=SHEETY_IZ_ENDPOINT)
data = response_iz.json()
docs = data['test']

# print(docs)


for i in range(len(docs)-1):
    if docs[i]['tnpa'] != '':
        ex[docs[i]['tnpa']] = {}

for i in range(len(docs)):
    if docs[i]['tnpa'] != '':
        flag = docs[i]['tnpa']
        name = docs[i]['oboznachTnpa']
        # До это заменен/отменен/действует, до заменяющие новые версии, предпоследнее даты введения, последнее изменение в документе.
        ex[flag][name] = [docs[i]['naimTnpa'].strip(), docs[i]['dataVved'].strip(), [docs[i]['bazDoc'].replace('\xa0', '')], [docs[i]['uchten']], '', [], [], []]
    else:
        if docs[i]['bazDoc'] != '':
            ex[flag][name][2].append(docs[i]['bazDoc'].replace('\xa0', ''))
        if docs[i]['uchten'] != '' and docs[i]['uchten'] != '-':
            ex[flag][name][3].append(docs[i]['uchten'])

# print(ex)
# fock = 1
cock = 1
for tnpas in ex:
    for tnpa in ex[tnpas]:
        # print(ex[tnpas][tnpa][2])
        print('начало')
        for i in ex[tnpas][tnpa][2]:
            try:
                result = vvod_ips(i)
                print('result прошел')
                if result[-1] == 'Недействующий НД':
                    # print('Не действующий')
                    basis_code(result[7], tnpas, tnpa)
                    print(cock, tnpa, ex[tnpas][tnpa])
                    now = datetime.datetime.now()
                    print('*****цикл пройден БЕЗ ОШИБОК', str(now))
                    print()
                elif result[-1] == 'Отменен':
                    # print('Базовый документ отменен')
                    ex[tnpas][tnpa][4] = 'Отменен'
                    print(cock, tnpa, ex[tnpas][tnpa])
                    now = datetime.datetime.now()
                    print('*****цикл пройден БЕЗ ОШИБОК', str(now))
                    print()
                else:
                    # print('Действующий')
                    ex[tnpas][tnpa][4] = 'Не заменен'
                    # print('версия проверки без учета анализа изменений')
                    # print(ex[tnpas][tnpa])
                    which_doc(i, tnpas, tnpa)
                    # print('указываем скорректированную версию')
                    print(cock, tnpa, ex[tnpas][tnpa])
                    now = datetime.datetime.now()
                    print('*****цикл пройден БЕЗ ОШИБОК', str(now))
                    print()

            except StaleElementReferenceException:
                result = vvod_ips(i)
                # print(result)
                if result[-1] == 'Недействующий НД':
                    # print('Не действующий')
                    basis_code(result[7], tnpas, tnpa)
                    print(cock, ex[tnpas][tnpa])
                    print('цикл')
                    now = datetime.datetime.now()
                    print(str(now))
                    print()
                elif result[-1] == 'Отменен':
                    # print('Базовый документ отменен')
                    ex[tnpas][tnpa][4] = 'Отменен'
                    print(cock, ex[tnpas][tnpa])
                    print('цикл')
                    now = datetime.datetime.now()
                    print(str(now))
                    print()
                else:
                    # print('Действующий')
                    ex[tnpas][tnpa][4] = 'Не заменен'
                    # print('версия проверки без учета анализа изменений')
                    # print(ex[tnpas][tnpa])
                    which_doc(i, tnpas, tnpa)
                    # print('указываем скорректированную версию')
                    print(cock, ex[tnpas][tnpa])
                    print('цикл')
                    now = datetime.datetime.now()
                    print(str(now))
                    print()

            cock += 1




# print(ex)
posl = 1
new_ex = []
for t in ex:
    for n in ex[t]:
        new_ex.append([posl, t, n, ex[t][n][0], ex[t][n][1], ', '.join(ex[t][n][2]), ', '.join(ex[t][n][3]), ex[t][n][4], ', '.join(ex[t][n][5]), ', '.join(ex[t][n][6]), ', '.join(ex[t][n][7])])
        posl += 1
# print(new_ex)
df = pd.DataFrame(new_ex)
df.to_excel('./test.xlsx')
