import os, sys
from pathlib import Path
from selenium import webdriver
from bs4 import BeautifulSoup
from time import sleep

# https://sites.google.com/chromium.org/driver/
# window의 경우 win32 버전 zip 파일 다운로드
# selenium에서 사용할 웹 드라이버 절대 경로 정보
 
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

def relative_path(data_file_name:str)->Path:
    ''' pyinstaller에 의한 실행일 경우 주소를 pyinstaller에서 생성한 주소인
    sys._MEIPASS를 포함한 주소로 변경
    '''
    # pyinstaller에 의한 one file 실행일 경우
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        bundle_dir = Path(sys._MEIPASS)
    else:
        try:
            bundle_dir = Path(__file__).parent
        except NameError:
            bundle_dir = os.getcwd()

    print(str(bundle_dir))

    path_to_dat = Path.cwd() / bundle_dir / data_file_name
    return path_to_dat

def get_cwd(data_file_name:str)->Path:
    ''' pyinstaller에 의한 실행일 경우 주소를 pyinstaller에서 생성한 주소인
    sys._MEIPASS를 포함한 주소로 변경
    '''
    # pyinstaller에 의한 one file 실행일 경우
    return Path.cwd() / data_file_name
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        os.getcwd()
    else:
        try:
            bundle_dir = Path(__file__).parent
        except NameError:
            bundle_dir = os.getcwd()

    path_to_dat = Path.cwd() / bundle_dir / data_file_name
    return path_to_dat



def main()->None:
    
    options = webdriver.ChromeOptions()
    options.add_argument('window-size=1920,1080')
    # options.add_argument('headless')

    chromedriver = 'chromedriver.exe'
    driver = webdriver.Chrome(get_cwd(chromedriver), options = options)
    kepco_url = 'https://en-ter.co.kr/main.do'
    driver.get(kepco_url)

    login = driver.find_element_by_css_selector(
        "#header > div.header_box > div > a:nth-child(1)")
    login.click()

    sleep(3)

    is_enterprise = driver.find_element_by_css_selector(
        '#login_tab01 > div.form_group > label:nth-child(4)')
    is_enterprise.click()    

    id_text = driver.find_element_by_css_selector(
        '#id')
    id_text.send_keys('######')


    pw_text = driver.find_element_by_css_selector(
        '#pwd')
    pw_text.send_keys('#######')

    log_bn = driver.find_element_by_css_selector(
        '#btn_login')
    log_bn.click()

    sleep(3)
    # id 변경과 관련하여 대기 필요
    driver.implicitly_wait(3)
    change_later = driver.find_element_by_css_selector(
        '#content > div > div.pw_info > div > a.btn.btn_cancel')
    change_later.click()

    sleep(3)
    energy_data = driver.find_element_by_css_selector(
        '#depth1_2 > a')
    energy_data.click()    

    driver.implicitly_wait(3)
    sleep(10)
    biz_portal = driver.find_element_by_css_selector(
        '#enterArea > div.content > div > div.tab_content > div > ul > li:nth-child(3) > a > span.thum_info > strong')
        # '#depth1_2 > div > div > div > ul > li:nth-child(3) > a')
    biz_portal.click()


    sleep(3)
    # tap이 변경되어 2번째 tap으로 이동 필요
    driver.switch_to.window(driver.window_handles[1])

    # driver.implicitly_wait(3)
    # html_source_code = driver.execute_script("return document.body.innerHTML;")
    # html_soup = BeautifulSoup(html_source_code, 'html.parser')
    # filename = 'source.txt'
    # with open(filename, mode="w",  encoding="utf8") as code:
    #         code.write(str(html_soup.prettify()))

    # print(html_soup)

    # 고객정보제공 동의현황
    driver.execute_script("Go_SubmenuTop(2,'000227','000063','Y');")

    # 로딩에 시간이 많이 걸림
    driver.implicitly_wait(30)
    list_of_service = Select(driver.find_element_by_css_selector(
        '#bodyArea > div.search_box > fieldset > div > div:nth-child(1) > span:nth-child(1) > select'))
    # list_of_service.select_by_value('에너지사용량신고 전력데이터')
    # 에너지사용량신고 전력데이터, index 2
    list_of_service.select_by_index(2)

    search_btn = driver.find_element_by_css_selector(
        '#bodyArea > div.search_box > fieldset > div > span > a')
    search_btn.click()
    sleep(3)

    download_excel_btn = driver.find_element_by_css_selector(
        '#printBtn')
    download_excel_btn.click()

    # 검색이 완료될 때까지 대기
    # driver.implicitly_wait(130)
    sleep(60*3)
    print('job done')

    # frames = [frame.get_attribute('id') for frame in driver.find_elements_by_tag_name('iframe')]

    # WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH,"//iframe[@id='successIframe' and @name='successIframe']")))

    #이동할 프레임 엘리먼트 지정
    # iframe = driver.find_element_by_id("Cal_iFrame")    
    # driver.switch_to.frame(iframe)

if __name__=='__main__':
    main()
