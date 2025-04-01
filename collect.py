from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException
# Chrome driver 자동 업데이트
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import openpyxl
import os

service = Service(executable_path=ChromeDriverManager().install())

class 양도세:
    해외 = "2"
    주식 = "61"
    양도 = "10"
    취득 = "01"

class 공통:
    @staticmethod
    def print_child_elements(driver, parent_locator, parent_type='xpath'):
        # 부모 요소 찾기
        if parent_type == 'xpath':
            parent_element = driver.find_element(By.XPATH, parent_locator)
        elif parent_type == 'css':
            parent_element = driver.find_element(By.CSS_SELECTOR, parent_locator)
        elif parent_type == 'id':
            parent_element = driver.find_element(By.ID, parent_locator)
        elif parent_type == 'class_name':
            parent_element = driver.find_element(By.CLASS_NAME, parent_locator)
        else:
            raise ValueError(f"지원하지 않는 locator 타입입니다: {parent_type}")

        # 자식 요소들 찾기
        child_elements = parent_element.find_elements(By.XPATH, './*')

        # 자식 요소들 출력하기
        for index, child in enumerate(child_elements):
            print(f"Child {index + 1}: {child.tag_name}, {child.text}")

    @staticmethod
    def print_table_selectors(driver, selector_type='xpath'):
        tables = driver.find_elements(By.TAG_NAME, "table")
        # 테이블의 XPath 또는 CSS Selector 출력하기
        for index, table in enumerate(tables):
            if selector_type == 'xpath':
                xpath = driver.execute_script(
                    "function absoluteXPath(element) {"
                    "var comp, comps = [];"
                    "var parent = null;"
                    "var xpath = '';"
                    "var getPos = function(element) {"
                    "var position = 1, curNode;"
                    "if (element.nodeType == Node.ATTRIBUTE_NODE) {"
                    "return null;"
                    "}"
                    "for (curNode = element.previousSibling; curNode; curNode = curNode.previousSibling){"
                    "if (curNode.nodeName == element.nodeName) {"
                    "++position;"
                    "}"
                    "}"
                    "return position;"
                    "};"
                    "if (element instanceof Document) {"
                    "return '/';"
                    "}"
                    "for (; element && !(element instanceof Document); element = element.nodeType == Node.ATTRIBUTE_NODE ? element.ownerElement : element.parentNode) {"
                    "comp = comps[comps.length] = {};"
                    "switch (element.nodeType) {"
                    "case Node.TEXT_NODE:"
                    "comp.name = 'text()';"
                    "break;"
                    "case Node.ATTRIBUTE_NODE:"
                    "comp.name = '@' + element.nodeName;"
                    "break;"
                    "case Node.PROCESSING_INSTRUCTION_NODE:"
                    "comp.name = 'processing-instruction()';"
                    "break;"
                    "case Node.COMMENT_NODE:"
                    "comp.name = 'comment()';"
                    "break;"
                    "case Node.ELEMENT_NODE:"
                    "comp.name = element.nodeName;"
                    "break;"
                    "}"
                    "comp.position = getPos(element);"
                    "}"
                    "for (var i = comps.length - 1; i >= 0; i--) {"
                    "comp = comps[i];"
                    "xpath += '/' + comp.name.toLowerCase();"
                    "if (comp.position !== null) {"
                    "xpath += '[' + comp.position + ']';"
                    "}"
                    "}"
                    "return xpath;"
                    "} return absoluteXPath(arguments[0]);", table)
                print(f"Table {index + 1} XPath: {xpath}")
            elif selector_type == 'css':
                # 테이블의 CSS Selector 생성
                classes = table.get_attribute("class").split()
                ids = table.get_attribute("id")
                if ids:
                    selector = f"#{ids}"
                elif classes:
                    selector = "." + ".".join(classes)
                else:
                    tag_name = table.tag_name
                    nth_child = driver.execute_script(
                        "var index = 0;"
                        "var el = arguments[0];"
                        "while ((el = el.previousElementSibling) != null) index++;"
                        "return index;", table)
                    selector = f"{tag_name}:nth-of-type({nth_child + 1})"
                print(f"Table {index + 1} CSS Selector: {selector}")
            else:
                print(f"지원하지 않는 selector 타입입니다: {selector_type}")


class 한국투자증권(양도세):
    번호 = "116-81-04504"
    COLLECT_PAGE = "https://securities.koreainvestment.com/main/banking/service/OverseaEvidence.jsp"
    def __init__(self, sec):
        self.sec = sec


    def __enter__(self):
        self.driver = webdriver.Chrome(service=service)
        driver = self.driver
        driver.get("https://securities.koreainvestment.com/main/member/login/login.jsp")
        time.sleep(self.sec)
        return self


    def is_login(self):
        driver = self.driver
        buttons = driver.find_elements(By.CSS_SELECTOR,".btn_logout")
        return buttons

    def go_collect_page(self):
        driver = self.driver
        driver.get(self.COLLECT_PAGE)

    def is_password(self):
        current_url = self.driver.current_url
        assert current_url == self.COLLECT_PAGE, f"URL이 일치하지 않습니다. 현재 URL은 {current_url}입니다."
        driver = self.driver
        pwd_field = driver.find_element(By.CSS_SELECTOR,"#I_PWD")
        pwd_value = driver.execute_script("return arguments[0].value;", pwd_field)
        print("pwd_value:", pwd_value)
        return pwd_value

    def collect(self):
        current_url = self.driver.current_url
        assert current_url == self.COLLECT_PAGE, f"URL이 일치하지 않습니다. 현재 URL은 {current_url}입니다."
        driver = self.driver
        driver.find_element(By.CSS_SELECTOR,".btnWhite").click()
        time.sleep(2)
        button = driver.find_element(By.CSS_SELECTOR,"#btnNext")
        #if not button.is_displayed():
        #    return pd.DataFrame(columns=[i for i in range(23)])
        while button.is_displayed():
            button.click()
            time.sleep(2)
        공통.print_table_selectors(driver, 'css')
        # 데이터 퍼오기
        data = []
        table = driver.find_element(By.CSS_SELECTOR,".CI-GRID-BODY-TABLE")
        for y, tr in enumerate(table.find_elements(By.TAG_NAME,"tr")):
            actions = ActionChains(driver)
            actions.move_to_element(tr).perform()
            td = tr.find_elements(By.TAG_NAME,"td")
            if y >= 2 and y % 2 == 0:
                row = [ "" for _ in range(23) ]
                row[1] = self.번호
                row[2] = self.해외
                row[3] = 0
                row[4] = self.주식
                row[5] = self.주식
                row[6] = self.양도
                row[7] = self.취득
                row[11] = '2000-01-01'
            for x, value in enumerate(td):
                #print("({},{}) - {}".format(y, x, td[x].text))
                if y >= 2:
                    if y % 2 == 0 and x == 2:
                        row[0] = td[x].text
                    elif y % 2 == 0 and x == 1:
                        row[8] = td[x].text.replace(".", "-")
                    elif y % 2 == 0 and x == 5:
                        row[9] = int(float(td[x].text.replace(",", "")))
                    elif y % 2 == 0 and x == 6:
                        row[10] = int(td[x].text.replace(",", ""))
                    elif y % 2 == 1 and x == 0:
                        row[12] = int(float(td[x].text.replace(",", "")))
                    elif y % 2 == 1 and x == 1:
                        row[13] = int(td[x].text.replace(",", ""))
                    elif y % 2 == 0 and x == 8:
                        row[14] = int(td[x].text.replace(",", ""))
                    elif y % 2 == 0 and x == 3:
                        row[20] = td[x].text
                        row[21] = td[x].text[:2]
                    elif y % 2 == 0 and x == 4:
                        row[3] = float(td[x].text)
            if y >= 2 and y % 2 == 1:
                print(row)
                data.append(row)
        if data:
            df = pd.DataFrame(data)
        else:
            df = pd.DataFrame(columns=[i for i in range(23)])
        return df

    def quit(self):
        self.driver.quit()

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.quit()



class 삼성증권(양도세):
    번호 = "202-81-48588"
    COLLECT_PAGE = "https://www.samsungpop.com/ux/kor/trading/overseasStock/overseasStockTransaction/overseasStockTransactionTax.do"
    def __init__(self, sec):
        self.sec = sec


    def __enter__(self):
        self.driver = webdriver.Chrome(service=service)
        driver = self.driver
        driver.get("https://www.samsungpop.com/")
        time.sleep(self.sec)
        driver.implicitly_wait(1)
        login_url = "https://www.samsungpop.com/login/login.pop"
        driver.execute_script(f"document.querySelector('#frmContent').src = '{login_url}';")
        driver.implicitly_wait(1)
        time.sleep(self.sec)
        return self


    def is_login(self):
        driver = self.driver
        frame = driver.find_element(By.ID, 'frmContent')
        driver.switch_to.frame(frame)
        try:
            buttons = driver.find_elements(By.CSS_SELECTOR,"a.logout")
        except Exception:
            wait = WebDriverWait(driver, 1)
            buttons = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.logout")))
        driver.switch_to.default_content()
        return buttons

    def go_collect_page(self):
        driver = self.driver
        driver.execute_script(f"document.querySelector('#frmContent').src = '{self.COLLECT_PAGE}';")

    def is_password(self):
        driver = self.driver
        frame = driver.find_element(By.ID, 'frmContent')
        frame_url = frame.get_attribute('src')
        assert frame_url == self.COLLECT_PAGE, f"URL이 일치하지 않습니다. 현재 URL은 {current_url}입니다."
        return True

    def collect(self):
        driver = self.driver
        frame = driver.find_element(By.ID, 'frmContent')
        frame_url = frame.get_attribute('src')
        assert frame_url == self.COLLECT_PAGE, f"URL이 일치하지 않습니다. 현재 URL은 {current_url}입니다."
        driver.switch_to.frame(frame)
        driver.find_element(By.CSS_SELECTOR,".btnLarge.gray").click()
        time.sleep(2)
        공통.print_table_selectors(driver, 'css')
        # 데이터 퍼오기
        data = []
        table = driver.find_element(By.CSS_SELECTOR,"#foreign_tb2")
        for y, tr in enumerate(table.find_elements(By.TAG_NAME,"tr")):
            actions = ActionChains(driver)
            actions.move_to_element(tr).perform()
            td = tr.find_elements(By.TAG_NAME,"td")
            if y >= 2 and y % 2 == 0:
                row = [ "" for _ in range(23) ]
                row[1] = self.번호
                row[2] = self.해외
                row[3] = 0
                row[4] = self.주식
                row[5] = self.주식
                row[6] = self.양도
                row[7] = self.취득
                row[11] = '2000-01-01'
            for x, value in enumerate(td):
                print("({},{}) - {}".format(y, x, td[x].text))
                if y >= 2:
                    if y % 2 == 0 and x == 0:
                        if td[x].text == "조회 내역이 없습니다.":
                            break
                        #종목명, ISIN코드
                        row[0], row[20] = td[x].text.split('\n')
                        # 국가코드
                        row[21] = row[20][:2]
                    elif y % 2 == 1 and x == 2:
                        # 양도일자
                        row[8] = td[x].text.replace("/", "-")
                    elif y % 2 == 0 and x == 4:
                        # 주당양도가액(int)
                        row[9] = int(float(td[x].text.replace(",", "")))
                    elif y % 2 == 1 and x == 3:
                        # 양도가액(int)
                        row[10] = int(td[x].text.replace(",", ""))
                    elif y % 2 == 0 and x == 5:
                        row[12] = int(float(td[x].text.replace(",", "")))
                    elif y % 2 == 1 and x == 4:
                        row[13] = int(td[x].text.replace(",", ""))
                    elif y % 2 == 0 and x == 6:
                        row[14] = int(td[x].text.replace(",", ""))
            if y >= 2 and y % 2 == 1:
                row[3] = float(row[10]) / float(row[9])
                print(row)
                data.append(row)
        if data:
            df = pd.DataFrame(data)
        else:
            df = pd.DataFrame(columns=[i for i in range(23)])
        return df


    def quit(self):
        self.driver.quit()


    def __exit__(self, exc_type, exc_val, exc_tb):
        self.quit()


class 키움증권(양도세):
    번호 = "107-81-76756"
    COLLECT_PAGE = "https://www1.kiwoom.com/h/banking/certificate/VFostockTransTax4View"
    def __init__(self, sec):
        self.sec = sec


    def __enter__(self):
        options = Options()
        options.add_argument("--disable-blink-features=AutomationControlled")  # 자동화 탐지 방지
        self.driver = webdriver.Chrome(service=service, options=options)
        driver = self.driver
        driver.get("https://www1.kiwoom.com/h/loginView")
        time.sleep(self.sec)
        return self


    def is_login(self):
        driver = self.driver
        divs = driver.find_elements(By.CSS_SELECTOR,".logout-area")
        return divs

    # 다음 함수까지 구현 나머지는 확인 필요
    def go_collect_page(self):
        driver = self.driver
        driver.get(self.COLLECT_PAGE)

    def is_password(self):
        current_url = self.driver.current_url
        assert current_url == self.COLLECT_PAGE, f"URL이 일치하지 않습니다. 현재 URL은 {current_url}입니다."
        driver = self.driver
        pwd_field = driver.find_element(By.CSS_SELECTOR,"#pswd")
        pwd_value = driver.execute_script("return arguments[0].value;", pwd_field)
        print("pwd_value:", pwd_value)
        return pwd_value

    def collect(self):
        current_url = self.driver.current_url
        assert current_url == self.COLLECT_PAGE, f"URL이 일치하지 않습니다. 현재 URL은 {current_url}입니다."
        driver = self.driver
        nation_dic = {
            '홍콩':'HK',
            '미국':'US',
            '일본':'JP'
        }
        driver.find_element(By.CSS_SELECTOR,"#btn_search").click()
        time.sleep(2)
        modals = driver.find_elements(By.CSS_SELECTOR,".modal-body-container")
        if len(modals) == 1:
            return pd.DataFrame(columns=[i for i in range(23)])
        # 공통.print_table_selectors(driver, 'css')
        # 데이터 퍼오기
        data = []
        table = driver.find_element(By.CSS_SELECTOR,"#tbl_view")
        tbody = table.find_element(By.TAG_NAME, "tbody")
        rows = tbody.find_elements(By.TAG_NAME, "tr")
        # 행 데이터를 저장할 리스트
        table_data = []
        # `rowspan` 위치를 추적하기 위한 임시 리스트
        rowspan_map = {}

        for row_idx, row in enumerate(rows):
            cells = row.find_elements(By.TAG_NAME, "td")
            actions = ActionChains(driver)
            actions.move_to_element(row).perform()
            row_data = []
            cell_idx = 0

            while cell_idx < len(cells):
                # `rowspan_map`에서 현재 행에 해당하는 병합된 셀이 있는지 확인
                if row_idx in rowspan_map:
                    row_data.append(rowspan_map[row_idx])
                    rowspan_map.pop(row_idx)

                # 현재 셀 데이터 추출
                cell = cells[cell_idx]
                cell_text = cell.text
                row_data.append(cell_text)

                # `rowspan` 및 `colspan` 속성 추출
                rowspan = int(cell.get_attribute("rowspan") or 1)
                colspan = int(cell.get_attribute("colspan") or 1)

                # rowspan이 있는 경우, 해당 셀 데이터를 rowspan만큼 추가
                if rowspan > 1:
                    for i in range(1, rowspan):
                        rowspan_map[row_idx + i] = cell_text

                # colspan이 있는 경우, 해당 열 수만큼 데이터를 중복 삽입
                for _ in range(1, colspan):
                    row_data.append(cell_text)

                cell_idx += 1

            # 현재 행 데이터를 최종 테이블 데이터에 추가
            table_data.append(row_data)


        # 전체 테이블 데이터 출력
        for y, lst in enumerate(table_data):
            #print(lst)
            if y % 2 == 0:
                row = [ "" for _ in range(23) ]
                row[1] = self.번호
                row[2] = self.해외
                # 양도주식수
                row[3] = int(lst[2].replace(",", ""))
                row[4] = self.주식
                row[5] = self.주식
                row[6] = self.양도
                row[7] = self.취득
                # 양도일자
                row[8] = lst[3].replace('.', '-')
                # 양도가액
                row[10] = int(lst[5].replace(",", ""))
                # 주당양도가액
                row[9] = int(row[10] / row[3])
                # 필요경비(int)
                row[14] = int(lst[6].replace(",", ""))
                # ISIN코드
                row[20] = lst[1]
                # 국외자산코드
                row[21] = nation_dic.get(lst[0], "")
            else:
                # 종목명
                row[0] = lst[1]
                # 취득일자
                row[11] = lst[2].replace('.', '-')
                # 취득가액
                row[13] = int(lst[4].replace(",", ""))
                # 주당취득가액
                row[12] = int(row[13] / row[3])
                data.append(row)
        df = pd.DataFrame(data)
        return df


    def quit(self):
        self.driver.quit()


    def __exit__(self, exc_type, exc_val, exc_tb):
        self.quit()


class 신한투자증권(양도세):
    번호 = "116-81-36684"
    COLLECT_PAGE = "https://www.shinhansec.com/siw/trading/foreign-equity/3826/view.do"
    def __init__(self, sec):
        self.sec = sec

    def __enter__(self):
        self.driver = webdriver.Chrome(service=service)
        driver = self.driver
        driver.get("https://www.shinhansec.com/index.html")
        time.sleep(self.sec)
        driver.implicitly_wait(1)
        login_url = "https://www.shinhansec.com/siw/etc/login/view.do"
        driver.execute_script(f"document.querySelector('#mainFrame').src = '{login_url}';")
        driver.implicitly_wait(1)
        time.sleep(self.sec)
        return self


    def is_login(self):
        driver = self.driver
        frame = driver.find_element(By.ID, 'mainFrame')
        driver.switch_to.frame(frame)
        js_var = driver.execute_script("return SIDATA;")
        #print(js_var)
        driver.switch_to.default_content()
        return js_var[1] != ""


    def go_collect_page(self):
        driver = self.driver
        driver.execute_script(f"document.querySelector('#mainFrame').src = '{self.COLLECT_PAGE}';")
        #driver.get(self.COLLECT_PAGE)

    def is_password(self):
        driver = self.driver
        frame = driver.find_element(By.ID, 'mainFrame')
        frame_url = frame.get_attribute('src')
        assert frame_url == self.COLLECT_PAGE, f"URL이 일치하지 않습니다. 현재 URL은 {current_url}입니다."
        driver.switch_to.frame(frame)
        driver.implicitly_wait(1)
        pwd_field = driver.find_element(By.CSS_SELECTOR,"#inq_pw")
        pwd_value = driver.execute_script("return arguments[0].value;", pwd_field)
        print("pwd_value:", pwd_value)
        driver.switch_to.default_content()
        return pwd_value

    def collect(self):
        driver = self.driver
        frame = driver.find_element(By.ID, 'mainFrame')
        frame_url = frame.get_attribute('src')
        assert frame_url == self.COLLECT_PAGE, f"URL이 일치하지 않습니다. 현재 URL은 {current_url}입니다."
        driver.switch_to.frame(frame)
        driver.find_element(By.CSS_SELECTOR,"#main > fieldset > button").click()
        time.sleep(2)
        공통.print_table_selectors(driver, 'css')
        # 데이터 퍼오기
        data = []
        table = driver.find_element(By.CSS_SELECTOR,"#main > table")
        for y, tr in enumerate(table.find_elements(By.TAG_NAME,"tr")):
            actions = ActionChains(driver)
            actions.move_to_element(tr).perform()
            td = tr.find_elements(By.TAG_NAME,"td")
            if y >= 4 and y % 2 == 0:
                row = [ "" for _ in range(23) ]
                row[1] = self.번호
                row[2] = self.해외
                row[3] = 0
                row[4] = self.주식
                row[5] = self.주식
                row[6] = self.양도
                row[7] = self.취득
            for x, value in enumerate(td):
                #print("({},{}) - {}".format(y, x, td[x].text))
                if y >= 4:
                    if y % 2 == 0 and x == 0:
                        if td[x].text == "검색된 내용이 없습니다.":
                            break
                        # 취득일자
                        row[11] = td[x].text.replace('.',"-")
                    elif y % 2 == 0 and x == 3:
                        # ISIN코드
                        row[20] = td[x].text
                        # 국가코드
                        row[21] = row[20][:2]
                    elif y % 2 == 0 and x == 4:
                        # 양도주식수
                        row[3] = int(td[x].text.replace(",", ""))
                    elif y % 2 == 0 and x == 5:
                        # 주당취득가액(int)
                        row[12] = int(td[x].text.replace(",", ""))
                    elif y % 2 == 0 and x == 6:
                        # 취득가액
                        row[13] = int(td[x].text.replace(",", ""))
                    elif y % 2 == 1 and x == 0:
                        # 양도일자
                        row[8] = td[x].text.replace(".", "-")
                    elif y % 2 == 1 and x == 3:
                        # 종목명
                        row[0] = td[x].text
                    elif y % 2 == 1 and x == 4:
                        # 제비용
                        row[14] = int(td[x].text.replace(",", ""))
                    elif y % 2 == 1 and x == 5:
                        # 주당양도가액(int)
                        row[9] = int(td[x].text.replace(",", ""))
                    elif y % 2 == 1 and x == 6:
                        # 양도가액(int)
                        row[10] = int(td[x].text.replace(",", ""))
            if y >= 4 and y % 2 == 1:
                print(row)
                data.append(row)
        if data:
            df = pd.DataFrame(data)
        else:
            df = pd.DataFrame(columns=[i for i in range(23)])
        return df


    def quit(self):
        self.driver.quit()


    def __exit__(self, exc_type, exc_val, exc_tb):
        self.quit()


class 하나증권(양도세):
    번호 = "116-81-05992"
    COLLECT_PAGE = "https://www.hanaw.com/main/bank/online/OB_042000.cmd"
    def __init__(self, sec):
        self.sec = sec

    def __enter__(self):
        download_path = os.getcwd()
        chrome_options = Options()
        prefs = {
            "download.default_directory": download_path,  # 다운로드 경로 설정
            "download.prompt_for_download": False,  # 다운로드 대화상자 비활성화
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        driver = self.driver
        driver.get("https://www.hanaw.com/main/customer/login/index.cmd")
        time.sleep(self.sec)
        return self


    def is_login(self):
        driver = self.driver
        s = driver.find_element(By.CSS_SELECTOR,"ul.memberlink.new-link > li > a > span")
        return s.text == "로그아웃"


    def go_collect_page(self):
        driver = self.driver
        driver.get(self.COLLECT_PAGE)
        time.sleep(2)
        overseas_stock_tab = driver.find_element(By.XPATH, '//ul[@class="tab-selector"]/li/a[contains(text(), "해외주식")]')
        overseas_stock_tab.click()

    def is_password(self):
        current_url = self.driver.current_url
        assert current_url == self.COLLECT_PAGE, f"URL이 일치하지 않습니다. 현재 URL은 {current_url}입니다."
        driver = self.driver
        pwd_field = driver.find_element(By.CSS_SELECTOR,"#b_scrt")
        pwd_value = driver.execute_script("return arguments[0].value;", pwd_field)
        print("pwd_value:", pwd_value)
        return pwd_value

    def collect(self):
        current_url = self.driver.current_url
        assert current_url == self.COLLECT_PAGE, f"URL이 일치하지 않습니다. 현재 URL은 {current_url}입니다."
        driver = self.driver
        time.sleep(2)

        button = driver.find_element(By.CSS_SELECTOR,"#btn1")
        button.click()
        time.sleep(10)
        frame = driver.find_element(By.CSS_SELECTOR,"#ozIframe")
        x = frame.location['x'] + 190
        y = frame.location['y'] + 25
        actions = ActionChains(driver)
        actions.move_by_offset(x, y).click().perform()
        time.sleep(5)
        actions.move_by_offset(385, 385).click().perform()
        time.sleep(3)

        data = []
        with open('noname.txt') as f:
            for line in f.readlines():
                row = [ "" for _ in range(23) ]
                row[1] = self.번호
                row[2] = self.해외
                row[3] = 0
                row[4] = self.주식
                row[5] = self.주식
                row[6] = self.양도
                row[7] = self.취득
                fields = line.split('\t')
                if len(fields) >= 28 and not line.startswith('cust_nm'):
                    #print(len(fields), fields)
                    # 취득일자
                    row[11] = fields[24][:4] + '-' + fields[24][4:6] + '-' + fields[24][6:8]
                    # ISIN코드
                    row[20] = fields[14]
                    # 국가코드
                    row[21] = fields[14][:2]
                    # 양도주식수
                    row[3] = int(fields[18])
                    # 주당취득가액(int)
                    row[12] = int(fields[22])
                    # 취득가액
                    row[13] = int(fields[23])
                    # 양도일자
                    row[8] = fields[12][:4] + '-' + fields[12][4:6] + '-' + fields[12][6:8]
                    # 종목명
                    row[0] = fields[15]
                    # 제비용
                    row[14] = int(fields[26])
                    # 주당양도가액(int)
                    row[9] = int(fields[19])
                    # 양도가액(int)
                    row[10] = int(fields[20])
                    data.append(row)
        os.remove('noname.txt')
        if data:
            df = pd.DataFrame(data)
        else:
            df = pd.DataFrame(columns=[i for i in range(23)])
        return df


    def quit(self):
        self.driver.quit()


    def __exit__(self, exc_type, exc_val, exc_tb):
        self.quit()


if __name__ == "__main__":
    dfs = []
    df_sums = []
    securities = (
        #("한국투자증권", 한국투자증권(0)),
        ("삼성증권", 삼성증권(0)),
        #("키움증권", 키움증권(0)),
        #("신한투자증권", 신한투자증권(0)),
        #("하나증권", 하나증권(0)),
    )
    #a = 한국투자증권(0) # 수동 로그인 시간
    #a = 삼성증권(0) # 수동 로그인 시간
    for name, sec in securities:
        print(name, sec)
        with sec as a:
            while True:
                time.sleep(10)
                if a.is_login():
                    break
                print("로그인이 필요합니다. 현재 로그인 안 된 상태")
            a.go_collect_page()
            while a.is_login():
                time.sleep(10)
                if a.is_password():
                    break
                print("비밀번호 입력이 완료되어야 데이터를 수집할 수 있습니다.")
            while True:
                df = a.collect()
                if isinstance(df, pd.DataFrame):
                    print(df)
                    break
                print("데이터 수집에 실패하였습니다. 10초후 재시도 합니다")
                time.sleep(10)
            dfs.append(df)
            #df = a.collect(10) # 수동 선택시간
            df_sum = df.sum()
            df_sums.append(df_sum)
            #df_sum.to_csv('data_sum.csv', index=False)
            손익 = df_sum.iloc[10] - df_sum.iloc[13]
            제비용 = df_sum.iloc[14]
            차감손익 = 손익-제비용
            print("손익 :", 손익)
            print("제비용 :", 제비용)
            print("차감손익 :", 차감손익)

    df_tax = pd.concat(dfs, ignore_index=True)
    # 기존 엑셀 파일 불러오기 (서식이 적용된 파일)
    file_path = '주식_엑셀업로드_양식.xlsx'
    wb = openpyxl.load_workbook(file_path)
    # 데이터를 추가할 시트 선택
    sheet_name = '자료'  # 데이터를 추가할 시트 이름
    ws = wb[sheet_name]
    for i, row in df_tax.iterrows():
        for j, value in enumerate(row, start=1):
            ws.cell(row=2+i, column=j, value=value)

    file_path = '주식_엑셀업로드_양식_20241130.xlsx'
    wb.save(file_path)
