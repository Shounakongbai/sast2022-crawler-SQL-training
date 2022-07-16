import time

# from requests import options
from selenium.webdriver.remote.webdriver import WebDriver as wd
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait as wdw
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains as AC
import selenium
import selenium.webdriver
from bs4 import BeautifulSoup as BS
import json
from IPython import embed


class WebVPN:

    def __init__(self, opt: dict, headless=False):
        self.root_handle = None
        self.driver: wd = None
        self.userid = opt["username"]
        self.passwd = opt["password"]
        self.headless = headless

    def login_webvpn(self):
        """
        Log in to WebVPN with the account specified in `self.userid` and `self.passwd`

        :return:
        """
        d = self.driver
        if d is not None:
            d.close()

        my_options = selenium.webdriver.ChromeOptions()
        if self.headless:
            my_options.add_argument("--headless")
        # options.binary_location = "/mnt/c/Program Files/Google/Chrome/Application"

        d = selenium.webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        d.get("https://webvpn.tsinghua.edu.cn/login")
        username = d.find_elements(By.XPATH,
                                   '//div[@class="login-form-item"]//input'
                                   )[0]
        password = d.find_elements(By.XPATH,
                                   '//div[@class="login-form-item password-field" and not(@id="captcha-wrap")]//input'
                                   )[0]

        username.send_keys(str(self.userid))
        password.send_keys(self.passwd)
        d.find_element(By.ID, "login").click()
        self.root_handle = d.current_window_handle
        self.driver = d
        return d

    def access(self, url_input):
        """
        Jump to the target URL in WebVPN

        :param url_input: target URL
        :return:
        """
        d = self.driver
        url = By.ID, "quick-access-input"
        btn = By.ID, "go"
        wdw(d, 5).until(EC.visibility_of_element_located(url))
        actions = AC(d)
        actions.move_to_element(d.find_element(*url))
        actions.click()
        actions. \
            key_down(Keys.CONTROL). \
            send_keys("A"). \
            key_up(Keys.CONTROL). \
            send_keys(Keys.DELETE). \
            perform()

        d.find_element(*url)
        d.find_element(*url).send_keys(url_input)
        d.find_element(*btn).click()

    def switch_another(self):
        """
        If there are only 2 windows handles, switch to the other one

        :return:
        """
        d = self.driver
        assert len(d.window_handles) == 2
        wdw(d, 5).until(EC.number_of_windows_to_be(2))
        for window_handle in d.window_handles:
            if window_handle != d.current_window_handle:
                d.switch_to.window(window_handle)
                return

    def to_root(self):
        """
        Switch to the home page of WebVPN

        :return:
        """
        self.driver.switch_to.window(self.root_handle)

    def close_all(self):
        """
        Close all window handles

        :return:
        """
        while True:
            try:
                l = len(self.driver.window_handles)
                if l == 0:
                    break
            except selenium.common.exceptions.InvalidSessionIdException:
                return
            self.driver.switch_to.window(self.driver.window_handles[0])
            self.driver.close()

    def login_info(self):
        """
        TODO: After successfully logged into WebVPN, login to info.tsinghua.edu.cn

        :return: None
        """

        # Hint: - Use `access` method to jump to info.tsinghua.edu.cn
        #       - Use `switch_another` method to change the window handle
        #       - Wait until the elements are ready, then preform your actions
        #       - Before return, make sure that you have logged in successfully
        self.login_webvpn()
        self.access("https://info.tsinghua.edu.cn")
        self.switch_another()
        d = self.driver

        userName_input = By.ID, "userName"
        password_input = By.NAME, "password"
        form = d.find_elements(By.XPATH, '//*[@id="login-form"]')[0]
        wdw(d, 5).until(EC.visibility_of_element_located(password_input))

        d.find_element(*userName_input)
        d.find_element(*userName_input).send_keys(self.userid)
        d.find_element(*password_input)
        d.find_element(*password_input).send_keys(self.passwd)
        form.submit()

        '''确认已经登录好了'''
        chengji = By.XPATH, '//*[@id="menu"]/li[2]/a[10]'
        wdw(d, 5).until(EC.visibility_of_element_located(chengji))

        self.driver = d

    def get_grades(self):
        """
        TODO: Get and calculate the GPA for each semester.

        Example return / print:
            2020-秋: *.**
            2021-春: *.**
            2021-夏: *.**
            2021-秋: *.**
            2022-春: *.**

        :return:
        """

        # Hint: - You can directly switch into
        #         `zhjw.cic.tsinghua.edu.cn/cj.cjCjbAll.do?m=bks_cjdcx&cjdlx=zw`
        #         after logged in
        #       - You can use Beautiful Soup to parse the HTML content or use
        #         XPath directly to get the contents
        #       - You can use `element.get_attribute("innerHTML")` to get its
        #         HTML code
        d = self.driver

        '''去成绩单页面'''
        self.to_root()
        self.access('zhjw.cic.tsinghua.edu.cn/cj.cjCjbAll.do?m=bks_cjdcx&cjdlx=zw')
        windows = d.window_handles
        d.switch_to.window(windows[2])
        time.sleep(2)

        '''分析成绩单（审判）'''
        chengji_dict = {}
        semesters = set()
        chengji_table = d.find_elements(By.XPATH, '//*[@id="table1"]/tbody')[0]
        chengji_table_html = chengji_table.get_attribute("innerHTML")
        soup = BS(chengji_table_html, 'lxml')
        trs = soup.find_all('tr')
        for i in range(1, len(trs)):
            tds = trs[i].find_all('td')
            # embed()
            semester = tds[5].string.strip()
            semesters.add(semester)
            if tds[4].string.strip() == 'N/A':
                continue  # 跳过N/A
            try:
                chengji_dict[semester + "学分"] += int(tds[2].string.strip())
            except Exception:
                chengji_dict[semester + "学分"] = int(tds[2].string.strip())
            # embed()
            try:
                chengji_dict[semester + "总分"] += int(tds[2].string.strip()) * float(tds[4].string.strip())
            except Exception:
                chengji_dict[semester + "总分"] = int(tds[2].string.strip()) * float(tds[4].string.strip())
        for semester in semesters:
            chengji_dict[semester + " GPA"] = round(chengji_dict[semester + "总分"] / chengji_dict[semester + "学分"], 3)

        '''打印成绩单'''
        for semester in semesters:
            print("您" + semester + "的GPA是： " + str(chengji_dict[semester + " GPA"]))

        self.driver = d


if __name__ == "__main__":
    # TODO: Write your own query process
    with open("./mysettings.json") as f:
        settings = json.load(f)
    vpn = WebVPN(settings, True)

    vpn.login_info()
    vpn.get_grades()
    vpn.close_all()
