from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from time import sleep
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
import logging
import time
import os
import threading
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter import messagebox


class App:
    def __init__(self, tk):
        current_time = time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime(time.time()))
        log_file_name = current_time + '.txt'
        logging.basicConfig(level=logging.DEBUG,
                            format='%(asctime)s [line:%(lineno)d] %(levelname)s %(message)s',
                            datefmt='%a, %d %b %Y %H:%M:%S',
                            filename=log_file_name,
                            filemode='w')

        self.mouse = None
        self.dir_name = ""
        self.dir_path = ""
        self.root = tk

        self.lb1 = LabelFrame(self.root, text="任务参数设置", padx=2, pady=2)
        self.lb1.grid(row=1, column=1, padx=10, pady=5, sticky=EW)

        self.mode_label = Label(self.lb1, text="模式选择：", padx=2, pady=2)
        self.mode_label.grid(row=1, column=1, padx=2, pady=2, sticky=W)
        self.mode_var = IntVar()
        self.new_mode_radiobutton = Radiobutton(self.lb1, text="全新模式", value=1, variable=self.mode_var)
        self.new_mode_radiobutton.grid(row=1, column=2, padx=10, pady=2, sticky=W)
        self.new_mode_radiobutton = Radiobutton(self.lb1, text="继续模式", value=2, variable=self.mode_var)
        self.new_mode_radiobutton.grid(row=1, column=3, padx=10, pady=2, sticky=W)

        self.dir_label = Label(self.lb1, text="输出目录：", padx=2, pady=2)
        self.dir_label.grid(row=2, column=1, padx=2, pady=2, sticky=W)
        self.dir_name_var = StringVar()
        self.dir_entry = Entry(self.lb1, textvariable=self.dir_name_var)
        self.dir_entry.grid(row=2, column=2, columnspan=2, padx=2, pady=2, sticky=EW)

        self.login_option = Label(self.lb1, text="登录选项：", padx=2, pady=2)
        self.login_option.grid(row=3, column=1, padx=2, pady=2, sticky=W)
        self.is_login_vpn_var = BooleanVar()
        self.login_vpn_checkbutton = Checkbutton(self.lb1, text="登录VPN", onvalue=True, offvalue=False,
                                                 variable=self.is_login_vpn_var, padx=2, pady=2)
        self.login_vpn_checkbutton.grid(row=3, column=2, padx=2, pady=2, sticky=W)

        self.lb2 = LabelFrame(self.root, text="任务执行消息", padx=2, pady=2)
        self.lb2.grid(row=2, column=1, padx=10, pady=5, sticky=EW)
        self.execute_msg_frame = ScrolledText(self.lb2, width=80, height=20)
        self.execute_msg_frame.grid(row=1, column=1, padx=2, pady=2)

        self.lb3 = LabelFrame(self.root, text="任务操作", padx=2, pady=2)
        self.lb3.grid(row=3, column=1, padx=10, pady=8, sticky=EW)
        self.start_button = Button(self.lb3, text="开始", width=10, height=1, padx=2, pady=2, command=self.start)
        self.start_button.grid(row=1, column=1, padx=2, pady=2, sticky=E)
        self.stop_button = Button(self.lb3, text="停止", width=10, height=1, padx=2, pady=2, command=self.stop)
        self.stop_button.grid(row=1, column=2, padx=2, pady=2, sticky=E)

        self.root.mainloop()

    def start(self):
        log_v("click start")
        log_v("dir:" + self.dir_name_var.get() + " , " + str(self.mode_var.get()))
        self.execute()

    def stop(self):
        log_v("click stop")

    def execute(self):
        if self.mode_var.get() == 0:
            messagebox.showerror('错误', '请选择模式！\n全新模式表示自动建立新目录进行下载。\n继续模式则从指定目录继续下载')
            return

        current_dir = os.getcwd()
        current_time = time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime(time.time()))
        self.dir_name = current_time
        if self.mode_var.get() == 2:
            if self.dir_name_var.get():
                if os.path.exists(os.path.join(current_dir, self.dir_name_var.get())):
                    self.dir_name = self.dir_name_var.get()
                else:
                    messagebox.showerror('错误', '指定的目录不存在。')
                    return

        # 创建目录
        self.dir_path = os.path.join(current_dir, self.dir_name)
        if not os.path.exists(self.dir_path):
            os.mkdir(self.dir_path)

        threading.Thread(target=self.do_execute).start()
        log_v("execute end")

    def do_execute(self):
        self.mouse = Mouse(self)
        self.mouse.execute()

    def log_info(self, msg):
        log_v(msg)
        current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        self.execute_msg_frame.insert(END, current_time + "  " + msg + '\n')
        self.execute_msg_frame.see(END)


def log_v(msg):
    logging.info(msg)
    current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    print(current_time + " : " + msg)


class Mouse:
    def __init__(self, app_obj):
        self.app = app_obj
        self.dir_name = app_obj.dir_name
        self.dir_path = app_obj.dir_path

        options = webdriver.ChromeOptions()
        # prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': './'}
        # options.add_experimental_option('prefs', prefs)

        options.add_experimental_option('prefs', {
            "profile.default_content_settings.popups": 0,
            "download.default_directory": self.dir_path,
            "download.prompt_for_download": False,  # 如果改为FALSE，则不会弹出下载确认框。To auto download the file
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True  # 如果为TRUE,则不会显示预览。It will not show PDF directly in chrome
        })

        self.driver = webdriver.Chrome(options=options)
        # 常量时间
        self.sleep_time = 1
        self.current_window = None
        # 业务变量
        self.current_bill_order_number = 0
        self.current_bill_limit_of_current_page = 100000
        # 账单数量的变量
        self.total_amount_of_bill_processed = 0
        self.amount_of_bill_downloaded_successfully = 0
        self.amount_of_bill_skipped_for_downloaded = 0
        self.amount_of_bill_skipped_for_no_info = 0
        self.amount_of_bill_skipped_for_rejected = 0

        self.current_bill_page_number = 0
        # 重试点击打印次数的变量
        self.time_for_clicking_print_button = 0
        # 业务阶段变量
        self.is_working = True
        self.is_current_bill_rejected = False
        # 业务文件变量
        self.current_bill_company_name = ""
        self.current_bill_enter_time = ""
        self.current_bill_main_number = ""
        self.current_file_name = ""
        self.current_file_path = ""
        self.prefix = "bill"
        # 初始化常量
        self.init_const()
        self.state = self.INIT_STATE

    def init_const(self):
        self.INIT_STATE = 0
        self.LOGIN_VPN = "LOGIN_VPN"
        self.LOGIN_GS = "LOGIN_GS"
        # 财务共享
        self.CLICK_FINANCE_SHARE_ITEM = "CLICK_FINANCE_SHARE"
        # 资金结算平台
        self.CLICK_FUND_SETTLEMENT_PLATFORM = "CLICK_FUND_SETTLEMENT_PLATFORM"
        # 单位结算办理
        self.CLICK_ORG_SETTLEMENT_MANAGEMENT = "CLICK_ORG_SETTLEMENT_MANAGEMENT"
        # 切换单位结算办理的frame
        self.TRANSFER_FRAME_ORG_SETTLEMENT_MANAGEMENT = 9
        # 点击已办
        self.CLICK_HANDLED = 10
        # 点击下一页，如果有必要
        self.CLICK_NEXT_PAGE_IF_NEEDED = 11
        # 查看需要打印的报账单元信息，如第几页第几条等
        self.CHECK_BILL_METADATA = 12
        # 获取勾选的报账单的信息，包括公司名，入系统时间，单据编号等
        self.CHECK_ONE_BILL_INFO = 13
        # 检查是否当前报账单是否已经下载，若已下载则跳过
        self.CHECK_IS_CURRENT_BILL_HAS_DOWNLOADED = 14
        # 点击报销单单选框
        self.SELECT_ONE_BILL_CHECKBOX = 15
        # 点击查看报销单
        self.CLICK_CHECK_BILL = "CLICK_CHECK_BILL"
        # 切换报销单frame
        self.TRANSFER_BILL_FRAME = "TRANSFER_BILL_FRAME"
        # 点击打印
        self.CLICK_PRINT_BUTTON = "CLICK_PRINT_BUTTON"
        # 切换稽核frame
        self.TRANSFER_FRAME_AUDIT_BILL = "TRANSFER_FRAME_AUDIT_BILL"
        # 点击稽核
        self.CLICK_AUDIT_BILL = "CLICK_AUDIT_BILL"
        # 点击下载确定按钮
        self.CLICK_DOWNLOAD_CONFIRM_BUTTON = 30

    def click_close_current_tab(self):
        current_tab_close_icon_locator = "//div[@id='mainTab']//li[@class='tabs-selected']/a[@class='tabs-close']"
        self.app.log_info("准备关闭当前标签")
        if self.wait_visible_by_xpath(current_tab_close_icon_locator):
            self.click_by_xpath(current_tab_close_icon_locator)
            self.app.log_info("成功关闭当前标签")

    def click_all_function_menu(self):
        self.driver.switch_to.default_content()
        all_function_menu_locator = "//div[@id='mainTab']//span[text()='所有功能']/../.."
        if self.wait_visible_by_xpath(all_function_menu_locator):
            self.click_by_xpath(all_function_menu_locator)

    def execute(self):
        if self.app.is_login_vpn_var.get():
            base_url = 'https://123.127.6.34/'
            self.driver.get(base_url)
            # 跳过不安全验证
            self.skip_verification()
            # 登录VPN
            self.login_vpn()
        else:
            base_url = 'http://10.224.0.59/cwbase/web/Login.aspx'
            self.driver.get(base_url)
        sleep(self.sleep_time)

        # 5.登录GS
        self.login_gs()
        # 6.点击财务共享
        self.click_finance_share()
        # 7.点击资金结算平台
        self.click_fund_settlement_platform()
        # 8.点击单位结算办理
        self.click_org_settlement_management()

        while self.is_working:
            # 9.切换单位结算办理的frame
            if self.state < self.TRANSFER_FRAME_ORG_SETTLEMENT_MANAGEMENT:
                self.switch_frame_to_org_settlement_manager()
            # 10.点击已办
            if self.state < self.CLICK_HANDLED:
                if self.click_handled():
                    continue
            # 11. 如有必要，点击下一页.
            # 在末尾下载时，或者此时，可能改变序号值，其余情况不改变。
            # todo,获取页码总数（和后面重复），如果到达最后一页，最有一条，则结束
            self.click_next_page_if_need()
            # 查看需要打印的报账单元信息，如第几页第几条等
            self.check_bill_metadata()
            # 获取当前的报账单的信息，包括公司名，入系统时间，单据编号等
            self.check_one_bill_info()
            # 检查是否当前报账单是否需要下载，若已下载或者是有异常的则跳过
            if self.check_is_need_skip_current_bill(self.current_file_path, self.current_file_name):
                # 在此处返回，实际是返回11步
                continue
            # 勾选之前确定的某一个报账单表格
            self.select_one_bill_checkbox()
            # 点击查看报账单
            self.click_check_bill()
            # 切到查看报账单的frame
            self.switch_frame_to_check_bill()
            # 点击打印按钮
            if self.click_print_button():
                # 在此处返回，实际是返回11步，所以要由于要关闭当前标签，所以要切回主页面，还要执行9步
                continue
            # 点击借款单稽核
            self.click_audit_bill()
            # 清空之前多余文件
            self.delete_pre_file(self.dir_path)
            # 点击下载确认按钮
            self.click_download_confirm_button()
            # 重命名
            self.rename_download_file(self.dir_path)
            # 切换回主页面
            self.driver.switch_to.default_content()
            # 关闭当前标签
            self.click_close_current_tab()

    def login_vpn(self):
        log_v("login_vpn begin")
        self.app.log_info("准备登陆VPN")
        self.current_window = self.driver.current_window_handle
        self.driver.find_element_by_id("svpn_name").send_keys("xueshan")
        # self.driver.find_element_by_id("svpn_name").send_keys("xiongli")
        sleep(self.sleep_time)
        self.driver.find_element_by_id("svpn_password").send_keys("XSshasha33!!")
        # self.driver.find_element_by_id("svpn_password").send_keys("Xl7893120!!")
        sleep(self.sleep_time)
        self.driver.find_element_by_id("logButton").click()
        sleep(self.sleep_time)
        self.driver.execute_script("javascript:SinforHandleIpHref('33')")
        log_v("login_vpn end")
        self.app.log_info("登陆VPN成功")
        sleep(3)
        self.switch_next_window()

    def skip_verification(self):
        log_v("skip_verification begin")
        detail_button = self.driver.find_element_by_id("details-button")
        detail_button.click()
        sleep(self.sleep_time)
        self.driver.find_element_by_id("proceed-link").click()
        sleep(self.sleep_time)
        log_v("skip_verification end")

    def login_gs(self):
        # 需要判断是否加载完成，等待
        log_v("login_gs begin")
        self.app.log_info("准备登陆浪潮财务系统")
        self.set_language_chinese()
        if self.wait_visible_by_id("txt_UserID"):
            self.driver.find_element_by_id("txt_UserID").send_keys("xueshan")
            self.driver.find_element_by_id("txt_PptextNo").send_keys("shasha33@")
            self.driver.find_element_by_id("btn-login").click()
        log_v("login_gs end")
        self.app.log_info("成功登陆浪潮财务系统")

    def set_language_chinese(self):
        log_v("set_language_chinese begin")
        if self.wait_visible_by_id("txt_UserID"):
            self.click_by_xpath("//span[@class='combo']")
            if self.wait_visible_by_id("_easyui_combobox_1"):
                self.click_by_id("_easyui_combobox_1")
                log_v("set_language_chinese successfully")
        sleep(self.sleep_time)
        log_v("set_language_chinese end")

    def click_download_confirm_button(self):
        log_v("click_download_confirm_button begin")
        download_tips_dialog = "//div[@class='dialog-button']"
        self.app.log_info("准备点击确定按钮进行下载")
        if self.wait_visible_by_xpath(download_tips_dialog):
            dialog = self.driver.find_element_by_xpath(download_tips_dialog)
            dialog.find_elements_by_xpath("./a")[0].click()
            self.app.log_info("成功点击确定按钮进行下载")
            # 报账单编号自增1
            self.current_bill_order_number += 1
            self.total_amount_of_bill_processed += 1
            self.amount_of_bill_downloaded_successfully += 1
            # 更新状态
            self.state = self.INIT_STATE
            sleep(5)
        log_v("click_download_confirm_button end")

    def click_audit_bill(self):
        log_v("click_audit_bill begin")
        self.app.log_info("准备点击稽核按钮")
        dialog_frame_locator = "//iframe[@id='viewIFrame']"
        self.switch_frame(dialog_frame_locator, False)
        bill_check_left_button_locator = "//div[@id='printDialog_left']/a[substring(text(), string-length(text()) - string-length('稽核') +1) = '稽核']"
        bill_check_right_button_locator = "//div[@id='printDialog_right']/a[substring(text(), string-length(text()) - string-length('稽核') +1) = '稽核']"
        bill_check_button_locator = "//div[@id='printDialog_left']"
        if self.wait_visible_by_xpath(bill_check_right_button_locator, wait_time=5):
            self.click_by_xpath(bill_check_right_button_locator)
            self.app.log_info("成功点击稽核按钮")
        elif self.wait_visible_by_xpath(bill_check_left_button_locator, wait_time=1):
            self.click_by_xpath(bill_check_left_button_locator)
            self.app.log_info("成功点击稽核按钮")
        elif self.wait_visible_by_xpath(bill_check_button_locator, wait_time=1):
            self.click_by_xpath(bill_check_button_locator)
            self.app.log_info("无法找到以稽核结尾的按钮，故点击第一个")
        sleep(4)
        log_v("click_audit_bill end")

    def click_print_button(self):
        log_v("click_print_button begin")
        self.app.log_info("准备点击打印按钮")
        try_click_print_button_time = 1
        while try_click_print_button_time <= 3:
            print_button_locator = "//div[@id='BarPubBill']//span[text()='打印']/../.."
            if self.wait_visible_by_xpath(print_button_locator):
                try:
                    self.click_by_xpath(print_button_locator)
                    self.app.log_info("成功点击打印按钮")
                    log_v("click_print_button end")
                    return False
                except ElementClickInterceptedException as e:
                    log_v('except: ' + str(e))
                    try_click_print_button_time += 1
                    sleep(5)
        log_v("click_print_button end")
        self.app.log_info("点击打印按钮失败，准备重试")
        self.time_for_clicking_print_button += 1
        # 切换回主页面
        self.driver.switch_to.default_content()
        sleep(1)
        # 关闭当前标签
        self.click_close_current_tab()
        # 切回单位结算办理的frame
        self.switch_frame_to_org_settlement_manager()
        return True

    def switch_frame_to_check_bill(self):
        log_v("switch_frame_to_check_bill begin")
        frame_locator = "//div[@id='mainTab']/div[@class='tabs-panels']/div[contains(@style,'block')]//iframe"
        self.switch_frame(frame_locator)
        sleep(7)
        log_v("switch_frame_to_check_bill end")

    def click_check_bill(self):
        log_v("click_check_bill begin")
        self.app.log_info("准备点击查看报销单")
        check_bill_button_locator = "//div[@id='Bar1']//span[text()='查看报账单']/../.."
        self.click_by_xpath(check_bill_button_locator)
        self.app.log_info("成功点击查看报销单")
        sleep(4)
        log_v("click_check_bill end")

    def get_footer_table_tds_web_element(self):
        log_v("get_footer_table_tds_web_element begin")
        res = None
        footer_div_locator = "//div[@id='Layout2_Main']//div[@class='datagrid-pager pagination']"
        if self.wait_visible_by_xpath(footer_div_locator):
            footer_div = self.driver.find_element_by_xpath(footer_div_locator)
            footer_table = footer_div.find_element_by_tag_name('table')
            footer_table_tds = footer_table.find_elements_by_xpath("./tbody/tr/td")
            res = footer_table_tds
        log_v("get_footer_table_tds_web_element end")
        return res

    def check_is_need_skip_current_bill(self, current_file_path, current_file_name):
        """
        检查当前报账单是否已经下载
        :return: 若已经下载则返回TRUE
        """
        log_v("check_is_need_skip_current_bill begin")
        self.app.log_info("检查当前报销单是否需要下载")
        res = False
        if os.path.exists(current_file_path):
            log_v("current_file_path exist")
            self.app.log_info("如下报账单已经下载：" + current_file_name)
            self.increment_bill_number()
            self.amount_of_bill_skipped_for_downloaded += 1
            res = True
        elif self.is_current_bill_rejected:
            log_v("current_file_path rejected")
            self.app.log_info("如下报账单已经退回：" + current_file_name)
            self.is_current_bill_rejected = False
            self.increment_bill_number()
            self.amount_of_bill_skipped_for_rejected += 1
            res = True
        elif self.time_for_clicking_print_button > 3:
            log_v("try to click print button more times")
            self.time_for_clicking_print_button = 0
            self.increment_bill_number()
            self.amount_of_bill_skipped_for_no_info += 1
            self.app.log_info("多次查看如下账单均无信息：" + current_file_name + " ，故跳过该账单. ")
            self.app.log_info("目前由于同样原因，已经跳过数量为: " + str(self.amount_of_bill_skipped_for_no_info))
            res = True
        log_v("check_is_need_skip_current_bill end")
        return res

    def increment_bill_number(self):
        log_v("increment_bill_number begin")
        self.current_bill_order_number += 1
        self.total_amount_of_bill_processed += 1
        log_v("increment_bill_number end")

    def check_bill_metadata(self):
        log_v("check_bill_metadata begin")
        footer_table_tds = self.get_footer_table_tds_web_element()
        if footer_table_tds:
            page_limit = footer_table_tds[8].find_element_by_tag_name('span').text

            footer_div_locator = "//div[@id='Layout2_Main']//div[@class='pagination-info']"
            bill_total_limit = self.driver.find_element_by_xpath(footer_div_locator).get_attribute("innerHTML")

            self.app.log_info("已经进入报销单列表页面。{0}，{1}。目前准备打印第{2}页的{3}条，总共处理了{4}条。".format(page_limit, bill_total_limit,
                                                                                       self.current_bill_page_number + 1,
                                                                                       self.current_bill_order_number + 1,
                                                                                       self.total_amount_of_bill_processed))
            self.app.log_info("当前数量统计。该次执行中成功下载{0}条，跳过已下载的{1}条，跳过无信息的账单{2}条，跳过已退回的{3}条".format(self.amount_of_bill_downloaded_successfully,
                                                                                       self.amount_of_bill_skipped_for_downloaded,
                                                                                       self.amount_of_bill_skipped_for_no_info,
                                                                                       self.amount_of_bill_skipped_for_rejected))
        log_v("check_bill_metadata end")

    def click_next_page_if_need(self):
        log_v("click_next_page_if_need begin")
        self.app.log_info("检查是否有必要点击下一页")
        self.current_bill_limit_of_current_page = len(self.get_checkbox_list_of_bill_number())
        if self.current_bill_order_number < self.current_bill_limit_of_current_page:
            log_v("click_next_page_if_need end. It's no need to click next page.")
            self.app.log_info("当前报销单还不是最后一个，无需点击下一页")
            return
        footer_table_tds = self.get_footer_table_tds_web_element()
        if footer_table_tds:
            next_page_button = footer_table_tds[10].find_element_by_tag_name('a')
            self.app.log_info("点击下一页")
            next_page_button.click()
            self.current_bill_order_number = 0
            self.current_bill_page_number += 1
            sleep(4)
        log_v("click_next_page_if_need end")

    def get_checkbox_list_of_bill_number(self):
        log_v("select_one_bill_checkbox begin")
        checkbox_table_locator = "//div[@id='Layout2_Main']//div[@class='datagrid-view1']/div[@class='datagrid-body']//table[@class='datagrid-btable']"
        if self.wait_visible_by_xpath(checkbox_table_locator):
            number_table = self.driver.find_element_by_xpath(checkbox_table_locator)
            checkbox_list = number_table.find_elements_by_xpath("./tbody/tr/td[@field='True']//input[@type='checkbox']")
            return checkbox_list
        return None

    def select_one_bill_checkbox(self):
        log_v("select_one_bill_checkbox begin")
        self.app.log_info("准备勾选当前报销单")
        checkbox_list = self.get_checkbox_list_of_bill_number()
        if checkbox_list:
            list_size = len(checkbox_list)
            self.current_bill_limit_of_current_page = list_size
            log_v("checkbox_list size: " + str(list_size))

            # 滑动滚动条到某个指定的元素
            js4 = "arguments[0].scrollIntoView();"
            # 将下拉滑动条滑动到当前div区域
            self.driver.execute_script(js4, checkbox_list[self.current_bill_order_number])
            sleep(1)
            checkbox_list[self.current_bill_order_number].click()
            self.app.log_info("成功勾选当前报销单")
            sleep(3)
        log_v("select_one_bill_checkbox end")

    def click_handled(self):
        log_v("click_handled begin")
        res = True
        self.app.log_info("准备点击已办")
        left_nav_locator = "//div[@id='IFrame1']//span[@class='tabs-title' and text()='已办']/.."
        if self.wait_visible_by_xpath(left_nav_locator):
            self.click_by_xpath(left_nav_locator)
            self.state = self.CLICK_HANDLED
            self.app.log_info("成功点击已办")
            res = False
        else:
            # 切换回主页面
            self.driver.switch_to.default_content()
            sleep(1)
            # 关闭当前标签
            self.click_close_current_tab()
            sleep(1)
            # 点击单位结算办理
            self.click_org_settlement_management()
        sleep(5)  # 强制等待刷新
        log_v("click_handled end. res: " + str(res))
        return res

    def switch_frame_to_org_settlement_manager(self):
        log_v("switch_frame_to_org_settlement_manager begin")
        frame_locator = "//div[@id='mainTab']//iframe"
        self.switch_frame(frame_locator, False)
        sleep(4)
        log_v("switch_frame_to_org_settlement_manager end")

    def switch_frame(self, frame_locator, is_switch_default_content=True):
        log_v("switch_frame: " + frame_locator)
        if is_switch_default_content:
            log_v("switch_default_content")
            self.driver.switch_to.default_content()
        if self.wait_visible_by_xpath(frame_locator):
            self.driver.switch_to.frame(self.driver.find_element_by_xpath(frame_locator))

    def click_org_settlement_management(self):
        log_v("click_org_settlement_management begin")
        # 点击单位结算办理
        self.app.log_info("准备点击单位结算办理")
        settle_process_locator = "//div[@title='单位结算办理']"
        if self.wait_visible_by_xpath(settle_process_locator):
            self.click_by_xpath(settle_process_locator)
            self.app.log_info("成功点击单位结算办理")
        log_v("click_org_settlement_management end")

    def click_fund_settlement_platform(self):
        log_v("click_fund_settlement_platform begin")
        # 点击资金结算平台
        self.app.log_info("准备点击资金结算平台")
        if self.wait_visible_by_id("_easyui_tree_32"):
            self.click_by_id("_easyui_tree_32")
            self.app.log_info("成功点击资金结算平台")
        log_v("click_fund_settlement_platform end")

    def click_finance_share(self):
        log_v("click_finance_share begin")
        # 点击财务共享
        self.app.log_info("准备点击财务共享")
        finance_share_locator = "//div[@id='_easyui_tree_29']"
        if self.wait_visible_by_xpath(finance_share_locator):
            finance_share = self.driver.find_element_by_xpath(finance_share_locator)
            ActionChains(self.driver).double_click(finance_share).perform()
            self.app.log_info("成功点击财务共享")
        log_v("click_finance_share end")

    def check_one_bill_info(self):
        log_v("check_one_bill_info begin")
        bills_table_locator = "//div[@id='Layout2_Main']//div[@class='datagrid-view2']//table[@class='datagrid-btable']"
        if self.wait_visible_by_xpath(bills_table_locator):
            bills_table = self.driver.find_element_by_xpath(bills_table_locator)
            specify_tr_element = bills_table.find_element_by_xpath(
                "./tbody/tr[" + str(self.current_bill_order_number + 1) + "]")
            company_td_element = specify_tr_element.find_element_by_xpath("./td[@field='LSTASKS_WFZHMC']")
            self.current_bill_company_name = company_td_element.find_element_by_tag_name('div').get_attribute(
                "innerHTML")
            log_v("check_one_bill_info company: " + self.current_bill_company_name)

            bill_enter_time_td_element = specify_tr_element.find_element_by_xpath("./td[@field='LSTASKS_RCSJ']")
            self.current_bill_enter_time = bill_enter_time_td_element.find_element_by_tag_name('div').get_attribute(
                "innerHTML")
            log_v("check_one_bill_info bill_enter_time: " + self.current_bill_enter_time)

            bill_main_number_td_element = specify_tr_element.find_element_by_xpath(
                "./td[@field='LSTASKS_DJBH']")
            self.current_bill_main_number = bill_main_number_td_element.find_element_by_tag_name('div').get_attribute(
                "innerHTML")
            log_v("check_one_bill_info current_bill_main_number: " + self.current_bill_main_number)
            self.current_file_name = self.prefix + self.current_bill_company_name + self.current_bill_enter_time + self.current_bill_main_number + ".pdf"
            self.current_file_path = os.path.join(self.dir_path, self.current_file_name)

            state_td_element = specify_tr_element.find_element_by_xpath("./td[@field='LSTASKS_ZT']")
            # 判断当前账单是否是退回的账单
            try:
                # 使用xpath,当前节点前，一定要有.
                bill_main_number_td_element.find_element_by_xpath(".//span[text()='退回']")
            except NoSuchElementException as e:
                log_v("current bill is normal")
                try:
                    state_td_element.find_element_by_xpath(".//div[text()='退回']")
                except NoSuchElementException as e:
                    log_v("current bill is normal")
                else:
                    log_v("current bill is rejected. ")
                    self.app.log_info("当前账单是退回的账单:" + self.current_bill_main_number)
                    # 发生了NoSuchElementException异常，说明页面中未找到该元素，返回False
                    self.is_current_bill_rejected = True
            else:
                log_v("current bill is rejected. ")
                self.app.log_info("当前账单是退回的账单:" + self.current_bill_main_number)
                # 发生了NoSuchElementException异常，说明页面中未找到该元素，返回False
                self.is_current_bill_rejected = True
        log_v("check_one_bill_info end")

    def delete_pre_file(self,dir_path):
        log_v("delete_pre_file begin")
        file_list = os.listdir(dir_path)
        self.app.log_info("准备删除多余文件")
        for file_name in file_list:
            file_path = os.path.join(dir_path, file_name)
            if not os.path.isdir(file_path):
                if not file_name.startswith(self.prefix):
                    os.remove(file_path)
                    self.app.log_info("完成删除多余文件:" + file_path)
                    break
        log_v("delete_pre_file end")

    def rename_download_file(self, dir_path):
        log_v("rename_download_file begin")
        file_list = os.listdir(dir_path)
        self.app.log_info("准备重命名文件")
        for file_name in file_list:
            file_path = os.path.join(dir_path, file_name)
            if not os.path.isdir(file_path):
                if not file_name.startswith(self.prefix):
                    os.rename(file_path, self.current_file_path)
                    self.app.log_info("完成下载和重命名:" + self.current_file_name)
                    break
        log_v("rename_download_file end")

    def click_by_xpath(self, settle_process_locator):
        self.driver.find_element_by_xpath(settle_process_locator).click()

    def click_by_id(self, widget_id):
        self.driver.find_element_by_id(widget_id).click()

    def wait_visible_by_id(self, wait_id):
        log_v("wait_visible_by_id:" + wait_id)
        try:
            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.ID, wait_id)))
            log_v("wait_visible_by_id: " + wait_id + " True")
            return True
        except TimeoutException:
            log_v("wait_visible_by_id: " + wait_id + " False")
            return False

    def wait_visible_by_xpath(self, xpath, wait_time=10):
        log_v("wait_visible_by_xpath:" + xpath)
        try:
            WebDriverWait(self.driver, wait_time).until(ec.visibility_of_element_located((By.XPATH, xpath)))
            log_v("wait_visible_by_xpath: " + xpath + " True")
            return True
        except TimeoutException:
            log_v("wait_visible_by_xpath: " + xpath + " False")
            return False

    def wait_presence_by_id(self, wait_id):
        log_v("wait_presence_by_id:" + wait_id)
        try:
            WebDriverWait(self.driver, 10).until(ec.presence_of_element_located((By.ID, wait_id)))
            log_v("wait_presence_by_id: True")
            return True
        except TimeoutException:
            log_v("wait_presence_by_id: False")
            return False

    def wait_presence_by_xpath(self, xpath):
        log_v("wait_presence_by_xpath:" + xpath)
        try:
            WebDriverWait(self.driver, 10).until(ec.presence_of_element_located((By.XPATH, xpath)))
            log_v("wait_presence_by_xpath: " + xpath + " True")
            return True
        except TimeoutException:
            log_v("wait_presence_by_xpath: " + xpath + " False")
            return False

    def switch_next_window(self):
        all_handles = self.driver.window_handles
        log_v("switch_next_window: " + str(len(all_handles)))
        for handle in all_handles:
            if handle != self.current_window:
                self.driver.switch_to.window(handle)
                self.current_window = handle
                break


if __name__ == "__main__":
    root = Tk()
    root.title("财务自动处理软件 - 薛珊~ 宝贝快乐")
    # root.geometry("600x400+200+200")

    app = App(root)
    # app.execute()
