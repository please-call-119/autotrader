# coding=utf-8
# 导入库
import io
import time
import pandas as pd
import pytesseract
from datetime import datetime
from pywinauto.keyboard import send_keys
from pywinauto import Application, clipboard

# 下单程序的路径
xiadan_exe_path = r'D:\长城证券同花顺版\xiadan.exe'

# 交易账户和密码
stock_account = ***
trader_password = ***
# 股票代码，数量，价格（价格同花顺自带，其实不用输入，这里只是方便实验）
stock_number = 162411
price = ***
amount = ***


# 验证码识别，参数：图片，返回字符串
def Ocr(png):
    verification_code = pytesseract.image_to_string(png, config="--psm 6 digits")
    return verification_code


# 清空编辑栏，参数：编辑控件
def empty_edit(edit_win):
    for clean_count in range(6):
        edit_win.type_keys('{BS}')
    time.sleep(1)


# 定义ThsTarder类
class ThsTarder:
    # 传入xiadan.exe的路径,股票账号，交易密码
    def __init__(self, exe_path, account, password):
        # 尝试连接已经存在的交易程序
        try:
            self.app = Application(backend="uia").connect(path=exe_path)
            # 查看顶级窗口标题
            win_name = self.app.top_window().set_focus().texts()[0]
            print(win_name)
            # 如果交易窗口存在
            if win_name == '网上股票交易系统V6.9':
                print('交易应用已存在')
                # 连接交易窗口
                self.main_window = self.app.window(title=u"网上股票交易系统V6.9", control_type="Window")
            # 如果存在的窗口不是交易窗口，则为登入窗口
            else:
                print('登录应用已存在，执行登录')
                # 执行登入程序，传递账户和密码
                self.login_exe(account, password)
            # self.app.window(title='网上股票交易系统V6.9').exists(timeout=10)

        # 如果以上连接已有窗口的操作都失败，则执行打开窗口
        except Exception as ex:
            print(ex)
            print('没有交易窗口在运行，启动应用')
            # 执行打开窗口操作，传递程序路径，账户，密码
            self.start_exe(exe_path, account, password)

    # 打开下单程序，传递程序路径，账户，密码
    def start_exe(self, exe_path, account, password):
        # 尝试打开下单程序，最多循环十次
        for open_count in range(10):
            # 尝试打开下单程序
            try:
                # 连接下单程序，给与10秒的缓存时间，如果成功连接，就开始登录步骤，如果没有成功就显示异常
                self.app = Application(backend="uia").start(exe_path, timeout=10)
                # 判断窗口是否存在
                is_bing = bool(self.app)
                # 如果存在，显示打开成功
                if is_bing:
                    print('打开下单程序成功，开始执行登入步骤')
                    # 执行登入操作
                    self.login_exe(account, password)
                # 跳出循环
                break
            # 如果以上操作失败，显示失败原因和尝试次数，执行下个循环
            except Exception as ex:
                is_bing = open_count + 1
                print('异常步骤：打开下单程序，请检查程序路径和文件完整性', ex)
                print('第[%s]次打开应用失败，3秒后将再尝试' % is_bing)
                time.sleep(3)

    # 定义登录函数，传入账号和密码
    def login_exe(self, account, password):
        login_count = 1
        # 循环登入
        for password_count in range(4):
            # self.win_name = self.app.top_window().set_focus().texts()
            print('第[%s]次登入中' % login_count)
            # 连接登入窗口
            login_windows = self.app['用户登录']
            login_windows.set_focus()
            login_windows.draw_outline(colour='green')
            # 定位账户编辑框控件
            account_win = login_windows.children()[0]
            account_win.draw_outline(colour='green')
            # 输入账号
            account_win.type_keys(account)
            time.sleep(0.5)
            # 定位密码编辑框控件
            password_win = login_windows.children()[1]
            password_win.draw_outline(colour='green')
            # 清空密码编辑框
            for login_clean_count in range(6):
                password_win.type_keys('{BS}')
            time.sleep(1)
            # 输入密码
            password_win.type_keys(password)
            time.sleep(0.5)
            # 定位验证码控件
            auth_code_png_win = login_windows.children()[3]
            auth_code_png_win.draw_outline(colour='green')
            # 调用pywinauto库的截图函数capture_as_image进行验证码截图
            auth_code_png = auth_code_png_win.capture_as_image()
            # 调用Ocr函数识别验证码，传递参数：图片
            auth_code = Ocr(auth_code_png)
            # 定位验证码编辑框
            auth_code_win = login_windows.children()[2]
            # 输入验证码
            auth_code_win.type_keys(auth_code)
            # 定位确定登入控件
            login_yes = login_windows.children()[4]
            login_yes.draw_outline(colour='green')
            # 点击登入
            login_yes.click_input()
            time.sleep(1)
            # 如果网上交易系统窗口存在，就显示登入成功
            if self.app.window(title='网上股票交易系统V6.9').exists(timeout=10):
                print('登入成功')
                open_main_window = self.app.window(title=u"网上股票交易系统V6.9", control_type="Window")
                # 定位第一个子控件
                daxin = open_main_window.children()[0]
                # 取控件信息，存储为字典格式
                daxin_dit = daxin.get_properties()
                # 如果第一个子控件名字为#32770，说明为打新或其他页面，关闭打新页面
                if daxin_dit['class_name'] == '#32770':
                    daxin.close()
                    print('检测到有打新页面，关闭中')
                else:
                    print('没有弹出打新页面，无需关闭')
                self.main_window = self.app.window(title=u"网上股票交易系统V6.9", control_type="Window")
                break
            # 如果没有连接到交易主窗口
            else:
                print('登入失败')
                # 连接到报错窗口，可能是验证码等失败
                error_win = login_windows.children()[0]
                error_win.draw_outline(colour='red')
                # 连接确定控件，点击确定
                yes_win = error_win.children()[0]
                yes_win.click_input()
                login_count = login_count + 1  # 循环计数器
                # 第二种判断方法
                # if self.app.top_window().set_focus().texts() == '网上股票交易系统V6.9':
                #     print('登入成功')
                #     break
                # else:
                #     # print(self.login_windows.children_texts())
                #     self.login_windows = self.app['用户登录']
                #     error_win = self.login_windows.children()[0]
                #     error_win.draw_outline(colour='red')
                #     yes_win=error_win.children()[0]
                #     yes_win.click_input()
                #     time.sleep(3)

    # 连接左侧菜单，返回左侧菜单控件
    def left_menu_win(self):
        left_menu_win = self.main_window.child_window(auto_id="129", control_type="Tree")
        left_menu_win().set_focus()
        # self.left_menu_win=self.main_window.children()[0].children()[0].children()[1].children()[0].children()[0]
        # print(left_menu_win.children_texts())
        print('连接左侧菜单成功')
        return left_menu_win

    # 连接右侧菜单，返回右侧菜单控件
    def right_data_win(self):
        right_menu_win = self.main_window.child_window(title="HexinScrollWnd", auto_id="1047", control_type="Pane")
        right_menu_win.set_focus()
        print('连接右侧菜单成功')
        return right_menu_win

    # 买入程序
    def buy(self, buy_stock_number, buy_price, buy_amount):
        # 复制临时主窗口控件，便于操作
        temporary_main_win = self.main_window
        temporary_main_win.set_focus()
        # 定位左侧买入按钮控件，并点击
        left_buy_ctrl = temporary_main_win.child_window(title="买入[F1]", control_type="TreeItem")
        left_buy_ctrl.click_input()
        # print(temporary_main_win.children_texts())
        # 定位证券代码编辑框
        stocks_num_win = temporary_main_win.child_window(auto_id="1032", control_type="Edit")
        # self.buy_win=self.temporary_main_win.children()[0].children()[1]
        stocks_num_win.draw_outline(colour='red')
        # 清空证券代码编辑框
        empty_edit(stocks_num_win)
        # 输入代码
        stocks_num_win.type_keys(buy_stock_number)
        time.sleep(1)
        # self.buy_win=self.temporary_main_win.children()[0].children()[3]
        # 定位价格窗口，其实可以不要，因为同花顺自己会输入当前价格
        price_win = temporary_main_win.child_window(auto_id="1033", control_type="Edit")
        # 输入价格
        price_win.type_keys(buy_price)
        time.sleep(1)
        # self.buy_win=self.temporary_main_win.children()[0].children()[5]
        # 定位数量窗口
        amount_win = temporary_main_win.child_window(auto_id="1034", control_type="Edit")
        # 输入数量
        amount_win.type_keys(buy_amount)
        time.sleep(1)
        # 定位买入控件
        buy_key = temporary_main_win.child_window(title="买入[B]", auto_id="1006", control_type="Button")
        buy_key.draw_outline(colour='red')
        # 点击买入
        buy_key.click_input()
        # 点击确定
        self.yes_win()
        time.sleep(1)
        now_time = datetime.now()
        # 点击确定
        self.yes_win()
        new = '通知：在[%s]时委托下单，以[%s]买入[%s]股票[%s]手' % (now_time, buy_price, buy_stock_number, buy_amount)
        print(new)

    # 点击确定
    def yes_win(self):
        # 定位第一个子控件的第一个控件，也就是yes按钮
        yes_win_yes = self.main_window.children()[0].children()[0]
        # 点击确定
        yes_win_yes.click_input()

    # 点击否
    def no_win(self):
        # 定位第一个子控件的第二个控件，也就是否按钮
        no_win_yes = self.main_window.children()[0].children()[1]
        # 点击否
        no_win_yes.click_input()

    # 查看今日交易
    def selet_deal_today(self):
        # self.selet_ctrl = self.left_menu_win().children()[5]
        # 定位查询控件
        selet_ctrl = self.main_window.child_window(title="查询[F4]", control_type="TreeItem")
        selet_ctrl.set_focus()
        # 点击查询
        selet_ctrl.click_input()
        # self.deal_today_ctrl=self.selet_ctrl.children()[1]
        # 定位当日成交控件
        deal_today_ctrl = selet_ctrl.child_window(title="当日成交", control_type="TreeItem")
        deal_today_ctrl.click_input()
        deal_today_ctrl.draw_outline(colour='red')
        # 调用右侧菜单，获得右侧界面控件
        deal_today_win = self.right_data_win()
        deal_today_win.draw_outline(colour='red')
        deal_today_win.set_focus()
        # 进行ctrl+v操作
        send_keys('^c')
        # 验证码循环，最多尝试10次
        for authentication_count in range(10):
            # auth_code_png_win=self.main_window.children()[0].children()[3]
            # 定位弹出验证码的窗口
            auth_code_png_win = self.main_window.child_window(title="检测到您正在拷贝数据，为保护您的账号数据安全，请", auto_id="2405",
                                                              control_type="Image")
            # 定位验证码图片控件
            auth_code_png_win.draw_outline(colour='red')
            # 截图
            auth_code_png = auth_code_png_win.capture_as_image()
            # 调用Ocr识别验证码
            auth_code = Ocr(auth_code_png)
            # auth_code_win=self.main_window.children()[0].children()[5]
            # 定位验证码编辑框
            auth_code_win = self.main_window.child_window(title="提示", auto_id="2404", control_type="Edit")
            auth_code_win.draw_outline(colour='red')
            # 输入验证码
            auth_code_win.type_keys(auth_code)
            # 点击确定
            self.yes_win()
            time.sleep(1)
            # 如果验证码查询窗口不存在了，则说明查询成功
            if self.main_window.children()[0].get_properties()['class_name'] != '#32770':
                print('查询验证码成功')
                # 从剪切板复制信息到data
                data = clipboard.GetData()
                # 以dataframe数据格式存储信息
                df = pd.read_csv(io.StringIO(data), delimiter='\t', na_filter=False)
                print(type(df))
                print(df)
                break
            # 否则验证码输入错误，验证码弹窗依旧存在
            else:
                print('第[%s]次输入验证码失败' % authentication_count)
                # 清空验证码，进行下次循环
                empty_edit(auth_code_win)

    # 查看资金
    def selet_money(self):
        # self.selet_ctrl = self.left_menu_win().children()[5]
        # 定位查询控件
        selet_ctrl = self.main_window.child_window(title="查询[F4]", control_type="TreeItem")
        selet_ctrl.set_focus()
        # 点击查询
        selet_ctrl.click_input()
        # self.deal_today_ctrl=self.selet_ctrl.children()[1]
        # 定位资金股票控件
        selet_money_ctrl = selet_ctrl.child_window(title="资金股票", control_type="TreeItem")
        # 点击确定
        selet_money_ctrl.click_input()
        selet_money_ctrl.draw_outline(colour='red')
        # 调用右侧菜单，获得右侧界面控件
        selet_money_win = self.right_data_win()
        selet_money_win.draw_outline(colour='red')
        selet_money_win.set_focus()
        # 进行ctrl+v操作
        send_keys('^c')
        # 定位弹出验证码的窗口
        for authentication_count in range(10):
            # auth_code_png_win=self.main_window.children()[0].children()[3]
            # 定位验证码图片控件
            auth_code_png_win = self.main_window.child_window(title="检测到您正在拷贝数据，为保护您的账号数据安全，请", auto_id="2405",
                                                              control_type="Image")
            auth_code_png_win.draw_outline(colour='red')
            # 截图
            auth_code_png = auth_code_png_win.capture_as_image()
            # 调用Ocr识别验证码
            auth_code = Ocr(auth_code_png)
            # auth_code_win=self.main_window.children()[0].children()[5]
            # 定位验证码编辑框
            auth_code_win = self.main_window.child_window(title="提示", auto_id="2404", control_type="Edit")
            auth_code_win.draw_outline(colour='red')
            # 输入验证码
            auth_code_win.type_keys(auth_code)
            # 点击确定
            self.yes_win()
            time.sleep(1)
            # 如果验证码查询窗口不存在了，则说明查询成功
            if self.main_window.children()[0].get_properties()['class_name'] != '#32770':
                print('查询验证码成功')
                # 从剪切板复制信息到data
                data = clipboard.GetData()
                # 以dataframe数据格式存储信息
                df = pd.read_csv(io.StringIO(data), delimiter='\t', na_filter=False)
                print(df['证券名称'])
                break
            # 否则验证码输入错误，验证码弹窗依旧存在
            else:
                print('第[%s]次输入验证码失败' % authentication_count)
                # 清空验证码，进行下次循环
                empty_edit(auth_code_win)

    # 买入操作，传递参数：代码，价格，数量
    def sell(self, sell_stock_number, sell_price, sell_amount):
        # 复制临时主窗口控件，便于操作
        temporary_main_win = self.main_window
        temporary_main_win.set_focus()
        # 定位左侧买入控件
        left_sell_ctrl = temporary_main_win.child_window(title="卖出[F2]", control_type="TreeItem")
        # 点击确定
        left_sell_ctrl.click_input()
        # 定位证券代码编辑框
        stocks_num_win = temporary_main_win.child_window(auto_id="1032", control_type="Edit")
        # self.buy_win=self.temporary_main_win.children()[0].children()[1]
        stocks_num_win.draw_outline(colour='red')
        # 清空证券代码编辑框
        empty_edit(stocks_num_win)
        # 输入证券代码
        stocks_num_win.type_keys(sell_stock_number)
        time.sleep(1)
        # self.buy_win=self.temporary_main_win.children()[0].children()[3]
        # 定位价格编辑框控件
        price_win = temporary_main_win.child_window(auto_id="1033", control_type="Edit")
        # 输入价格
        price_win.type_keys(sell_price)
        time.sleep(1)
        # self.buy_win=self.temporary_main_win.children()[0].children()[5]
        # 定位数量编辑框控件
        amount_win = temporary_main_win.child_window(auto_id="1034", control_type="Edit")
        # 输入数量
        amount_win.type_keys(sell_amount)
        time.sleep(1)
        # 定位卖出控件
        sell_key = temporary_main_win.child_window(title="卖出[S]", auto_id="1006", control_type="Button")
        sell_key.draw_outline(colour='red')
        # 点击卖出
        sell_key.click_input()
        # 点击确定
        self.yes_win()
        time.sleep(1)
        # 点击确定
        self.yes_win()
        now_time = datetime.now()
        new = '通知：在[%s]时委托下单，以[%s]卖出[%s]股票[%s]手' % (now_time, sell_price, sell_stock_number, sell_amount)
        print(new)

    # 市价买入，传递参数：股票代码，数量
    def market_buy(self, market_buy_stock_number, market_buy_amount):
        # 复制临时主窗口控件，便于操作
        temporary_main_win = self.main_window
        temporary_main_win.set_focus()
        # 定位左侧市价委托控件
        left_sell_ctrl = temporary_main_win.child_window(title="市价委托", control_type="TreeItem")
        left_sell_ctrl.draw_outline(colour='red')
        # 点击市价委托
        left_sell_ctrl.click_input()
        # 定位买入控件
        market_buy_ctrl = self.main_window.child_window(title="买入", control_type="TreeItem")
        market_buy_ctrl.draw_outline(colour='red')
        # 点击买入
        market_buy_ctrl.click_input()
        # 定位证券代码编辑框
        stocks_num_win = temporary_main_win.child_window(auto_id="1032", control_type="Edit")
        time.sleep(1)
        # 输入证券代码
        stocks_num_win.type_keys(market_buy_stock_number)
        # 定位数量编辑框
        amount_win = temporary_main_win.child_window(auto_id="1034", control_type="Edit")
        # 输入数量
        amount_win.type_keys(market_buy_amount)
        # 定位买入按钮
        market_buy_yes = self.main_window.child_window(title="买入[B]", auto_id="1006", control_type="Button")
        # 确定买入
        market_buy_yes.click_input()
        # 点击确定
        self.yes_win()
        time.sleep(1)
        # 点击确定
        self.yes_win()
        now_time = datetime.now()
        new = '通知：在[%s]时委托下单，以市价买入[%s]股票[%s]手' % (now_time, market_buy_stock_number, market_buy_amount)
        print(new)

    # 市价卖出
    def market_sell(self, market_sell_stock_number, market__sell_amount):
        # 复制临时主窗口控件，便于操作
        temporary_main_win = self.main_window
        temporary_main_win.set_focus()
        # 定位左侧市价委托控件
        left_sell_ctrl = temporary_main_win.child_window(title="市价委托", control_type="TreeItem")
        left_sell_ctrl.draw_outline(colour='red')
        # 点击市价委托
        left_sell_ctrl.click_input()
        # 定位卖出控件
        market_sell_ctrl = self.main_window.child_window(title="卖出", control_type="TreeItem")
        market_sell_ctrl.draw_outline(colour='red')
        # 点击卖出
        market_sell_ctrl.click_input()
        # 定位证券代码编辑框
        stocks_num_win = temporary_main_win.child_window(auto_id="1032", control_type="Edit")
        time.sleep(1)
        # 输入证券代码
        stocks_num_win.type_keys(market_sell_stock_number)
        # 定位数量编辑框
        amount_win = temporary_main_win.child_window(auto_id="1034", control_type="Edit")
        # 输入数量
        amount_win.type_keys(market__sell_amount)
        # 定位卖出按钮
        market_buy_yes = self.main_window.child_window(title="卖出[S]", auto_id="1006", control_type="Button")
        # 点击卖出
        market_buy_yes.click_input()
        # 确定
        self.yes_win()
        time.sleep(1)
        # 确定
        self.yes_win()
        now_time = datetime.now()
        new = '通知：在[%s]时委托下单，以市价卖出[%s]股票[%s]手' % (now_time, market_sell_stock_number, market__sell_amount)
        print(new)


# 实例化ThsTarder
a = ThsTarder(xiadan_exe_path, stock_account, trader_password)
# 执行买入操作
a.buy(stock_number, price, amount)
# 执行卖出操作
a.sell(stock_number, price, amount)
#  执行市价买入
a.market_buy(stock_number, amount)
# 市价卖出
a.market_sell(stock_number, amount)
# 查看今天交易
a.selet_deal_today()
# 查看资金明细
a.selet_money()
