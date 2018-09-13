import csv
from decimal import Decimal
import os
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QGridLayout, QWidget, QPushButton, QLabel, QInputDialog, QFileDialog, QMessageBox, QCheckBox
from PyQt5.QtGui import QIcon, QFont
import images_pyqt

basepath = os.path.abspath(os.path.dirname(__file__))  # 当前模块文件的根目录


def get_reduced_money(s):
    '''获取每笔交易的使用了优惠券的金额'''
    int_part = s.split('.')[0]
    return int(int_part.replace(int_part[-1], '0'))


def get_combination(reduced_money, c_par_1, c_par_2, c_par_3, c_max_1, c_max_2, c_max_3):
    '''算出各优惠券的数量组合'''
    combination = []
    for i in range(c_max_1+1):
        for j in range(c_max_2+1):
            for k in range(c_max_3+1):
                if c_par_1 * i + c_par_2 * j + c_par_3 * k == reduced_money:
                    combination.append((i, j, k))
    return combination


class Ui_MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.c_par_1 = 0  # 优惠券1的面额
        self.c_max_1 = 0  # 优惠券1最大叠加的张数
        self.c_par_2 = 0  # 优惠券2的面额
        self.c_max_2 = 0  # 优惠券2最大叠加的张数
        self.c_par_3 = 0  # 优惠券3的面额
        self.c_max_3 = 0  # 优惠券3最大叠加的张数
        self.data_csv = None  # 财务明细数据 csv 文件
        self.data = {}  # 操作员账号是key, value是需要的列组成的字典
        self.faild_rows = []  # 程序无法计算出使用了多少张优惠券时，将财务明细数据 csv 文件中的该条记录保存到这里
        self.shops_csv = None  # 店铺名称 csv 文件
        self.shops = {}  # 操作员账号是key, 店铺名称是value
        self.rate = 0.6  # 实收金额的手续费率 %
        self.init_ui()

    def init_ui(self):
        self.setFixedSize(960, 640)  # 窗口大小, 固定值，窗口不能放大
        self.setWindowIcon(QIcon(':/img/logo.ico'))  # 图标
        self.setWindowTitle('王颜公子 - Python3 Process CSV Data')  # 窗口标题
        self.statusBar().showMessage('Madman 2018. © All Rights Reserved  Blog: http://www.madmalls.com')  # 状态栏消息

        # 网格布局, 假设分成 12 列
        self.main_widget = QWidget()  # 创建窗口主部件
        self.main_widget.setObjectName('main_widget')
        self.main_layout = QGridLayout()  # 创建主部件的网格布局
        self.main_widget.setLayout(self.main_layout)  # 设置窗口主部件布局为网格布局
        self.setCentralWidget(self.main_widget)  # 设置窗口主部件

        # 第0行
        self.label_1 = QLabel('1. 请指定优惠券面额及最大叠加张数', self)
        self.label_1.setFont(QFont('微软雅黑', 10))
        self.label_1.setStyleSheet('color: #ff9999')
        self.main_layout.addWidget(self.label_1, 0, 0, 1, 12)
        # 第1行
        self.btn_1 = QPushButton('修改', self)
        self.btn_1.clicked.connect(self.modify_coupon)
        self.main_layout.addWidget(self.btn_1, 1, 0, 1, 2)
        self.label_2 = QLabel('优惠券1面额(元): ', self)
        self.main_layout.addWidget(self.label_2, 1, 3, 1, 3)
        self.label_3 = QLabel('0', self)
        self.label_3.setFont(QFont('微软雅黑', 12))
        self.label_3.setStyleSheet('color: #ff9999')
        self.main_layout.addWidget(self.label_3, 1, 6, 1, 1)
        # 第2行
        self.btn_2 = QPushButton('修改', self)
        self.btn_2.clicked.connect(self.modify_coupon)
        self.main_layout.addWidget(self.btn_2, 2, 0, 1, 2)
        self.label_4 = QLabel('优惠券1最大叠加数(张): ', self)
        self.main_layout.addWidget(self.label_4, 2, 3, 1, 3)
        self.label_5 = QLabel('0', self)
        self.label_5.setFont(QFont('微软雅黑', 12))
        self.label_5.setStyleSheet('color: #ff9999')
        self.main_layout.addWidget(self.label_5, 2, 6, 1, 1)
        # 第3行
        self.btn_3 = QPushButton('修改', self)
        self.btn_3.clicked.connect(self.modify_coupon)
        self.main_layout.addWidget(self.btn_3, 3, 0, 1, 2)
        self.label_6 = QLabel('优惠券2面额(元): ', self)
        self.main_layout.addWidget(self.label_6, 3, 3, 1, 3)
        self.label_7 = QLabel('0', self)
        self.label_7.setFont(QFont('微软雅黑', 12))
        self.label_7.setStyleSheet('color: #ff9999')
        self.main_layout.addWidget(self.label_7, 3, 6, 1, 1)
        # 第4行
        self.btn_4 = QPushButton('修改', self)
        self.btn_4.clicked.connect(self.modify_coupon)
        self.main_layout.addWidget(self.btn_4, 4, 0, 1, 2)
        self.label_8 = QLabel('优惠券2最大叠加数(张): ', self)
        self.main_layout.addWidget(self.label_8, 4, 3, 1, 3)
        self.label_9 = QLabel('0', self)
        self.label_9.setFont(QFont('微软雅黑', 12))
        self.label_9.setStyleSheet('color: #ff9999')
        self.main_layout.addWidget(self.label_9, 4, 6, 1, 1)
        # 第5行
        self.btn_5 = QPushButton('修改', self)
        self.btn_5.clicked.connect(self.modify_coupon)
        self.main_layout.addWidget(self.btn_5, 5, 0, 1, 2)
        self.label_10 = QLabel('优惠券3面额(元): ', self)
        self.main_layout.addWidget(self.label_10, 5, 3, 1, 3)
        self.label_11 = QLabel('0', self)
        self.label_11.setFont(QFont('微软雅黑', 12))
        self.label_11.setStyleSheet('color: #ff9999')
        self.main_layout.addWidget(self.label_11, 5, 6, 1, 1)
        # 第6行
        self.btn_6 = QPushButton('修改', self)
        self.btn_6.clicked.connect(self.modify_coupon)
        self.main_layout.addWidget(self.btn_6, 6, 0, 1, 2)
        self.label_12 = QLabel('优惠券3最大叠加数(张): ', self)
        self.main_layout.addWidget(self.label_12, 6, 3, 1, 3)
        self.label_13 = QLabel('0', self)
        self.label_13.setFont(QFont('微软雅黑', 12))
        self.label_13.setStyleSheet('color: #ff9999')
        self.main_layout.addWidget(self.label_13, 6, 6, 1, 1)

        # 第7行
        self.label_14 = QLabel('2. 请选择快收后台导出的财务明细数据 csv 文件', self)
        self.label_14.setFont(QFont('微软雅黑', 10))
        self.label_14.setStyleSheet('color: #ff9999')
        self.main_layout.addWidget(self.label_14, 7, 0, 1, 12)
        # 第8行
        self.btn_7 = QPushButton('选择', self)
        self.btn_7.clicked.connect(self.chose_data_csv)
        self.main_layout.addWidget(self.btn_7, 8, 0, 1, 2)
        self.label_15 = QLabel('', self)  # 用来显示选择的文件名
        self.main_layout.addWidget(self.label_15, 8, 3, 1, 11)

        # 第9行
        self.label_16 = QLabel('3. (可选) 请选择包含店铺名称与操作员账号对应关系的 csv 文件', self)
        self.label_16.setFont(QFont('微软雅黑', 10))
        self.label_16.setStyleSheet('color: #ff9999')
        self.main_layout.addWidget(self.label_16, 9, 0, 1, 12)
        # 第10行
        self.btn_8 = QPushButton('选择', self)
        self.btn_8.clicked.connect(self.chose_shops_csv)
        self.main_layout.addWidget(self.btn_8, 10, 0, 1, 2)
        self.label_17 = QLabel('', self)  # 用来显示选择的文件名
        self.main_layout.addWidget(self.label_17, 10, 3, 1, 11)

        # 第11行
        self.label_18 = QLabel('4. 请勾选统计报表的列', self)
        self.label_18.setFont(QFont('微软雅黑', 10))
        self.label_18.setStyleSheet('color: #ff9999')
        self.main_layout.addWidget(self.label_18, 11, 0, 1, 12)
        # 第12行
        self.cb_1 = QCheckBox('操作员账号', self)
        self.cb_1.toggle()
        self.main_layout.addWidget(self.cb_1, 12, 0, 1, 3)
        self.cb_2 = QCheckBox('店铺名称', self)
        self.main_layout.addWidget(self.cb_2, 12, 3, 1, 3)
        self.cb_3 = QCheckBox('交易笔数', self)
        self.cb_3.toggle()
        self.main_layout.addWidget(self.cb_3, 12, 6, 1, 3)
        self.cb_4 = QCheckBox('优惠笔数', self)
        self.cb_4.toggle()
        self.main_layout.addWidget(self.cb_4, 12, 9, 1, 3)
        # 第13行
        self.cb_5 = QCheckBox('优惠金额(元)', self)
        self.cb_5.toggle()
        self.main_layout.addWidget(self.cb_5, 13, 0, 1, 3)
        self.cb_6 = QCheckBox('优惠券1的笔数(0元)', self)
        self.main_layout.addWidget(self.cb_6, 13, 3, 1, 3)
        self.cb_7 = QCheckBox('优惠券2的笔数(0元)', self)
        self.main_layout.addWidget(self.cb_7, 13, 6, 1, 3)
        self.cb_8 = QCheckBox('优惠券3的笔数(0元)', self)
        self.main_layout.addWidget(self.cb_8, 13, 9, 1, 3)
        # 第14行
        self.cb_9 = QCheckBox('实收金额(元)', self)
        self.cb_9.toggle()
        self.main_layout.addWidget(self.cb_9, 14, 0, 1, 3)
        self.cb_10 = QCheckBox('退款笔数', self)
        self.cb_10.toggle()
        self.main_layout.addWidget(self.cb_10, 14, 3, 1, 3)
        self.cb_11 = QCheckBox('退款总额(元)', self)
        self.cb_11.toggle()
        self.main_layout.addWidget(self.cb_11, 14, 6, 1, 3)

        # 第15行
        self.label_19 = QLabel('5. 请指定实收金额手续费率', self)
        self.label_19.setFont(QFont('微软雅黑', 10))
        self.label_19.setStyleSheet('color: #ff9999')
        self.main_layout.addWidget(self.label_19, 15, 0, 1, 12)
        # 第16行
        self.btn_9 = QPushButton('修改', self)
        self.btn_9.clicked.connect(self.modify_rate)
        self.main_layout.addWidget(self.btn_9, 16, 0, 1, 2)
        self.label_20 = QLabel('实收金额手续费率(%): ', self)
        self.main_layout.addWidget(self.label_20, 16, 3, 1, 3)
        self.label_21 = QLabel('0.6', self)
        self.label_21.setFont(QFont('微软雅黑', 12))
        self.label_21.setStyleSheet('color: #ff9999')
        self.main_layout.addWidget(self.label_21, 16, 6, 1, 1)

        # 第17行
        self.label_22 = QLabel('6. 导出最终统计报表', self)
        self.label_22.setFont(QFont('微软雅黑', 10))
        self.label_22.setStyleSheet('color: #ff9999')
        self.main_layout.addWidget(self.label_22, 17, 0, 1, 12)
        # 第18行
        self.btn_10 = QPushButton('导出', self)
        self.btn_10.clicked.connect(self.output_reslut)
        self.main_layout.addWidget(self.btn_10, 18, 0, 1, 2)

        self.show()

    def modify_coupon(self):
        sender = self.sender()
        if sender == self.btn_1:
            text, ok = QInputDialog.getInt(self, '修改优惠券1面额', '请输入面额：', min=10, step=10)
            if ok:
                self.c_par_1 = text  # 保存面额
                self.label_3.setText(str(text))
                self.c_max_1 = 1
                self.label_5.setText('1')  # 自动显示初始张数为1
                self.cb_6.setChecked(True)  # 勾选导出的对应列
                self.cb_6.setText('优惠券1的笔数({}元)'.format(text))  # 修改列的名称
        elif sender == self.btn_2:
            text, ok = QInputDialog.getInt(self, '修改优惠券1最大叠加张数', '请输入最大叠加张数：', min=1)
            if ok:
                self.c_max_1 = text  # 保存最大叠加张数
                self.label_5.setText(str(text))
        elif sender == self.btn_3:
            text, ok = QInputDialog.getInt(self, '修改优惠券2面额', '请输入面额：', min=10, step=10)
            if ok:
                self.c_par_2 = text  # 保存面额
                self.label_7.setText(str(text))
                self.c_max_2 = 1
                self.label_9.setText('1')  # 自动显示初始张数为1
                self.cb_7.setChecked(True)  # 勾选导出的对应列
                self.cb_7.setText('优惠券2的笔数({}元)'.format(text))  # 修改列的名称
        elif sender == self.btn_4:
            text, ok = QInputDialog.getInt(self, '修改优惠券2最大叠加张数', '请输入最大叠加张数：', min=1)
            if ok:
                self.c_max_2 = text  # 保存最大叠加张数
                self.label_9.setText(str(text))
        elif sender == self.btn_5:
            text, ok = QInputDialog.getInt(self, '修改优惠券3面额', '请输入面额：', min=10, step=10)
            if ok:
                self.c_par_3 = text  # 保存面额
                self.label_11.setText(str(text))
                self.c_max_3 = 1
                self.label_13.setText('1')  # 自动显示初始张数为1
                self.cb_8.setChecked(True)  # 勾选导出的对应列
                self.cb_8.setText('优惠券3的笔数({}元)'.format(text))  # 修改列的名称
        elif sender == self.btn_6:
            text, ok = QInputDialog.getInt(self, '修改优惠券3最大叠加张数', '请输入最大叠加张数：', min=1)
            if ok:
                self.c_max_3 = text  # 保存最大叠加张数
                self.label_13.setText(str(text))

        # 如果是导出几次后，重新修改优惠券信息，需要重新触发 get_data() 方法
        if self.data_csv:  # 如果没选择 财务明细数据 csv，则条件为False
            self.get_data(self.data_csv)

    def chose_data_csv(self):
        filename, filetype = QFileDialog.getOpenFileName(self, '选择文件', basepath)
        if filename:
            self.data_csv = filename
            self.get_data(self.data_csv)

    def get_data(self, filename):
        self.data = {}  # 每次调用该函数时，都需要先清空之前的数据
        self.faild_rows = []  # 每次调用该函数时，都需要先清空之前的数据
        try:
            with open(filename) as f:
                f_csv = csv.DictReader(f)
                for row in f_csv:  # row表示每一行数据，每一列都是str类型
                    if not row['操作员账号']:  # csv文件最后一行，是汇总数据，没有操作员账号
                        continue

                    operator_id = row['操作员账号'].strip()
                    if operator_id not in self.data:  # 第一次碰到该商户
                        self.data[operator_id] = {}
                        self.data[operator_id]['shop_name'] = ''  # 该商户的店铺名称
                        self.data[operator_id]['trade_count'] = 0  # 该商户的交易笔数
                        self.data[operator_id]['reduced_count'] = 0  # 该商户的优惠笔数
                        self.data[operator_id]['reduced_succeed_count'] = 0  # 成功计算成优惠券组合的笔数
                        self.data[operator_id]['reduced_faild_count'] = 0  # 未能成功计算成优惠券组合的笔数
                        self.data[operator_id]['reduced_total'] = 0  # 该商户的优惠总额
                        self.data[operator_id]['coupon_1'] = 0  # 优惠券1的张数
                        self.data[operator_id]['coupon_2'] = 0  # 优惠券2的张数
                        self.data[operator_id]['coupon_3'] = 0  # 优惠券3的张数
                        self.data[operator_id]['received_total'] = Decimal('0.0')  # 该商户的实收总额
                        self.data[operator_id]['refund_count'] = 0  # 该商户的退款笔数
                        self.data[operator_id]['refund_total'] = Decimal('0.0')  # 该商户的退款总额

                    self.data[operator_id]['trade_count'] += 1  # 该商户的交易笔数加1

                    reduced_price = row['优惠金额(元)']
                    if Decimal(reduced_price) >= min(self.c_par_1, self.c_par_2, self.c_par_3):  # 只有大于3张优惠券最小值的才是使用了优惠券，才计算为优惠笔数
                        reduced_money = get_reduced_money(reduced_price)  # 没使用优惠券时返回None
                        if reduced_money:
                            self.data[operator_id]['reduced_count'] += 1  # 该商户的优惠笔数加1
                            self.data[operator_id]['reduced_total'] += reduced_money  # 该商户的优惠总额累加

                            combination = get_combination(reduced_money, self.c_par_1, self.c_par_2, self.c_par_3, self.c_max_1, self.c_max_2, self.c_max_3)
                            if len(combination) == 1:
                                self.data[operator_id]['coupon_1'] += combination[0][0]  # 优惠券1的张数累加
                                self.data[operator_id]['coupon_2'] += combination[0][1]  # 优惠券2的张数累加
                                self.data[operator_id]['coupon_3'] += combination[0][2]  # 优惠券3的张数累加
                                self.data[operator_id]['reduced_succeed_count'] += 1
                            else:  # 如果组合没有或者多于一种，请求人工核对，将这一条 row 保存到 faild.csv 中
                                self.faild_rows.append(dict(row))  # 先将OrderedDict转换成Dict
                                self.data[operator_id]['reduced_faild_count'] += 1

                    received_price = Decimal(row['实收金额(元)'])
                    self.data[operator_id]['received_total'] += received_price  # 该商户的实收总额累加

                    if row['状态'].strip() == '退款':  # 状态为 '退款' 时，退款金额看 实收金额(元) 的值
                        self.data[operator_id]['refund_count'] += 1  # 该商户的退款笔数加1
                        self.data[operator_id]['refund_total'] += received_price  # 该商户的退款总额累加

            self.label_15.setText(filename)  # 显示选择的文件名
        except (UnicodeDecodeError, KeyError) as e:
            QMessageBox.warning(self, '警告', '请选择正确的 [快收后台的账务明细数据] 文件!')
        # print(self.data)
        # print(self.faild_rows)

    def chose_shops_csv(self):
        filename, filetype = QFileDialog.getOpenFileName(self, '选择文件', basepath)
        if filename:
            self.shops_csv = filename
            self.get_shops(self.shops_csv)

    def get_shops(self, filename):
        self.shops = {}  # 每次调用该函数时，都需要先清空之前的数据
        try:
            with open(filename) as f:
                f_csv = csv.DictReader(f)
                for row in f_csv:
                    if not row['操作员账号']:  # csv文件最后一行为空
                        continue
                    operator_id = row['操作员账号'].strip()
                    shop = row['店铺名称']
                    if operator_id not in self.shops:
                        self.shops[operator_id] = shop
            self.label_17.setText(filename)  # 显示选择的文件名
            self.cb_2.setChecked(True)  # 勾选列
        except (UnicodeDecodeError, KeyError) as e:
            QMessageBox.warning(self, '警告', '请选择正确的 [店铺名称与操作员账号对应关系] 文件!')
        # print(self.shops)

    def modify_rate(self):
        # 后面四个数字的作用依次是 初始值 最小值 最大值 小数点后位数  
        text, ok = QInputDialog.getDouble(self, '修改费率', '请输入费率(单位 %)：', 0.60, -10000, 10000, 2)
        if ok:
            self.rate = text
            self.label_21.setText(str(text))

    def output_reslut(self):
        if not self.data:
            QMessageBox.warning(self, '警告', '请选择 [快收后台的账务明细数据] 文件!')
            return
        if self.cb_2.isChecked() and not self.shops:
            QMessageBox.warning(self, '警告', '由于您勾选了[店铺名称]，所以请选择 [操作员账号与店铺名称对应表] 文件!')
            return

        filename, filetype = QFileDialog.getSaveFileName(self, '导出报表', basepath, 'CSV Files (*.csv)')
        if filename:
            self.rows = []  # 最终要插入 result.csv 文件的记录，每个元素也是一个列表

            # 合计总数
            total_trade_count = 0
            total_reduced_count = 0
            total_reduced_succeed_count = 0
            total_reduced_faild_count = 0
            total_reduced = Decimal('0')
            total_coupon_1 = 0
            total_coupon_2 = 0
            total_coupon_3 = 0
            total_received = Decimal('0')
            total_refund_count = 0
            total_refund = Decimal('0')

            for key, value in self.data.items():
                row = []  # 统计结果csv文件的每一行数据

                # 如果未选择复选框，则统计报表中不输出该列
                if self.cb_1.isChecked():
                    row.append(key)

                if self.cb_2.isChecked():
                    if key in self.shops:
                        row.append(self.shops[key])
                    elif key.lstrip('0') in self.shops:
                        row.append(self.shops[key.lstrip('0')])
                    else:
                        row.append('')

                if self.cb_3.isChecked():
                    row.append(value['trade_count'])

                if self.cb_4.isChecked():
                    row.append(value['reduced_count'])

                if self.cb_5.isChecked():
                    row.append(value['reduced_total'])

                if self.cb_6.isChecked():
                    row.append(value['coupon_1'])

                if self.cb_7.isChecked():
                    row.append(value['coupon_2'])

                if self.cb_8.isChecked():
                    row.append(value['coupon_3'])

                if self.cb_9.isChecked():
                    row.append(value['received_total'])

                if self.cb_10.isChecked():
                    row.append(value['refund_count'])

                if self.cb_11.isChecked():
                    row.append(value['refund_total'])

                self.rows.append(row)

                # 合计统计
                total_trade_count += value['trade_count']
                total_reduced_count += value['reduced_count']
                total_reduced_succeed_count += value['reduced_succeed_count']
                total_reduced_faild_count += value['reduced_faild_count']
                total_reduced += value['reduced_total']
                total_coupon_1 += value['coupon_1']
                total_coupon_2 += value['coupon_2']
                total_coupon_3 += value['coupon_3']
                total_received += value['received_total']
                total_refund_count += value['refund_count']
                total_refund += value['refund_total']

            headers = []  # csv表头
            totals = []  # csv合计

            if self.cb_1.isChecked():
                headers.append('操作员账号')
                totals.append('合计')
            if self.cb_2.isChecked():
                headers.append('店铺名称')
                totals.append('(Total):')
            if self.cb_3.isChecked():
                headers.append('交易笔数')
                totals.append(total_trade_count)
            if self.cb_4.isChecked():
                headers.append('优惠笔数')
                totals.append(total_reduced_count)
            if self.cb_5.isChecked():
                headers.append('优惠金额(元)')
                totals.append(total_reduced)
            if self.cb_6.isChecked():
                headers.append(self.cb_6.text())
                totals.append(total_coupon_1)
            if self.cb_7.isChecked():
                headers.append(self.cb_7.text())
                totals.append(total_coupon_2)
            if self.cb_8.isChecked():
                headers.append(self.cb_8.text())
                totals.append(total_coupon_3)
            if self.cb_9.isChecked():
                headers.append('实收金额(元)')
                totals.append(total_received)
            if self.cb_10.isChecked():
                headers.append('退款笔数')
                totals.append(total_refund_count)
            if self.cb_11.isChecked():
                headers.append('退款总额(元)')
                totals.append(total_refund)

            try:
                with open(filename, 'w', newline='') as f:
                    f_csv = csv.writer(f)
                    f_csv.writerow(headers)
                    f_csv.writerows(self.rows)
                    f_csv.writerow(totals)

                    f_csv.writerow([])  # 空一行
                    f_csv.writerow(['商场活动数据', '实收金额', '实收金额手续费', '实际结算金额', '商场到账金额', '差额'])
                    rate_price = total_received * Decimal(str(self.rate / 100))
                    f_csv.writerow(['', total_received, rate_price, total_received-rate_price, '/', '/'])

                    f_csv.writerow([])  # 空一行
                    f_csv.writerow(['结算公式:'])
                    f_csv.writerow(['实收金额=快收后台实收金额'])
                    f_csv.writerow(['实收金额手续费=实收金额*{}%'.format(self.rate)])
                    f_csv.writerow(['实际结算金额=实收金额-实收金额手续费'])
                    QMessageBox.information(self, '导出成功', '统计报表为 [{}]，请用WPS等程序查看!'.format(filename))
            except PermissionError as e:
                QMessageBox.warning(self, '导出失败', '指定的导出文件 [{}] 可能正在被WPS等程序打开，请先关闭它!'.format(filename))
                return

            # 导出优惠券无法计算的行
            if self.faild_rows:
                QMessageBox.warning(self, '警告', '有 [{}] 笔优惠金额无法计算出使用的优惠券组合( 成功计算出 [{}] 笔 )，请人工核对。请单击 [OK] 按钮后，选择要保存这些记录到哪?'.format(total_reduced_faild_count, total_reduced_succeed_count))
                filename, filetype = QFileDialog.getSaveFileName(self, '导出异常的优惠券信息', basepath, 'CSV Files (*.csv)')
                if filename:
                    try:
                        headers = [key for key in self.faild_rows[0].keys()]  # 随便拿一条，拿Dict的keys，生成csv表头
                        with open(filename, 'w', newline='') as f:
                            f_csv = csv.DictWriter(f, headers)
                            f_csv.writeheader()
                            f_csv.writerows(self.faild_rows)
                            QMessageBox.warning(self, '警告', '优惠券数量无法正确计算的交易记录导出为 [{}]，请用WPS等程序查看!'.format(filename))
                    except PermissionError as e:
                        QMessageBox.warning(self, '警告', '用来保存 [优惠券数量无法正确计算的交易记录] 的csv文件 [{}] 可能正在被WPS等程序打开，请先关闭它!'.format(filename))
                        return
                else:
                    QMessageBox.warning(self, '警告', '有部分交易中的优惠券数量无法计算，需要人工核对，但是您没有选择要导出的文件!')
        else:
            QMessageBox.warning(self, '警告', '统计报表未导出，请指定导出文件!')


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = Ui_MainWindow()
    ex.show()
    sys.exit(app.exec_())
