# -*- coding: utf-8 -*-

import sys, shelve, time, os
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QFileDialog, QSplashScreen, QMainWindow, QApplication
from py.Ui_xsd import Ui_MainWindow
from py.funs import XW_sms, Xls


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
        # 检测py目录是否存在，如不存在则新建
        path = [x for x in os.listdir('.') if os.path.isdir(x)]
        if 'py' not in path:
            os.mkdir('py')
        time.sleep(0.5)
        L = {'BLACK': '通道黑名单;该手机号码为通道黑名单，请确认是否一定要让其接收信息，如是，请取消黑名单后再进行下发',
             'DB:0140': '用户不在白名单中;白名单通道中发送的用户不在白名单中，请先上传白名单再进行下发', 'DELIVRD': '成功;短信发送成功。',
             'ID:0076': '信息安全鉴权失败(包含敏感字);有可能运营商侧有更新新的敏感字，请检查内容中，换一种表述或者去掉敏感字，并将新的敏感字添加入系统的敏感字过滤中',
             'ILLEGALKEY': 'UMP-敏感字;短信中包含敏感字，被系统拦截。', 'MC:0055': '基站与手机通讯超时;手机与基站通讯时，可能因天气、磁场等环境因素导致基站与手机通讯异常',
             'MI:0000': '空号;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             'MI:0004': '空号;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             'MI:0010': '关机;、停机确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             'MI:0013': '停机;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             'MI:0017': '运营商短信推送至手机超时;手机长时间不在服务区或者信号不稳定导致，如果需要，可考虑重发',
             'MI:0022': '运营商短信推送至手机超时;手机与基站通讯时，可能因天气、磁场等环境因素导致基站与手机通讯异常',
             'MI:0024': '关机/停机;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             'MI:0029': '关机/不在服务区;因手机长时间关机/不在服务区，超过短信有效期时间(通常为48小时),建议核查该号码是否仍然有效',
             'MK:0001': '空号;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             'MK:0004': '号码暂停使用;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             'MK:0005': '停机/号码失效;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             'MK:0012': '空号;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             'MK:0024': '用户关机;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             'MK:0029': '无法接通手机;手机长时间不在服务区或者信号不稳定导致，如果需要，可考虑重发',
             'MN:0001': '空号;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             'MN:0017': '手机内存已满;手机收件箱中的内存已满，无法保存接收回来的短信，建议让接收者清理一下手机短信收件箱中的内容再尝试下发',
             'NOTITLE': '未加退订指令;短信内容最后面需要加上回复TD退订',
             'NP:1243': '携号转网客户;移动的号码携号转至联通或者电信，一般在天津和海南号码会出现这种情况，请在系统中将该号码添加至对应的运营商号段路由中', '0': '成功;短信发送成功;',
             '1': '关机或空号;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '10': '手机无法接通或者关机;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '101': '空号;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '11': '欠费停机;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '12': '空号/停机;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收', '13': '呼叫受限;',
             '15': '手机终端通讯故障或者信号问题;建议拨打10010通知联通客服处理',
             '23': '空号;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '24': '空号/停机;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '50': '短消息内容非法;有可能运营商侧有更新新的敏感字，请检查内容中，换一种表述或者去掉敏感字，并将新的敏感字添加入系统的敏感字过滤中。',
             '56': '空号、停机;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '59': '关机;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收', '67': '通道黑名单;建议手机号用户直接拨打当地10010客服解除屏蔽',
             '86': '通道黑名单;有可能运营商侧有更新新的敏感字，请检查内容中，换一种表述或者去掉敏感字，并将新的敏感字添加入系统的敏感字过滤中。',
             'R:00090': '手机终端故障; 短信存储满手机收件箱中的内存已满，无法保存接收回来的短信，建议让接收者清理一下手机短信收件箱中的内容再尝试下发',
             'sk:8102': '通道地区限制;(该通道已没有在用)', '000': '成功;手机正常收到短信',
             '006': '空号/关键字/黑名单;请检查该号码使用状态,如为空号，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收。 否则检查内容是否有敏感字，如无，请检查是否在通道黑名单',
             '039': '空号;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '601': '空号/停机;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '602': '关机;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '614': '手机端接收短信失败;有可能是基站与手机通讯时，网络原因导致短信接收失败，请确认该手机是否经常信号不稳定',
             '660': '短信被运营商拦截;有可能运营商侧有更新新的敏感字，请检查内容中，换一种表述或者去掉敏感字，并将新的敏感字添加入系统的敏感字过滤中。',
             '701': '关机/空号;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '705': '空号;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '713': '关机，欠费停机;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '760': '空号/信号不稳定;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '771': '空号/停机;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '801': '关机/空号;确认该号码的使用状态，如果是空号/停机/长时间关机，建议考虑更新该收件人的手机接收号码，以便后续能让正常接收',
             '999': '业务网关黑名单;该手机号码为通道黑名单，请确认是否一定要让其接收信息，如是，请取消黑名单后再进行下发', 'INSUFFICIENT_BALANCE': '余额不足;账号余额不足，请及时充值',
             'MORETHAN_SENDTIMES': '限制6条;系统限制半小时最多发送6条，多余的将被系统拦截。', 'REPEAT_SEND': '重复提交;短时间(同一秒)内提交同一内容给同一号码。'}
        file_db = shelve.open('./py/cache.db')
        file_db['status'] = L
        file_db.close()

    @pyqtSlot()
    def on_pushButton_clicked(self):
        file_path = \
            QFileDialog.getOpenFileName(self, '校验文件', 'C:/Users/heiying/Desktop',
                                        'xlsx Files (*.xls);;XLS Files (*.xlsx)')[
                0]
        file_db = shelve.open('./py/cache.db')
        file_db['file_path'] = [file_path]
        file_db.close()

    @pyqtSlot()
    def on_pushButton_2_clicked(self):
        Echo_line = ""
        file_db = shelve.open('./py/cache.db')
        try:
            file_path = file_db['file_path'][0]
            phone_data = Xls(file_path)
            xw = XW_sms()
            phone_list = phone_data.Phone_list()
            for x in phone_list:
                phone = xw.sms_end(str(x)[:11], phone_list[x])
                if '成功' not in phone:
                    Echo_line = Echo_line + phone
            Echo_line = Echo_line + 'END'
            self.textBrowser.setText(Echo_line)
        except FileNotFoundError as a:
            self.textBrowser.setText('导入文件为空，请导入检测文件!!!')
        except KeyError as b:
            self.textBrowser.setText('导入文件为空，请导入检测文件!!!')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    splash = QSplashScreen(QPixmap(':/img/xsd_tm.png'))
    splash.show()
    app.processEvents()
    ui = MainWindow()
    ui.show()
    splash.finish(ui)
    sys.exit(app.exec_())
