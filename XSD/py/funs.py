# -*- coding: utf-8 -*-
import xlrd, datetime, shelve
from suds.client import Client
from py import config

new_time = datetime.datetime.now().strftime('%Y-%m-%d')


class Xls:
    def __init__(self, path):
        self.path = path

    def Open_sheet(self):
        # 打开这个Excel
        xl = xlrd.open_workbook(self.path)
        sheet = xl.sheet_by_index(0)  # 打开这个Excel的第一个sheet
        hang_num = sheet.nrows - 1  # Excel的行数
        return sheet, hang_num

    def Phone_list(self):
        # 获取{'13452904579': '程翔'}列表
        sheet = self.Open_sheet()[0]
        num = sheet.ncols
        for x in range(0, num):
            if sheet.col_values(x)[0] == '联系地址':
                M = sheet.col_values(x)[1:]
            if sheet.col_values(x)[0] == '客户姓名':
                N = sheet.col_values(x)[1:]
        L = dict(zip(M, N))  # {u'phone': u'姓名'}
        # file_log.write(u'%s获取到号码列表有%s条。\n' % (new_time, len(L)))
        return L


class XW_sms:
    def __init__(self):
        # 定义账号密码
        self.account, self.password = config.SMS_user, config.SMS_password
        url = config.url
        self.client = Client(url)

    def find_report(self, phone, name):
        info = self.client.service.FindReport(account=self.account, password=self.password, batchid='', mobile=phone,
                                              pageindex='1', flag='1')
        T = F = W = 0
        DE_ = None
        try:
            data = info['MTReport']
            for x in data:
                status = x['reserve']
                # DE_ = x['originResult']
                DE_ = str(x['originResult'].split(",")[1])
                a, b, c, d = status[0], status[6], status[12], status[2]  # 4/4成功;0/4失败;0/4等待
                if a == d:
                    T += 1
                elif b == d:
                    F += 1
                elif b == 0 and c != 0:
                    W += 1
                else:
                    T = F = W = 6  # 都为6为异常
        except TypeError as e:
            W = 111  # 都为提交未返回
        stat = {
            'phone': phone,
            'name': name,
            'DELIVRD': DE_,
            'status': [T, F, W],
        }
        return stat

    def sms_end(self, phone, name):
        text = self.find_report(phone, name)
        file_db = shelve.open('./py/cache.db')
        asd = text['DELIVRD']
        try:
            Status = file_db['status'][asd]
        except KeyError as e:
            Status = '未知状态码!!!'
        if text['status'][0] != 0:
            return '%s\t%s\t%s\t%s\t%s\n' % (text['name'], text['phone'], '成功', text['DELIVRD'], Status)
        if text['status'][1] != 0:
            return '%s\t%s\t%s\t%s\t%s\n' % (text['name'], text['phone'], '失败', text['DELIVRD'], Status)
        if text['status'][2] != 0:
            status = self.client.service.FindResponse(account=config.SMS_user, password=config.SMS_password, batchid='',
                                                      mobile=phone, pageindex='1', flag='1')
            if status != None:
                return '%s\t%s\t%s\t%s\t%s\n' % (text['name'], text['phone'], '已提交', text['DELIVRD'], Status)
            else:
                return '%s\t%s\t%s\t%s\t%s\n' % (text['name'], text['phone'], '短信未到玄武', text['DELIVRD'], Status)
