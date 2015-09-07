# -*- coding: utf-8 -*-

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import xlrd
import xlwt
import os
import xlutils.copy
import shutil
import base64
####################################
#######打开一个新xls，并初始化######
####################################
style_list = xlwt.easyxf("font: bold on,colour_index white,height 250;" \
                         "align: horizontal center;borders:left 1,right 1,top 1,bottom 1;" \
                         "pattern: pattern solid_pattern, fore_colour green, back_colour green")
style_ordinary = xlwt.easyxf("font: bold off,colour_index black,height 200;" \
                         "align: horizontal center;" )


def combine_sheet():
    file_path = os.getcwd().decode('GBK')
    name_list = []
    for i in os.listdir(file_path):
        name_list.append(i)
    combine_list = []
    xls_name = ['GGSN', 'SGSN']
    for xls in xls_name:
        for m in name_list:
            if xls in m:
                combine_list.append(m)
        if len(combine_list) == 1:
            shutil.copyfile(combine_list[0], xls[0:2] + '.xls')

        #print combine_list
        else:
            ggsn_1 = read_excel(combine_list[0])
            ggsn_2 = read_excel(combine_list[1])
            ggsn_1_table = ggsn_1.sheet_by_index(0)
            ggsn_2_table = ggsn_2.sheet_by_index(0)
            if ggsn_1_table.nrows > ggsn_2_table.nrows:
                big = ggsn_1_table
                small = ggsn_2_table
                com_xls = ggsn_1
            else:
                big = ggsn_2_table
                small = ggsn_1_table
                com_xls = ggsn_2
            com_xls = xlutils.copy.copy(com_xls)
            com_sheet = com_xls.get_sheet(0)
            small_row = 8
            for i in range(big.nrows, big.nrows + small.nrows - 8):
                small_col = 0
                for j in range(big.ncols):
                    com_sheet.write(i, j, small.cell_value(small_row, small_col))
                    small_col += 1
                small_row += 1
            file_name = xls[0:2] + '.xls'
            com_xls.save(file_name)
        combine_list = []
    return name_list
def open_excel():
    try:
        book = xlwt.Workbook()
        return book
    except Exception, e:
        print str(e)


def bb_excel_init(book, day_list):
    #print end_day
    sheet_daily = book.add_sheet(u'日吞吐量')
    sheet_user = book.add_sheet(u'最大用户激活数')
    sheet_daily.write(0, 0, u'日期', style_list)
    sheet_daily.write(0, 1, u'4G接入（总）', style_list)
    sheet_daily.write(0, 2, u'3G接入（总）', style_list)
    sheet_daily.write(0, 3, u'2G接入（总）', style_list)
    sheet_daily.write(0, 4, u'本省(2G\\3G\\4G)', style_list)
    sheet_daily.write(0, 5, u'4G接入（huawei）', style_list)
    sheet_daily.write(0, 6, u'3G接入（huawei）', style_list)
    sheet_daily.write(0, 7, u'2G接入（huawei）', style_list)
    sheet_daily.write(0, 8, u'4G接入（ZTE）', style_list)
    sheet_daily.write(0, 9, u'3G接入（ZTE）', style_list)
    sheet_daily.write(0, 10, u'2G接入（ZTE）', style_list)
    sheet_user.write(0, 0, u'日期', style_list)
    sheet_user.write(0, 1, u'4G（总）', style_list)
    sheet_user.write(0, 2, u'3G（总）', style_list)
    sheet_user.write(0, 3, u'2G（总）', style_list)
    sheet_user.write(0, 4, u'本省', style_list)
    sheet_user.write(0, 5, u'4G（huawei）', style_list)
    sheet_user.write(0, 6, u'3G（huawei）', style_list)
    sheet_user.write(0, 7, u'2G（huawei）', style_list)
    sheet_user.write(0, 8, u'4G（ZTE）', style_list)
    sheet_user.write(0, 9, u'3G（ZTE）', style_list)
    sheet_user.write(0, 10, u'2G（ZTE）', style_list)

    for i in range(0, 8):
        sheet_daily.col(i).width = 0x0d00 + 2500
        sheet_user.col(i).width = 0x0d00 + 2500
    m_index = 1
    for m in day_list:
        insert_context = m[0] + u'月' + m[1] + u'日'
        sheet_daily.write(m_index, 0, insert_context, style_ordinary)
        sheet_user.write(m_index, 0, insert_context, style_ordinary)
        m_index += 1
    return book


def save_excel(book, day_list, mail_user, mail_pwd):
    excel_name = str(day_list[0][2]) + u'年' + str(day_list[0][0]) + u'月月报数据.xls'
    book.save(excel_name)
    msg = MIMEMultipart()
    mail_host = "hellokitty.com"
    mail_port = 25
    mail_to = "###@hellokitty.com"
    att = MIMEText(open((excel_name),'rb').read(),'base64','utf-8')
    #att["Content-Type"] = 'application/octet-stream'
    #content = '"' + excel_name.encode("utf-8") + '"'
    att["content-Disposition"] = 'attachment; filename =' + '"'+ (excel_name).encode("utf-8")+'"'
    msg.attach(att)
    #message = 'content part'
    #body = MIMEText(message)
    #msg.attach(body)
    msg['To'] = mail_to
    msg['from'] = mail_user
    msg['subject'] = unicode(excel_name)

    try:
        smtp = smtplib.SMTP(mail_host, mail_port)
        smtp.set_debuglevel(1)
        #smtp.ehlo()
        #smtp.starttls()
        smtp.login(mail_user, mail_pwd)
        smtp.sendmail(mail_user, mail_to, msg.as_string())
        smtp.close
    except Exception, e:
        print e




####################################
############读取现有xls#############
####################################
def read_excel(excel_name):
    try:
        data = xlrd.open_workbook(excel_name)
        return data
    except Exception, e:
        print str(e)


def get_time(table, row):
    cell_value = table.cell_value(row-1, 0)
    day_time_list = cell_value.split(' ')
    day_time_list[0] = day_time_list[0].split('/')
    day_time_list[1] = day_time_list[1].split(':')
    if day_time_list[0][0][0] == '0':
        day_time_list[0][0] = day_time_list[0][0][1]
    if day_time_list[0][1][0] == '0':
        day_time_list[0][1] = day_time_list[0][1][1]
    return day_time_list


def get_day_list(table):
    day_list = []
    for i in range(9, table.nrows+1):
        if get_time(table, i)[0] not in day_list:
            day_list.append(get_time(table, i)[0])
    return day_list


def insert_data_sheet1(table, baobiao, day_list, col_1, col_2, col_w):
    start_row = 9
    m_index = 0
    sheet_daily = baobiao.get_sheet(0)
    for m in day_list:
        sum_value = 0.0
        n = get_time(table, start_row)[0]
        while m == n:
            if col_w == 4:
                if table.cell_value(start_row-1, col_2+5) == 'NIL':
                    value_col_1 = 0
                    value_col_2 = 0
                else:
                    value_col_2 = float(table.cell_value(start_row-1, col_2+5))
                    value_col_1 = float(table.cell_value(start_row-1, col_2+6))
                sum_value += (float(table.cell_value(start_row-1, col_1))+float(table.cell_value(start_row-1, col_2)))/1024+(value_col_1+value_col_2)/1024/1024
            else:
                if table.cell_value(start_row-1, col_1) == 'NIL'or table.cell_value(start_row-1, col_2) == 'NIL':
                    value_col_1 = 0
                    value_col_2 = 0
                else:
                    value_col_1 = float(table.cell_value(start_row-1, col_1))
                    value_col_2 = float(table.cell_value(start_row-1, col_2))
                #value_col_2 = float(table.cell_value(start_row-1, col_2))
                #value_col_1 = float(table.cell_value(start_row-1, col_1))
                sum_value += (value_col_1+value_col_2)/1024/1024
            start_row += 1
            if start_row > table.nrows:
                n = None
                continue
            n = get_time(table, start_row)[0]
        else:
            m_index += 1
            sheet_daily.write(m_index, col_w, round(sum_value, 3), style_ordinary)
            continue


def insert_data_sheet2(table, baobiao, day_list, col, col_w):
    start_row = 9
    sheet_user = baobiao.get_sheet(1)
    m_index = 0
    time_list = []
    for i in range(9, table.nrows+1):
        if get_time(table, i)[1] not in time_list:
            time_list.append(get_time(table, i)[1])
    for m in day_list:
        n = get_time(table, start_row)[0]
        if m == n:
            max_tmp = 0
            for x in time_list:
                t = get_time(table, start_row)[1]
                tmp_sum = 0
                while x == t:
                    if col_w ==4:
                        tmp_sum += table.cell_value(start_row-1, col)+table.cell_value(start_row-1, col+9)
                    else:
                        tmp_sum += table.cell_value(start_row-1, col)
                    start_row += 1
                    if start_row > table.nrows:
                        t = None
                        continue
                    t = get_time(table, start_row)[1]
                else:
                    if tmp_sum > max_tmp:
                        max_tmp = tmp_sum
                    continue
            m_index += 1
            sheet_user.write(m_index, col_w, max_tmp, style_ordinary)

            #n = get_time(table, start_row)[0]

# zte actived content:
def insert_data_sheet2_zte(table, baobiao, col, col_w):
    start_row = 6
    sheet_user = baobiao.get_sheet(1)
    m_index = 1
    print u"中兴有效的数据量为"+str((table.nrows-6)/2.0)+u'天'
    for i in range(start_row, table.nrows):
        if i%2 == 0:
            sum_value = table.cell_value(i, col) + table.cell_value(i+1, col)
            sheet_user.write(m_index, col_w, sum_value, style_ordinary)
            m_index += 1

def insert_data_sheet1_zte(table, baobiao, col, col_w):
    start_row = 6
    sheet_user = baobiao.get_sheet(0)
    m_index = 1
    print u"中兴有效的数据量为"+str((table.nrows-6)/2.0)+u'天'
    if col_w == 9:
        for i in range(start_row, table.nrows):
            if i%2 == 0:
                sum_value = table.cell_value(i, col) + table.cell_value(i+1, col)
                sum_value += table.cell_value(i, col+1) + table.cell_value(i+1, col+1)
                sum_value = float(sum_value)*24*3600/1024/8
                sheet_user.write(m_index, col_w, round(sum_value, 3), style_ordinary)
                m_index += 1
    elif col_w == 10:
        for i in range(start_row, table.nrows):
            if i%2 == 0:
                sum_value = table.cell_value(i, col) + table.cell_value(i+1, col)
                sum_value += table.cell_value(i, col+1) + table.cell_value(i+1, col+1)
                sum_value += table.cell_value(i, col+4) + table.cell_value(i+1, col+4)
                sum_value += table.cell_value(i, col+5) + table.cell_value(i+1, col+5)
                sum_value = float(sum_value)*24/1024/8*3600
                sheet_user.write(m_index, col_w, round(sum_value, 3), style_ordinary)
                m_index += 1

def combine_huawei_zte(baobiao, day_list):
    sheet_IO = baobiao.get_sheet(0)
    sheet_user = baobiao.get_sheet(1)
    for i in range(day_list.__len__()):
        IO_formula_3G = "G"+str(i+2)+"+"+"J"+str(i+2)
        IO_formula_2G = "H"+str(i+2)+"+"+"K"+str(i+2)
        IO_formula_4G = "F"+str(i+2)+"+"+"I"+str(i+2)
        user_formula_3G = "G"+str(i+2)+"+"+"J"+str(i+2)
        user_formula_2G = "H"+str(i+2)+"+"+"K"+str(i+2)
        user_formula_4G = "F"+str(i+2)+"+"+"I"+str(i+2)
        sheet_IO.write(i+1, 1, xlwt.Formula(IO_formula_4G))
        sheet_IO.write(i+1, 2, xlwt.Formula(IO_formula_3G))
        sheet_IO.write(i+1, 3, xlwt.Formula(IO_formula_2G))
        sheet_user.write(i+1, 1, xlwt.Formula(user_formula_4G))
        sheet_user.write(i+1, 2, xlwt.Formula(user_formula_3G))
        sheet_user.write(i+1, 3, xlwt.Formula(user_formula_2G))
    formula_4G_avg = "round("+"AVERAGE(B2:B" + str(day_list.__len__()+1)+")"+ ",3)"
    formula_3G_avg = "round("+"AVERAGE(C2:C" + str(day_list.__len__()+1)+")"+ ",3)"
    formula_2G_avg = "round("+"AVERAGE(D2:D" + str(day_list.__len__()+1)+")"+ ",3)"
    formula_GGSN_avg = "round("+"AVERAGE(E2:E" + str(day_list.__len__()+1)+")"+ ",3)"
    sheet_IO.write(35, 0, "average", style_list)
    sheet_IO.write(35, 1, xlwt.Formula(formula_4G_avg), style_list)
    sheet_IO.write(35, 2, xlwt.Formula(formula_3G_avg), style_list)
    sheet_IO.write(35, 3, xlwt.Formula(formula_2G_avg), style_list)
    sheet_IO.write(35, 4, xlwt.Formula(formula_GGSN_avg), style_list)
    sheet_user.write(35, 0, "average", style_list)
    sheet_user.write(35, 1, xlwt.Formula(formula_3G_avg), style_list)
    sheet_user.write(35, 2, xlwt.Formula(formula_2G_avg), style_list)
    sheet_user.write(35, 3, xlwt.Formula(formula_GGSN_avg), style_list)

def login_mail():
    if os.path.isfile('config.ini'):
        fs = open('config.ini', 'r')
        account = fs.readlines()
        mail = base64.decodestring(account[0][:-1]).split(':')[1]
        passwd = base64.decodestring(account[1][:-1]).split(':')[1]
    else:
        print u'请输入你的邮箱及密码，程序将会自动将做好的报表发给牟师傅。'
        print u'目前只支持联通的邮箱'
        print u'如果不小心输错了只要将当前文件夹下config.ini文件删除即可'
        fs = open('config.ini', 'w')
        mail = raw_input("mail:")
        passwd = raw_input("passwd:")
        fs.write(base64.encodestring("mail:"+mail))
        fs.write(base64.encodestring("passwd:"+passwd))
        fs.close()
    return mail,passwd
def main():
    login = login_mail()
    mail_user = login[0]
    mail_pwd = login[1]
    print u"正在处理中。。。。。。"
    try:
        os.remove('GG.xls')
        os.remove('SG.xls')
    except  WindowsError:
        pass
    file_list = combine_sheet()
    for file in file_list:
        if 'GnGp-IuPS-Gb' in file:
            data_zte_inout = read_excel(file)
        elif '234G' in file:
            data_zte_active = read_excel(file)
    ggsn_name = u'GG.xls'
    sgsn_name = u'SG.xls'
    data_ggsn = read_excel(ggsn_name)
    data_sgsn = read_excel(sgsn_name)

    table_ggsn = data_ggsn.sheet_by_index(0)
    table_sgsn = data_sgsn.sheet_by_index(0)
    table_zte_inout = data_zte_inout.sheet_by_index(0)
    table_zte_active = data_zte_active.sheet_by_index(0)
    day_list = get_day_list(table_ggsn)
    '''max_row = table_ggsn.nrows
    day_time_list = get_time(table_ggsn, max_row)
    end_day = day_time_list[0]
    year = end_day[2]
    month = end_day[0]
    d = "/"
    name_list = os.listdir(d)
    print name_list'''
    baobiao_xls = open_excel()
    baobiao_xls = bb_excel_init(baobiao_xls, day_list)
    insert_data_sheet1(table_ggsn, baobiao_xls, day_list, 6, 10, 4)
    insert_data_sheet1(table_sgsn, baobiao_xls, day_list, 13, 14, 6)
    insert_data_sheet1(table_ggsn, baobiao_xls, day_list, 18, 19, 5)
    insert_data_sheet1(table_sgsn, baobiao_xls, day_list, 15, 16, 7)
    insert_data_sheet1_zte(table_zte_inout, baobiao_xls, 5, 9)
    insert_data_sheet1_zte(table_zte_inout, baobiao_xls, 11, 10)
    insert_data_sheet2(table_sgsn, baobiao_xls, day_list, 6, 7)
    insert_data_sheet2(table_sgsn, baobiao_xls, day_list, 12, 6)
    insert_data_sheet2(table_ggsn, baobiao_xls, day_list, 8, 4)
    insert_data_sheet2(table_sgsn, baobiao_xls, day_list, 25, 5)
    insert_data_sheet2_zte(table_zte_active, baobiao_xls, 5, 10)
    insert_data_sheet2_zte(table_zte_active, baobiao_xls, 6, 9)
    insert_data_sheet2_zte(table_zte_active, baobiao_xls, 7, 8)
    combine_huawei_zte(baobiao_xls, day_list)
    os.remove('GG.xls')
    os.remove('SG.xls')
    save_excel(baobiao_xls, day_list, mail_user, mail_pwd)
    print day_list
    raw_input()
if __name__ == "__main__":
    main()



