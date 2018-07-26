import requests
import re
import pickle
from bs4 import BeautifulSoup
import os

# 模拟浏览器头
HEADER = {'User-Agent': ('Mozilla/4.0 (compatible; MSIE 7.0; Windows NT'
                         '6.1; Trident/7.0; SLCC2; .NET CLR 2.0.50727; '
                         '.NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET4'
                         '.0C; .NET4.0E)')}
# 工号姓名信息从文件汇总反序列化出来
sf = os.path.join(os.path.dirname(__file__), 'staff2name')
with open(sf, 'rb') as f:
    staff2name = pickle.load(f)
# 内网地址
url = 'http://www.tc.net'


def nw_chakan(shouid):
    '''内网查看
       shouid   个人事务接收信息ID
    '''
    rep = requests.get('%s/grsw/shou/chakan.aspx?id=%s' % (url, shouid),
                       cookies={'tcdx_admin': ''})
    if re.search('查看信息', rep.text):
        return True
    else:
        return False


def nw_zhuanfa(receivers,
               shouid,
               fjid,
               title='',
               content='',
               sender='600299'):
    '''转发函数
       receivers  接收工号列表
       shouid     个人事务接收信息ID
       fjid       附件ID
       title      发送信息标题
       content    发送信息内容
       sender     发送人工号
    '''
    # 发送人COOKIE
    cookie = {'tcdx_admin': 'id={0}&work_no={1}'.format(
        staff2name[sender][0], sender)}
    # 通过get请求获取2个参数值
    rep = requests.get('%s/grsw/shou/zhuanfa.aspx?id=%s' % (url, shouid),
                       cookies=cookie, headers=HEADER)
    soup = BeautifulSoup(rep.text, 'html.parser')
    VIEWSTATE = soup.find(id='__VIEWSTATE')['value']
    EVENTVALIDATION = soup.find(id='__EVENTVALIDATION')['value']
    # 接收人列表
    person = ''
    for gh in receivers:
        person += '[{0}]{1};'.format(gh, staff2name[gh][1])
    # 构造发送信息
    files = {
        '__VIEWSTATE': (None, VIEWSTATE),
        '__EVENTVALIDATION': (None, EVENTVALIDATION),
        'HiddenField2': (None, person),
        'TextBox2': (None, title),
        'editor1$TextBox1': (None, content),
        'Button1': (None, '发送'),
    }
    if fjid:
        files['fjid_%s' % fjid] = (None, fjid)
    # 转发信息
    rep = requests.post('%s/grsw/shou/zhuanfa.aspx?id=%s' % (url, shouid),
                        files=files, cookies=cookie, headers=HEADER)
    if re.search('发送成功', rep.text):
        return True
    else:
        return False


def nw_fa(receivers, title='', content='', filename=None, sender='600299'):
    '''发送函数
       receivers  接收工号列表
       title      发送信息标题
       content    发送信息内容
       filename   发送文件
       sender     发送人工号
    '''
    # 发送人COOKIE
    cookie = {'tcdx_admin': (
        'id={0}&work_no={1}'.format(staff2name[sender][0], sender))}
    # 通过get请求获取2个参数值
    rep = requests.get('%s/grsw/fa/add.aspx' % url,
                       cookies=cookie, headers=HEADER)
    soup = BeautifulSoup(rep.text, 'html.parser')
    VIEWSTATE = soup.find(id='__VIEWSTATE')['value']
    EVENTVALIDATION = soup.find(id='__EVENTVALIDATION')['value']
    # 读取发送文件
    if filename:
        with open(filename, 'rb') as f:
            file = f.read()
    else:
        file = ''
    # 接收人列表
    person = ''
    for gh in receivers:
        person += '[{0}]{1};'.format(gh, staff2name[gh][1])
    # 构造发送信息
    files = {
        '__VIEWSTATE': (None, VIEWSTATE),
        '__EVENTVALIDATION': (None, EVENTVALIDATION),
        'HiddenField2': (None, person),
        'TextBox2': (None, title),
        'editor1$TextBox1': (None, content),
        'Button1': (None, '发送'),
        'file_1': (filename, file)
    }
    # 发送信息
    rep = requests.post('%s/grsw/fa/add.aspx' % url,
                        files=files, cookies=cookie, headers=HEADER)
    if re.search('发送成功', rep.text):
        return True
    else:
        return False


if __name__ == '__main__':
    staff = input('send to: ')
    title = input('title: ')
    content = input('content: ')
    filename = input('file: ')
    if not title:
        title = filename.split('.')[0]
    staff = [staff]
    if nw_fa(staff, title, content, filename):
        print('发送成功')
    input('按任意键退出')
