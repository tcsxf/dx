import requests
import re
import json
from bs4 import BeautifulSoup
import os


class NW():
    def __init__(self, staff='600299', url='http://www.tc.net'):
        # 模拟登陆cookie,根据staff从文件读取信息
        sfile = os.path.join(os.path.dirname(__file__), 'staff')
        with open(sfile, 'r') as f:
            self.staff_list = json.load(f)
        for staff_ in self.staff_list:
            if staff_['nm'] == staff or staff_['staff'] == staff:
                s = staff_
                break
        self.cookie = {
            'tcdx_admin': 'id={0}&work_no={1}'.format(s['id'], s['staff'])
        }
        # 模拟浏览器头
        self.header = {
            'User-Agent': ('Mozilla/5.0 (Windows NT 6.1; WOW64) '
                           'AppleWebKit/537.36 (KHTML, like Gecko'
                           ') Chrome/69.0.3497.100 Safari/537.36')}
        # url地址
        self.url = url

    def get_param(self, url):
        # 模拟cookie获取相应页面的2个关键值
        rep = requests.get(url,
                           cookies=self.cookie,
                           headers=self.header)
        soup = BeautifulSoup(rep.text, 'html.parser')
        VIEWSTATE = soup.find(id='__VIEWSTATE')['value']
        EVENTVALIDATION = soup.find(id='__EVENTVALIDATION')['value']
        return VIEWSTATE, EVENTVALIDATION

    def nw_chakan(self, shouid):
        '''内网查看
           shouid   个人事务接收信息ID
        '''
        rep = requests.get(f'{self.url}/grsw/shou/chakan.aspx?id={shouid}',
                           cookies=self.cookie)
        if re.search('查看信息', rep.text):
            return True
        else:
            return False

    def nw_zhuanfa(self,
                   receivers,
                   shouid,
                   fjid,
                   title='',
                   content='',):
        '''转发函数
           receivers  接收工号列表
           shouid     个人事务接收信息ID
           fjid       附件ID
           title      发送信息标题
           content    发送信息内容
        '''
        # 获取2个参数值
        v, e = self.get_param(f'{self.url}/grsw/shou/zhuanfa.aspx?id={shouid}')
        # 接收人列表
        persons = []
        for staff in self.staff_list:
            if staff['staff'] in receivers or staff['nm'] in receivers:
                persons.append('[{0}]{1};'.format(staff['staff'], staff['nm']))
        # 构造发送信息
        files = {
            '__VIEWSTATE': (None, v),
            '__EVENTVALIDATION': (None, e),
            'HiddenField2': (None, ''.join(persons)),
            'TextBox2': (None, title),
            'editor1$TextBox1': (None, content),
            'Button1': (None, '发送'),
        }
        if fjid:
            files['fjid_%s' % fjid] = (None, fjid)
        # 转发信息
        rep = requests.post(f'{self.url}/grsw/shou/zhuanfa.aspx?id={shouid}',
                            files=files,
                            cookies=self.cookie,
                            headers=self.header)
        if re.search('发送成功', rep.text):
            return True
        else:
            return False

    def nw_fa(self,
              receivers,
              title='',
              content='',
              filename=None,):
        '''发送函数
           receivers  接收工号列表
           title      发送信息标题
           content    发送信息内容
           filename   发送文件
        '''
        # 获取2个参数值
        v, e = self.get_param(f'{self.url}/grsw/fa/add.aspx')
        # 读取发送文件
        if filename:
            with open(filename, 'rb') as f:
                file = f.read()
        else:
            file = ''
        # 接收人列表
        persons = []
        for staff in self.staff_list:
            if staff['staff'] in receivers or staff['nm'] in receivers:
                persons.append('[{0}]{1};'.format(staff['staff'], staff['nm']))
        # 构造发送信息
        files = {
            '__VIEWSTATE': (None, v),
            '__EVENTVALIDATION': (None, e),
            'HiddenField2': (None, ''.join(persons)),
            'TextBox2': (None, title),
            'editor1$TextBox1': (None, content),
            'Button1': (None, '发送'),
            'file_1': (filename, file)
        }
        # 发送信息
        rep = requests.post(f'{self.url}/grsw/fa/add.aspx',
                            files=files,
                            cookies=self.cookie,
                            headers=self.header)
        if re.search('发送成功', rep.text):
            return True
        else:
            return False


if __name__ == '__main__':
    staff = input('send to: ')
    title = input('title: ')
    content = input('content: ')
    filename = input('file: ')
    nw = NW()
    if nw.nw_fa(staff.split(','), title, content, filename):
        print('发送成功')
    input('按任意键退出')
