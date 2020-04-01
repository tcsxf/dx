import requests
from bs4 import BeautifulSoup
import re


class NW():
    def __init__(self, staff='600299', staff_list=None, dls=False):
        # 选择传参或者DB读取工号列表
        if staff_list:
            self.staff_list = staff_list
        else:
            self.staff_list = self.db_list()
        self.staff = self.staff_list.get(staff)

        # 选择内网或者代理商网址并模拟cookie
        if dls:
            self.url = 'http://10.148.251.144'
            self.cookie = {'dx_dl': f'id={self.staff[3]}&work_no={staff}'}
        else:
            self.url = 'http://www.tc.net'
            self.cookie = {'tcdx_admin': f'id={self.staff[2]}&work_no={staff}'}
        # 模拟浏览器头
        self.header = {
            'User-Agent': ('Mozilla/5.0 (Windows NT 6.1; WOW64) '
                           'AppleWebKit/537.36 (KHTML, like Gecko) '
                           'Chrome/80.0.3987.132 Safari/537.36')}

    def db_list(self):
        # 从DB获取工号列表
        import cx_Oracle
        scb6 = cx_Oracle.connect('luyu', 'meafe123', 'tcscb6')
        cur = scb6.cursor()
        sql = cur.execute('''select staff_no,
                                    staff_nm,
                                    phone,
                                    nwid,
                                    dlsid
                               from ly_sys_user_table''').fetchall()
        staff_list = dict([[x[0], x[1:]] for x in sql])
        cur.close()
        scb6.close()
        return staff_list

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
                           cookies=self.cookie,
                           headers=self.header)
        if re.search('查看信息', rep.text):
            return True
        else:
            return False

    def nw_zhuanfa(self,
                   receivers,
                   title,
                   content,
                   shouid,
                   fjid=''):
        '''转发函数
           receivers  接收工号列表
           title      发送信息标题
           content    发送信息内容
           shouid     个人事务接收信息ID
           fjid       附件ID,字符串类型,多个ID以|分隔
        '''
        # 获取2个参数值
        v, e = self.get_param(f'{self.url}/grsw/shou/zhuanfa.aspx?id={shouid}')
        # 接收人列表
        persons = [f'[{x}]{self.staff_list.get(x)[0]}' for x in receivers]
        # 构造发送信息
        data = {
            '__VIEWSTATE': v,
            '__EVENTVALIDATION': e,
            'HiddenField2': ';'.join(persons),
            'TextBox2': title,
            'editor1$TextBox1': content,
            'Button1': '发 送'
        }
        # 转发的附件ID，2个网址不同处理方式
        if fjid:
            fjlist = fjid.split('|')
            if self.url == 'http://www.tc.net':
                for f in fjlist:
                    data[f'fjid_{f}'] = f
            else:
                data['fj_id'] = fjlist
        # 转发信息
        rep = requests.post(f'{self.url}/grsw/shou/zhuanfa.aspx?id={shouid}',
                            data=data,
                            cookies=self.cookie,
                            headers=self.header)
        if re.search('发送成功', rep.text):
            return True
        else:
            return False

    def nw_fa(self,
              receivers,
              title,
              content,
              filename=''):
        '''发送函数
           receivers  接收工号列表
           title      发送信息标题
           content    发送信息内容
           filename   发送文件
        '''
        # 获取2个参数值
        v, e = self.get_param(f'{self.url}/grsw/fa/add.aspx')
        # 接收人列表
        persons = [f'[{x}]{self.staff_list.get(x)[0]}' for x in receivers]
        # 构造发送信息
        post_data = {
            '__VIEWSTATE': v,
            '__EVENTVALIDATION': e,
            'HiddenField2': ';'.join(persons),
            'TextBox2': title,
            'editor1$TextBox1': content,
            'Button1': '发 送'
        }
        # 读取发送文件
        if filename:
            with open(filename, 'rb') as f:
                file = {'undefined': (filename, f.read())}
        else:
            file = ''
        # 发送信息
        rep = requests.post(f'{self.url}/grsw/fa/add.aspx',
                            data=post_data,
                            files=file,
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
