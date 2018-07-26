from bs4 import BeautifulSoup
import requests
import re
import pandas as pd


def str2num(func):
    def wrapper(*args, **kwargs):
        ret = func(*args, **kwargs)
        try:
            return float(ret)
        except ValueError:
            return ret
    return wrapper


@str2num
def find_bj(table, restr):
    name = table('tr')[0]('td')
    value = table('tr')[1]('td')
    for k, v in zip(name, value):
        if re.search(restr, k.text):
            return v.text
    else:
        return '0'


def spider_money(begin, end, mm=1, desc=None, addr='http://www.tc.net'):
    out = pd.DataFrame(
        columns=['staff', 'name', 'desc', 'money', 'tax', 'cut', 'gjj'])
    for gh in range(begin, end + 1):
        cookie = {'tcdx_admin': 'work_no={0}'.format(gh)}
        url = '%s/gzcx/index.asp' % addr
        rep = requests.get(url, cookies=cookie)
        rep.encoding = 'utf-8'
        soup = BeautifulSoup(rep.text, 'html.parser')

        one = {}
        cnt = 0
        for t in soup.html.body('table', recursive=False):
            one['desc'] = t.previous_sibling.string.strip()
            if desc and not re.search(desc, one['desc']):
                continue
            one['name'] = find_bj(t.table, '姓名')
            one['staff'] = find_bj(t.table, '工号')
            one['money'] = float(t('td', text=re.compile('\d'))[-1].string)
            one['tax'] = find_bj(t.table, '税')
            one['cut'] = find_bj(t.table, '扣款合计')
            one['gjj'] = find_bj(t.table, '公积金')
            out = out.append(one, ignore_index=True)
            cnt = cnt + 1
            if cnt >= mm:
                break

    out['name'] = out.merge(out.drop_duplicates('staff'), on='staff')['name_y']
    excel = pd.ExcelWriter('2018.xlsx')
    hz = out.groupby(['staff', 'name']).sum()
    hz['rank'] = hz.money.rank(method='first', ascending=False)
    hz.to_excel(excel, sheet_name='汇总')
    out.to_excel(excel, index=None, float_format='%g', sheet_name='清单')
    excel.save()


spider_money(600299, 600303, mm=100, desc='2018')
