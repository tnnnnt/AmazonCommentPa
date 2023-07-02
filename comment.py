import os
import re
import openpyxl
import pandas as pd
import requests
import urllib3.contrib
import winsound
from lxml import etree
from tqdm import tqdm
import time

# 用户设定
id = "B078K3B7HY"   # 商品的id
excel_name = "test.xlsx"    # 保存的文件名
save_path = "E:"    # 保存路径
timeout = 15.05     # 请求超时设置，单位为秒，建议设置为略大于3的倍数，如3.05、6.05、9.05等，网络状态差的话建议长一点
sleep_time = 5      # 睡眠时间设置，单位为秒，太短容易被反爬机制检测到，太长就效率低
headers = {}  # 用你自己的


# 开发者设定
def get_response(URL):
    while True:
        # noinspection PyBroadException
        try:
            urllib3.disable_warnings()
            urllib3.contrib.pyopenssl.inject_into_urllib3()
            requests.packages.urllib3.disable_warnings()
            requests.DEFAULT_RETRIES = 10
            requests.session().keep_alive = False
            response = requests.get(URL, headers=headers, verify=False, timeout=timeout)
            response.close()
            return response
        except Exception:
            time.sleep(sleep_time)


def get_html(URL):
    return get_response(URL).text


def change_time(t):
    t = t.split(' ')
    return t[2] + '/' + months[t[0]] + '/' + t[1][:-1]


def change_num(a):
    return a.replace(',', '')


tmp_str = "https://www.amazon.com/dp/"
get_types_html = get_html(tmp_str + id)
tmp_str = '"dimensionValuesDisplayData" : (.*?)]},'
try:
    types = eval(re.findall(tmp_str, get_types_html)[0] + ']}')
    tmp_str = '"dimensionsDisplay" : (.*?)],'
    dimensions = eval(re.findall(tmp_str, get_types_html)[0] + ']')
except IndexError:
    types = {id: []}
    dimensions = []

first_row = ['日期', '地区', '用户名', '评价标题', '评价内容', '评价星级']
for dim in dimensions:
    first_row.append(dim)

while os.path.isfile(save_path + '//' + excel_name):
    excel_name = excel_name.split('.')[0] + '_.xlsx'
os.chdir(save_path)
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.append(first_row)

str1 = "https://www.amazon.com/product-reviews/"
str2 = "/ref?filterByStar="
str3 = "&sortBy="
str4 = "&reviewerType="
str5 = "&mediaType="
str6 = "&formatType=current_format"
str7 = 'total rating.*? (.*?) with review'
str8 = "&pageNumber="
str9 = '/html/body/div[1]/div[3]/div/div[1]/div/div[1]/div[5]/div[3]/div/div['
str10 = ']/div/div/div[1]/a/div[2]/span'
str11 = ']/div/div/div[4]/span/span/text()'
str12 = 'Reviewed in (.*?) on (.*?)114514'
str13 = ']/div/div/span'
str14 = '114514'
str15 = '/html/body/div[1]/div[3]/div/div[1]/div/div[1]/div[5]/div[3]/div//div/div/div[2]/a/span[2]'
str16 = '/html/body/div[1]/div[3]/div/div[1]/div/div[1]/div[5]/div[3]/div//div/div/div[2]/span[2]/span'
str17 = ']/div/div/div[1]/div/div[2]/span'
str18 = ']/div/div/div[1]/div[1]/div/a/div[2]/span'

months = {'January': '1', 'February': '2', 'March': '3', 'April': '4', 'May': '5', 'June': '6', 'July': '7',
          'August': '8', 'September': '9', 'October': '10', 'November': '11', 'December': '12'}
filterByStar = ['one_star', 'two_star', 'three_star', 'four_star', 'five_star']
sortBy = ['helpful', 'recent']
reviewerType = ['all_reviews', 'avp_only_reviews']
mediaType = ['all_contents', 'media_reviews_only']

for k in tqdm(types):
    for star in range(5):
        fbs = filterByStar[star]
        users, titles, comments, locations, times = [], [], [], [], []
        for sb in sortBy:
            for rt in reviewerType:
                for mt in mediaType:
                    base_url = str1 + k + str2 + fbs + str3 + sb + str4 + rt + str5 + mt + str6
                    print(base_url)
                    try:
                        pages = min(10, (int(change_num(
                            re.findall(str7, get_html(base_url))[0])) + 9) // 10)
                    except IndexError:
                        pages = 1
                    for p in range(1, pages + 1):
                        url = base_url + str8 + str(p)
                        html = get_html(url)
                        xhtml = etree.HTML(html)
                        for i in range(1, 13):
                            # noinspection PyBroadException
                            try:
                                users.append(xhtml.xpath(str9 + str(i) + str10)[0].text)
                            except Exception:
                                # noinspection PyBroadException
                                try:
                                    users.append(xhtml.xpath(str9 + str(i) + str17)[0].text)
                                except Exception:
                                    # noinspection PyBroadException
                                    try:
                                        users.append(xhtml.xpath(str9 + str(i) + str18)[0].text)
                                    except Exception:
                                        pass
                            # noinspection PyBroadException
                            try:
                                comments.append(xhtml.xpath(str9 + str(i) + str11)[0])
                            except Exception:
                                if len(comments) < len(users):
                                    comments.append("")
                            # noinspection PyBroadException
                            try:
                                data = re.findall(str12, xhtml.xpath(str9 + str(i) + str13)[0].text + str14)[0]
                                locations.append(data[0])
                                times.append(change_time(data[1]))
                            except Exception:
                                pass
                        datas = xhtml.xpath(str15)
                        for data in datas:
                            titles.append(data.text)
                        datas = xhtml.xpath(str16)
                        for data in datas:
                            if data.text is None:
                                continue
                            titles.append(data.text)
                    url = str1 + k + str2 + fbs + str3 + sb + str4 + rt + str5 + mt + str6
                    if int(change_num(re.findall(str7, get_html(url))[0])) <= 100:
                        break
                url = str1 + k + str2 + fbs + str3 + sb + str4 + rt + str6
                if int(change_num(re.findall(str7, get_html(url))[0])) <= 100:
                    break
            url = str1 + k + str2 + fbs + str3 + sb + str6
            if int(change_num(re.findall(str7, get_html(url))[0])) <= 100:
                break
        nums = len(users)
        for ind in range(nums):
            row = [times[ind], locations[ind], users[ind], titles[ind], comments[ind], star + 1]
            for add in types[k]:
                row.append(add)
            sheet.append(row)
workbook.save(excel_name)
pd.DataFrame(pd.read_excel(save_path + '\\' + excel_name, 'Sheet')).drop_duplicates().to_excel(
    save_path + '\\' + excel_name)
workbook = openpyxl.load_workbook(save_path + '\\' + excel_name)
sheet = workbook.active
sheet.delete_cols(idx=1)
workbook.save(save_path + '\\' + excel_name)
winsound.Beep(1000, 1000)
exit()
