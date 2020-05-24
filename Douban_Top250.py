'''
file: Douban.movies.py
encoding: utf-8
data: 2019.12.31
name: xxx
email: xxx
introduction: 爬取豆瓣热门电影数据，存入excel文件，并统计数据生成图表和词云
url: https://movie.douban.com/explore#!type=movie&tag=%E7%83%AD%E9%97%A8&sort=recommend&page_limit=20&page_start=0
'''


import requests
import json
import re
import time
import xlwt
import xlrd
import matplotlib.pyplot as plt
import matplotlib as mpl
import numpy as np
from wordcloud import WordCloud
import PIL.Image as image


def respone(url):

    # 模拟谷歌浏览器请求头
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36"}
    # 发请求并获取响应
    res = requests.get(url, headers=headers)
    # 将响应的编码格式设置为utf-8
    res.encoding = 'utf-8'
    try:
        # json格式转码
        html = json.loads(res.text)
    except json.decoder.JSONDecodeError:
        # 若不是json格式则直接获取html源码
        html = res.text
    
    return html


# 正则匹配影片外国名
def re_foreign_name(html):

    name = re.findall(r'<span property="v:itemreviewed">(.+?)</span>', html)[0]
    name = name.split(' ')
    name.pop(0)
    foreign_name = ''
    for ch in name:
        if ch == name[-1]:
            foreign_name += ch
        else:
            foreign_name += (ch + ' ')
    
    return foreign_name  # 影片外国名


# 正则匹配评价数
def re_evaluation(html):
    
    evaluation = re.findall(r'property="v:votes">(.+?)</span>', html)[0]
    
    return evaluation  # 评价数


# 正则匹配概况
def re_introduction(html):

    introduction = re.findall(r'v:summary" class="">(.+?)</span>', html, re.S)[0]
    introduction = introduction.replace(' ','').replace('\n','')[2:]
    
    return introduction  # 概况


# 正则匹配导演
def re_director(html):
    
    director = re.findall(r'v:directedBy">(.+?)</a>', html, re.S)[0]
    
    return director  # 导演


# 正则匹配主演
def re_actors(html):

    actors_data = re.findall(r'v:starring">(.+?)</a>', html, re.S)
    actors = ''
    for actor in actors_data:
        if actor == actors_data[-1]:
            actors += actor
        else:
            actors += (actor + ' | ')
    
    return actors  # 主演


# 正则匹配年份
def re_years(html):

    years_data = re.findall(r'initialReleaseDate" content="(.+?)"', html)
    years = ''
    for year in years_data:
        if year == years_data[-1]:
            years += year
        else:
            years += (year + ' | ')
    
    return years  # 年份


# 正则匹配地区
def re_region(html):

    region = re.findall(r'\u5236\u7247\u56fd\u5bb6\u002f\u5730\u533a:</span>(.+?)<br/>', html)[0]
    region = region.replace(' ','')
    
    return region  # 地区


# 正则匹配类别
def re_types(html):

    types_data = re.findall(r'v:genre">(.+?)</span>', html)
    types = ''
    for Type in types_data:
        if Type == types_data[-1]:
            types += Type
        else:
            types += (Type + ' | ')

    return types  # 类别


# 将匹配到的文件写入文件中
def write_data(movie_datas):

    workbook = xlwt.Workbook(encoding = 'utf-8')  # 创建写入文件
    worksheet = workbook.add_sheet('movies')  # 创建工作表

    line1 = ['排名','影片中文名','评分','评价数','导演','主演','年份','地区','类别']
    for index,col_name in enumerate(line1):
        worksheet.write(0,index, label = col_name)  #写入第一行

    line = 1
    rates = []  # 存所有评分
    names = []  # 存所有电影名
    urls = []  # 存所有电影详情链接
    evaluations = []  # 存所有的评价数

    for movie_data in movie_datas:
        # 写排名
        worksheet.write(line, 0, label=str(line))
        for index,data in enumerate(movie_data):
            # 写电影所有数据
            worksheet.write(line, index+1, label=data)
        line += 1  # 每写一行，行数+1
    
    workbook.save('Douban_movies.xls')  # 保存写入文件


# 从文件中获取数据
def read_data():

    # 打开数据文件
    data = xlrd.open_workbook('Douban_movies.xls')
    table = data.sheets()[0]  # 操作工作表一
    
    # 获取第5列数据
    directors = table.col_values(4)[1:]
    director_count = {}
    for director in directors:
        # 得key为导演，value为负责的电影数量的字典
        director_count[director] = director_count.get(director, 0) + 1

    actors = table.col_values(5)[1:]
    actor_count = {}
    for actor in actors:
        actor_list = actor.split(' | ')
        for actor in actor_list:
            # 得key为主演，value为参演的电影数量的字典
            actor_count[actor] = actor_count.get(actor, 0) + 1

    years = table.col_values(6)[1:]
    year_count = {}
    for year in years:
        year_list = year.split(' | ')
        #for year in year_list:
        year = year_list[0]
        year = year[:4]
        if int(year) <= 1980:
            year = '1980之前'
        elif 1980 < int(year) <= 1990:
            year = '1981-1990'
        elif 1990 < int(year) <= 2000:
            year = '1991-2000'
        elif 2000 < int(year) <= 2010:
            year = '2001-2010'
        elif 2010 < int(year) <= 2020:
            year = '2011-2020'
        # 得key为年代，value为上映的电影数量的字典
        year_count[year] = year_count.get(year, 0) + 1

    return director_count,actor_count,year_count


# 生成导演的柱状图
def bar_director(director_count):

    plt.rcParams["font.sans-serif"]=["SimHei"]  #正常显示中文标签
    plt.rcParams["axes.unicode_minus"]=False

    x = np.arange(10)  # x轴
    xticks = []  # x轴说明文字
    y = []

    items = list(director_count.items())
    items.sort(key=lambda x:x[1], reverse=True)
    for name,count in items[:10]:
        xticks.append(name)
        y.append(count)

    fig = plt.figure(figsize=(15,12))  # 控制画布大小
    
    bar_width = 0.3  # 控制条形柱宽度
    # align控制条形柱位置
    # color控制条形柱颜色
    # label为该条形柱对应图示
    # alpha控制条形柱透明度
    plt.bar(x, y, bar_width, align="center", color="c", alpha=0.5)
    # 控制x轴显示什么
    plt.xticks(x, xticks, rotation=-30)
    # 控制x轴刻度距离
    plt.xlim(-0.5, 10)
    plt.ylim(0, 8)

    ax = fig.add_subplot(1, 1, 1)
    ax.set_ylabel('电影数')

    for a, b in zip(x, y):
        plt.text(a, b, '%d'%b, ha='center', va='bottom', fontsize=10)

    plt.title('十大最佳导演',fontsize=20)  # 添加标题
    
    #plt.show()  # 是否打开图形输出器，如有需要再打开
    plt.savefig('十大最佳导演.png')  # 保存为当前目录下的图片
    plt.clf()  #关闭图形编辑器


# 生成主演的柱状图
def bar_actor(actor_count):

    plt.rcParams["font.sans-serif"]=["SimHei"]  #正常显示中文标签
    plt.rcParams["axes.unicode_minus"]=False

    x = np.arange(10)  # x轴
    xticks = []  # x轴说明文字
    y = []

    items = list(actor_count.items())
    items.sort(key=lambda x:x[1], reverse=True)
    for name,count in items[:10]:
        xticks.append(name)
        y.append(count)

    fig = plt.figure(figsize=(15,12))  # 控制画布大小
    
    bar_width = 0.3  # 控制条形柱宽度
    # align控制条形柱位置
    # color控制条形柱颜色
    # label为该条形柱对应图示
    # alpha控制条形柱透明度
    plt.bar(x, y, bar_width, align="center", color="c", alpha=0.5)
    # 控制x轴显示什么
    plt.xticks(x, xticks, rotation=-30)
    # 控制x轴刻度距离
    plt.xlim(-0.5, 10)
    plt.ylim(0, 9)

    ax = fig.add_subplot(1, 1, 1)
    ax.set_ylabel('出演电影数')

    for a, b in zip(x, y):
        plt.text(a, b, '%d'%b, ha='center', va='bottom', fontsize=10)

    plt.title('十大最佳主演',fontsize=20)  # 添加标题
    
    #plt.show()  # 是否打开图形输出器，如有需要再打开
    plt.savefig('十大最佳主演.png')  # 保存为当前目录下的图片
    plt.clf()  #关闭图形编辑器


# 生成年代的柱状图
def bar_year(year_count):

    plt.rcParams["font.sans-serif"]=["SimHei"]  #正常显示中文标签
    plt.rcParams["axes.unicode_minus"]=False

    x = np.arange(len(year_count))  # x轴
    xticks = []  # x轴说明文字
    y = []

    items = list(year_count.items())
    # 按年代由小到大排序
    items.sort(key=lambda x:x[0][:4], reverse=False)

    for name,count in items:
        xticks.append(name)
        y.append(count)

    fig = plt.figure(figsize=(10,10))  # 控制画布大小
    
    bar_width = 0.3  # 控制条形柱宽度
    # align控制条形柱位置
    # color控制条形柱颜色
    # label为该条形柱对应图示
    # alpha控制条形柱透明度
    plt.bar(x, y, bar_width, align="center", color="c", alpha=0.5)
    # 控制x轴显示什么
    plt.xticks(x, xticks, rotation=0)

    ax = fig.add_subplot(1, 1, 1)
    ax.set_ylabel('上映电影数')

    for a, b in zip(x, y):
        plt.text(a, b, '%d'%b, ha='center', va='bottom', fontsize=10)

    plt.title('电影上映数年代分布图',fontsize=20)  # 添加标题
    
    #plt.show()  # 是否打开图形输出器，如有需要再打开
    plt.savefig('电影上映数年代分布图.png')  # 保存为当前目录下的图片
    plt.clf()  #关闭图形编辑器


# 生成词云
def word_cloud(actor_count):

    items = list(actor_count.items())
    items.sort(key=lambda x:x[1], reverse=True)
    
    word_list = []
    for item in items:
        word_list.append(item[0])

    text = " ".join(word_list)

    # 打开词云背景图
    mask = np.array(image.open("wordcloud.jpg"))
    # 生成词云图
    wordcloud = WordCloud(mask=mask, font_path="C:\\Windows\\Fonts\\msyh.ttc").generate(text)
    # 保存词云图
    wordcloud.to_file("词云.jpg")


# 主循环
def main():
    
    movie_counts = int(input('\n>>> 请输入你想获取的电影数量：'))
    start = time.time()  # 爬取开始时间
    print('\n>>> 开始爬取...预计耗时：%d 秒\n'%(movie_counts*2.5))
    url = 'https://movie.douban.com/j/search_subjects?type=movie&tag=%E5%8F%AF%E6%92%AD%E6%94%BE&sort=rank&page_limit={}&page_start=0'.format(movie_counts)
    
    movie_datas = []  # 存所有电影信息
    failure = 0  # 失败电影数
    
    html = respone(url)
    for datas in html["subjects"]:
        try:
            movie_data = []  # 存某电影所有信息
            movie_data.append(datas['title'])  # 影片中文名
            url = datas['url']
            html = respone(url)
            movie_data.append(datas['rate'])  # 评分
            evaluation = re_evaluation(html)
            movie_data.append(evaluation)  # 评价数
            director = re_director(html)
            movie_data.append(director)  # 导演
            actors = re_actors(html)
            movie_data.append(actors)  # 主演
            years = re_years(html)
            movie_data.append(years)  # 年份
            region = re_region(html)
            movie_data.append(region)  # 地区
            types = re_types(html)
            movie_data.append(types)  # 类别
            movie_datas.append(movie_data)
            print('>>> 链接：%s    电影名称：%s ---> 爬取成功'%(datas['url'],datas['title']))

        except Exception as error:
            print('>>> 链接：%s    电影名称：%s ---> 爬取失败'%(datas['url'],datas['title']))
            print('>>> 失败原因：',error)
            failure += 1
        
        time.sleep(2)  # 每次爬取睡眠2秒

    end = time.time()  # 爬取结束时间
    times = end - start
    print('\n>>> 全部电影爬取完毕，成功爬取 %d 部，失败 %d 部'%(movie_counts-failure, failure))
    print('    总耗时：%d 秒\n'%times)
    
    print('>>> 开始将全部爬取的数据写入excel文件中...')
    write_data(movie_datas)
    print('    数据已全部写入excel文件中 ---> Douban_movies.xls\n')
    
    print('>>> 开始进行数据可视化...')
    director_count,actor_count,year_count = read_data()
    bar_director(director_count)
    print('>>> 已统计出上映电影次数最多的十位导演，并保存为图片 ---> 十大最佳导演.png')
    bar_actor(actor_count)
    print('>>> 已统计出参演电影次数最多的十位主演，并保存为图片 ---> 十大最佳主演.png')
    bar_year(year_count)
    print('>>> 已统计出各个年代电影数量，并保存为图片 ---> 电影上映数年代分布图.png\n')
    
    print('>>> 开始根据电影类型生成词云图片...')
    word_cloud(actor_count)
    print('>>> 词云图片已生成完毕，保存为 ---> 词云.jpg\n')
    
   
if __name__ == '__main__':

    main()