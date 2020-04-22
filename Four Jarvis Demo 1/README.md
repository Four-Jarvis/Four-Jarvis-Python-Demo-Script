## Four Jarvis Auto Workflow Python Demo Script 1

简单的自动数据报告工作流Demo：从网页提取所需数据，自动生成舆情报告，定时邮件发送

简介：从东方财富网资讯频道，根据自定义关键词（公司名称），通过selenium采集数据，用re正则表达式清洗数据，提取新闻标题、时间、链接、新闻摘要、公司名称一起存入Mysql数据库，再生成数据报告，定时邮件发送

目的：拒绝无脑重复工作

流程：数据采集、数据清洗、数据储存、数据导出、定时邮件

##### 数据采集 Tool: selenium/谷歌浏览器/ChromeDriver

##### 数据清洗 Tool: 正则表达式re

##### 数据储存 Tool: 数据库mysql/pymysql

##### 数据导出 Tool: pymysql/pandas/matplotlib/python-docx/openpyxl/python-pptx

##### 定时邮件 Tool: SMTP/schedule

##### 其他工具 Tool: 多进程multiprocessing.Pool/日志处理loguru/python内置库

##### 实现过程

def selenium(): 传入url，传回网页源码

def eastmoney2mysql(): 通过正则表达式提取网页数据，并清洗导入mysql数据库

class Fmysql: 自定义mysql常用语句的类，执行sql语句

def get_Ystd() / def get_Today(): 生成昨天/今天的日期字符串

def save2picture(): 传入日期和文件保存路径，通过pandas的read_sql()方法生成数据，进而通过matplotlib生成图片

def save2excel():  / def save2docx(): / def save2pptx(): 传入日期/公司/文件保存路径，调用Fmysql类，执行sql语句生成数据，写入excel/word/ppt

def create_email(): 传入发件人昵称/收件人昵称/主题/正文/附件地址,附件名称，生成一封邮件

def send_email(): 传入发件人邮箱账号/密码/收件人邮箱/邮件内容，发送邮件

def main_by_schedule():自定义定时函数

def main():  实现以上函数，通过multiprocessing开启多进程实现数据采集，最后生成数据报告和发送邮件，再通过自定义定时函数执行main函数

##### 合法性说明

本DEMO是为了学习研究的目的，所有读者可以参考执行思路和程序代码，但不能用于恶意和非法目的（恶意攻击网站服务器、非法盈利等），如有违者请自行负责。
本DEMO所获取的数据都是仅用于个人学习，并非于恶意抓取数据来攫取不正当竞争的优势，也未用于商业目的牟取不法利益，并非大范围地爬取，同时严格控制爬取的速率、争取不对服务器造成压力，如侵犯当事者（即被抓取的网络主体）的利益，请联系更改或删除。

##### 本DEMO环境配置

Python：3.6
Mysql：8.0
Google Chrome版本：81.0.4044.113
ChromeDriver：81.0.4044.69
Jupyter Notebook/ Visual Studio Code 

#### 微信公众号：Four Jarvis

##### DEMO图片



![image-20200423000745075](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four%20Jarvis%20Demo%201/demo_report_img/image-20200423000745075.png](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four Jarvis Demo 1/demo_report_img/image-20200423000745075.png)

![image-20200423000620695](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four%20Jarvis%20Demo%201/demo_report_img/image-20200423000620695.png](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four Jarvis Demo 1/demo_report_img/image-20200423000620695.png)

![image-20200423000816556](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four%20Jarvis%20Demo%201/demo_report_img/image-20200423000816556.png](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four Jarvis Demo 1/demo_report_img/image-20200423000816556.png)

![image-20200423000842771](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four%20Jarvis%20Demo%201/demo_report_img/image-20200423000842771.png](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four Jarvis Demo 1/demo_report_img/image-20200423000842771.png)

![image-20200423000921325](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four%20Jarvis%20Demo%201/demo_report_img/image-20200423000921325.png](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four Jarvis Demo 1/demo_report_img/image-20200423000921325.png)

![image-20200423000936807](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four%20Jarvis%20Demo%201/demo_report_img/image-20200423000936807.png](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four Jarvis Demo 1/demo_report_img/image-20200423000936807.png)

![image-20200423000951586](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four%20Jarvis%20Demo%201/demo_report_img/image-20200423000951586.png](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four Jarvis Demo 1/demo_report_img/image-20200423000951586.png)

![image-20200423001007731](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four%20Jarvis%20Demo%201/demo_report_img/image-20200423001007731.png](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four Jarvis Demo 1/demo_report_img/image-20200423001007731.png)

![image-20200423001111039](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four%20Jarvis%20Demo%201/demo_report_img/image-20200423001111039.png](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four Jarvis Demo 1/demo_report_img/image-20200423001111039.png)

![image-20200423001159681](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four%20Jarvis%20Demo%201/demo_report_img/image-20200423001159681.png](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four Jarvis Demo 1/demo_report_img/image-20200423001159681.png)

![Four Jarvis Auto Demo1 2020-04-22](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four%20Jarvis%20Demo%201/demo_report_img/image-20200423001323079.png](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four Jarvis Demo 1/demo_report_img/image-20200423001323079.png)

![image-20200423015125201](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four%20Jarvis%20Demo%201/demo_report_img/image-20200423015125201.png](https://github.com/Four-Jarvis/Four-Jarvis-Python-Demo-Script/blob/master/Four Jarvis Demo 1/demo_report_img/image-20200423015125201.png)
