# 简介
通过爬虫批量从https://yts.mx/下载1080p电影种子
# 需要导入的包
```
import time
import requests
import lxml.html
import re
import os
import xlrd
from multiprocessing.dummy import Pool
import openpyxl

```
# 需要准备的文件
+ excel表格：内容为电影名。如图所示：

![](http://typora.joyboy2.top/image/20210613222457.png)

# 注意事项
+ 因为我的Excel中电影名都是从豆瓣直接爬取的，所以电影名要满足中文名+英文名+年份的要求。有自己特殊的格式的话可以在get_colum()方法自己改
+ 有些种子不知道为什么为0kb，有能力的可以自己解决
+ 200个线程太少的话，可以自己加，不过因为https://yts.mx/为国外网站，所以网速可能比较慢，建议用梯子
+ 重要的地方都写了注释，有需要的可以自己改改
