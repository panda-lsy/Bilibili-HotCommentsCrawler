# Bilibili（B站）视频评论爬虫

````
使用bilibili评论api爬取多个视频评论，能更快更全面的爬取数据，比Selenium爬取更快，避免浏览器崩溃的问题，不需要安装额外的驱动。
````

**功能：**

1. 自动抓取热点视频、自动爬取评论。
2. 能够全面爬取评论和用户数据：能够爬取到一级评论和对应的二级评论，及评论相关数据（用户昵称、用户当前等级、用户性别、评论内容、被回复用户昵称、评论层级、评论点赞数量和回复时间）
3. 保存的文件都在result目录下，并整合到Excel文档中，不同的视频在不同的Sheet中。
4. 多线程支持，默认5线程：**有风控风险**

**安装：**

```

```

**requests**是 一个用于发送HTTP请求的库，非Python标准库，需要额外安装。 安装方法：打开终端或命令提示符，输入以下命令：

~~~cmd
pip install -r requirements.txt
~~~

**爬取结果示例：**

```
输出结果是UTF-8格式，如果打开是乱码请更换打开方式（pycharm，记事本等）或者检查编码格式。
```

![image](https://github.com/LSQYES/BilibiliCommentsCrawler/blob/main/example.png)
