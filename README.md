# 微信分析 WeChatAnalysis

使用python进行微信好友或群聊的分析(Using Python to analyze WeChat friends or group chat)

## 微信好友分析

### 1. 博客地址：[博客园](https://www.cnblogs.com/bpf-1024/p/10962332.html)

### 2. 编程环境：python3.7

### 3. 使用的库

> os, wxpy, openpyxl, pyecharts, wordcloud, matplotlib, jieba
>
> 图库包：(需安装后重启)
> > echarts-china-cities-pypkg&emsp;&emsp; `0.0.9`
> >
> > echarts-china-provinces-pypkg `0.0.3`
> >
> > echarts-countries-pypkg&emsp;&emsp;&emsp; `0.1.6`

### 4. 运行结果

> /out/before 2019-6-3的运行结果
>
> /out/bpf    2019-7-19的运行结果

&emsp;&emsp;输出文件包含一张关于市级的词云图、一个全国分布的网页地图、一个广东省分布的网页地图、一个微信好友信息的Excel表格（这个涉及个人隐私，我就不放了）

## 微信群聊好友分析

### 1. 编程环境：python3.7

### 2. 功能
&emsp;&emsp;获取微信群聊所有成员的备注和群昵称，并输出到excel表格，文件以群聊名称命名。

### 3. 使用的库

> os, wxpy, xlwt

### 4. 使用前提

> 微信能够登陆网页版
>
> 群聊添加至通讯录

### 5. 使用步骤

> 5.1 运行WeChatGroupMember.py
>
> 5.2 第一次运行需要下载安装外部函数库，等待安装完成
>
> 5.3 登陆微信，扫描二维码，此时电脑端的微信会自动退出
>
> 5.4 选择群聊，通过序号选择即可
>
> 5.5 等待程序运行完成

### 6. 运行效果

> /out/WeChat_group_member.mp4 记录了一次运行的过程
