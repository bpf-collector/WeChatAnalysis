# -*- coding: utf-8 -*-
'''
 * @Author       : bpf
 * @Date         : 2020-09-29 00:52:12
 * @Description  : 微信群组成员分析
 * @LastEditTime : 2020-09-29 02:14:11
'''
import os
try:
    import xlwt
except:
    os.system("pip install xlwt")
    import xlwt

try:
    import wxpy
except:
    os.system("pip install wxpy")
    import wxpy

# 登陆
bot = wxpy.Bot()

# 获取所有群聊
print("获取信息中...")
groups = bot.groups()

# 输出所有群聊
for i in range(len(groups)):
    print(i+1, groups[i])

# 选择群聊
while True:
    try:
        index = int(input("请选择数字："))
        if index > 0 and index <= len(groups):
            break
    except:
        print("输入有误，请重新输入：")
groupName = groups[index-1].__str__().split()[1][:-1]
group = groups.search(groupName)[0]

# 获取群聊所有成员的 备注、群昵称
nickName, displayName = [], []
for member in group:
    nickName.append(member.raw.get('NickName'))
    displayName.append(member.raw.get('DisplayName'))

# 创建文件
workbook = xlwt.Workbook(encoding="utf-8")

# 创建表格
worksheet = workbook.add_sheet(groupName)

# 表头
title = ["序号", "备注", "群昵称"]
# 写入表头
for i in range(len(title)):
    worksheet.write(0, i, title[i])

# 写入数据
for i in range(len(nickName)):
    worksheet.write(i+1, 0, i+1)
    worksheet.write(i+1, 1, nickName[i])
    worksheet.write(i+1, 2, displayName[i])

# 保存文件
savepath = os.getcwd() + "\\" + groupName + ".xls"
workbook.save(savepath)
print("文件已保存到", savepath)