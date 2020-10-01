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

class WeChatGroupMember:
    def __init__(self):
        # 登陆
        self.bot = wxpy.Bot()
        self.nickName = []      # 备注
        self.displayName = []   # 群昵称
    
    def getGroups(self):
        # 获取所有群聊
        print("获取信息中...")
        self.groups = self.bot.groups()

    def chooseGroup(self):
        # 输出所有群聊
        for i in range(len(self.groups)):
            print(i+1, self.groups[i])

        # 选择群聊
        while True:
            try:
                index = int(input("请选择数字："))
                if index > 0 and index <= len(self.groups):
                    break
            except:
                print("输入有误，请重新输入：")
        groupName = self.groups[index-1].__str__().split()[1][:-1]
        group = self.groups.search(groupName)[0]
        return group, groupName

    def getGroupMember(self, group):
        # 获取群聊所有成员的 备注、群昵称
        for member in group:
            self.nickName.append(member.raw.get('NickName'))
            self.displayName.append(member.raw.get('DisplayName'))

    def saveAsXls(self, groupName):
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
        for i in range(len(self.nickName)):
            worksheet.write(i+1, 0, i+1)
            worksheet.write(i+1, 1, self.nickName[i])
            worksheet.write(i+1, 2, self.displayName[i])

        # 保存文件
        savepath = os.getcwd() + "\\" + groupName + ".xls"
        try:
            workbook.save(savepath)
            print("文件已保存到", savepath)
        except Exception as e:
            print("[Error]: ", e)
    
    def run(self):
        self.getGroups()
        group, groupName = self.chooseGroup()
        self.getGroupMember(group)
        self.saveAsXls(groupName)

if __name__ == "__main__":
    wc = WeChatGroupMember()
    wc.getGroups()
    while True:
        group, groupName = wc.chooseGroup()
        wc.getGroupMember(group)
        wc.saveAsXls(groupName)

        next = input("是否继续？(Y/N)")
        if next.lower() == 'n':
            break
