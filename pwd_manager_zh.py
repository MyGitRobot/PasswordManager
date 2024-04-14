# -*- coding: utf-8 -*-

"""
Program: 简易密码管理器
Author: MrCrawL
Created Time: 2024-04-11
Last Modified: 2024-04-13
Modified by: MrCrawL
"""

from warnings import simplefilter
from configparser import ConfigParser
import pandas as pd
import os, sys, logging

simplefilter('ignore', FutureWarning)


class PwdManager:
    def __init__(self, file_dir:str):
        self.fileDir = file_dir
        if os.path.exists(self.fileDir):
            self.df = pd.read_excel(self.fileDir, 'Sheet1', index_col=0)
            self.df.reset_index(drop=True, inplace=True)
        else:
            data = {
                'address': [],
                'account': [],
                'password': []
            }
            self.df = pd.DataFrame(data)
            self.df.to_excel(self.fileDir, index=False)
        print('[Info] 数据加载中...')

    def ask_action(self):
        action = input('>>> 请问您想 [1]查询 [2]添加 [3]删除 [4]修改 [0]退出: ')
        if action == '1':
            self.query_pwd()
        elif action == '2':
            self.add_pwd()
        elif action == '3':
            self.del_pwd()
        elif action == '4':
            self.chg_pwd()
        elif action == '0' or action == '':
            return None
        else:
            print('[Error] 请输入正确的序号.')
        self.ask_action()

    def query_pwd(self):
        words = input('>>> 请输入要查询的地址: ').lower()
        if words:
            self.match_words(words)
        else:
            print(self.df)

    def match_words(self, words:str):
        matches = []
        for i in range(len(self.df)):
            if words in self.df.iloc[i, 0].lower():
                matches.append(i)
        if matches:
            for match in matches:
                print(
                    f'[{match}] 地址: {self.df.iloc[match, 0]}  '
                    f'账号: {self.df.iloc[match, 1]}  '
                    f'密码: {self.df.iloc[match, 2]}')
            return matches
        else:
            print(f'[Info] 您输入的地址不存在.')
            return None

    def add_pwd(self):
        address = input('>>> 请输入地址: ')
        account = input('>>> 请输入账号: ')
        password = input('>>> 请输入密码: ')
        self.df.loc[len(self.df)] = [address, account, password]
        self.save()

    def del_pwd(self):
        words = input('>>> 请输入要删除的地址: ').lower()
        matches = self.match_words(words)
        if matches:
            index = input('>>> 请选择要删除的条目: ')
            if index == '' and len(matches) == 1:
                ensure = input(f'>>> 确认删除 {self.df.iloc[matches[0], 0]} ? (y/n): ').lower()
                if ensure == 'y' or ensure == '':
                    self.df.drop(matches[0], axis=0, inplace=True)
                    print('[Info] 删除成功.')
                    self.df.reset_index(drop=True, inplace=True)
                    self.save()
                else:
                    print('[Info] 取消删除.')
            else:
                self.del_index(index, matches)

    def del_index(self, index:str, matches:list):
        if index.isdigit():
            trush_index = int(index)
            if trush_index in matches:
                ensure = input(f'>>> 确认删除 {self.df.iloc[trush_index, 0]} ? (y/n): ').lower()
                if ensure == 'y' or ensure == '':
                    self.df.drop(trush_index, axis=0, inplace=True)
                    print('[Info] 删除成功.')
                    self.df.reset_index(drop=True, inplace=True)
                    self.save()
                else:
                    print('[Info] 取消删除.')
            else:
                trush_index = input('>>> 请输入正确的序号: ')
                self.del_index(trush_index, matches)
        else:
            trush_index = input('>>> 请输入正确的序号: ')
            self.del_index(trush_index, matches)

    def chg_pwd(self):
        chg_address = input('>>> 请输入要修改密码的地址: ').lower()
        matches = self.match_words(chg_address)
        if matches:
            index = input('>>> 请选择要修改密码的条目: ')
            if index == '' and len(matches) == 1:
                password = input('>>> 请输入新密码: ')
                ensure = input(f'>>> 确认修改密码为 {password} ?(y/n): ').lower()
                if ensure == 'y' or ensure == '':
                    self.df.iloc[matches[0], 2] = password
                    print('[Info] 修改成功.')
                    self.save()
                else:
                    print('[Info] 取消修改.')
            else:
                self.chg_index(index ,matches)

    def chg_index(self, index:str, matches:list):
        if index.isdigit():
            chg_index = int(index)
            if chg_index in matches:
                password = input('>>> 请输入新密码: ')
                ensure = input(f'>>> 确认修改密码为 {password} ?(y/n): ').lower()
                if ensure == 'y':
                    self.df.iloc[chg_index, 2] = password
                    print('[Info] 修改成功.')
                    self.save()
                else:
                    print('[Info] 取消修改.')
            else:
                chg_index = input('>>> 请输入正确的序号: ')
                self.chg_index(chg_index, matches)
        else:
            chg_index = input('>>> 请输入正确的序号: ')
            self.chg_index(chg_index, matches)

    def save(self):
        with pd.ExcelWriter(self.fileDir) as wt:
            self.df.to_excel(wt, 'Sheet1')
        print('[Info] 保存成功.')


def first_ues():
    if os.path.exists('config.ini'):
        return False
    else:
        if getattr(sys, 'frozen', False):
            fileDir = os.path.join(os.path.dirname(sys.executable), 'myPassword.xlsx')
        else:
            fileDir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'myPassword.xlsx')
        print("[Info] 这是您第一次运行程序，系统将自动创建config.ini文件，您可以在该文件修改密码保存的位置.")
        with open('config.ini', 'w', encoding='utf-8') as file:
            file.write(f'[config]\nfile_dir = {fileDir}\n')
        print("[Info] 如需修改，快去修改吧！")
        input('>>> (按下回车退出程序...)')
        return True


def main():
    if not first_ues():
        config = ConfigParser()
        config.read('config.ini', encoding='utf-8')
        fileDir = config['config']['file_dir']
        print(f'[Info] 文件位置: {fileDir}')
        mng = PwdManager(fileDir)
        mng.ask_action()


def setup_logging():
    logging.basicConfig(filename='error.log', level=logging.ERROR,
                        format='%(asctime)s - %(levelname)s - %(message)s',
                        encoding='utf-8')


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'[Error] {e}')
        setup_logging()
        logging.error(f"Error occurred: {e}", exc_info=True)