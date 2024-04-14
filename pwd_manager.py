# -*- coding: utf-8 -*-

"""
Program: Simple Password Manager
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
        print('[Info] Loading data...')

    def ask_action(self):
        action = input('>>> You want to [1]Query [2]Add [3]Delete [4]Change [0]Exit: ')
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
            print('[Error] Please input correct number.')
        self.ask_action()

    def query_pwd(self):
        words = input('>>> Input address to query: ').lower()
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
                    f'[{match}] Address: {self.df.iloc[match, 0]}  '
                    f'Account: {self.df.iloc[match, 1]}  '
                    f'Password: {self.df.iloc[match, 2]}')
            return matches
        else:
            print(f'[Info] Address not found.')
            return None

    def add_pwd(self):
        address = input('>>> Input address: ')
        account = input('>>> Input account: ')
        password = input('>>> Input password: ')
        self.df.loc[len(self.df)] = [address, account, password]
        self.save()

    def del_pwd(self):
        words = input('>>> Input address to delete: ').lower()
        matches = self.match_words(words)
        if matches:
            index = input('>>> Choose which to delete: ')
            if index == '' and len(matches) == 1:
                ensure = input(f'>>> Delete address {self.df.iloc[matches[0], 0]} ? (y/n): ').lower()
                if ensure == 'y' or ensure == '':
                    self.df.drop(matches[0], axis=0, inplace=True)
                    print('[Info] Deleted successfully.')
                    self.df.reset_index(drop=True, inplace=True)
                    self.save()
                else:
                    print('[Info] Deletion canceled.')
            else:
                self.del_index(index, matches)

    def del_index(self, index:str, matches:list):
        if index.isdigit():
            trush_index = int(index)
            if trush_index in matches:
                ensure = input(f'>>> Delete address {self.df.iloc[trush_index, 0]} ? (y/n): ').lower()
                if ensure == 'y' or ensure == '':
                    self.df.drop(matches[0], axis=0, inplace=True)
                    print('[Info] Deleted successfully.')
                    self.df.reset_index(drop=True, inplace=True)
                    self.save()
                else:
                    print('[Info] Deletion canceled.')
            else:
                trush_index = input('>>> Please input correct [number]: ')
                self.del_index(trush_index, matches)
        else:
            trush_index = input('>>> Please input correct [number]: ')
            self.del_index(trush_index, matches)

    def chg_pwd(self):
        chg_address = input('>>> Input address to change password: ').lower()
        matches = self.match_words(chg_address)
        if matches:
            index = input('>>> Choose which to change password: ')
            if index == '' and len(matches) == 1:
                password = input('>>> Input new password: ')
                ensure = input(f'>>> Change password to {password} ? (y/n): ').lower()
                if ensure == 'y' or ensure == '':
                    self.df.iloc[matches[0], 2] = password
                    print('[Info] Changed successfully.')
                    self.save()
                else:
                    print('[Info] Change canceled.')
            else:
                self.chg_index(index ,matches)

    def chg_index(self, index:str, matches:list):
        if index.isdigit():
            chg_index = int(index)
            if chg_index in matches:
                password = input('>>> Input new password: ')
                ensure = input(f'>>> Change password to {password} ? (y/n): ').lower()
                if ensure == 'y' or ensure == '':
                    self.df.iloc[chg_index, 2] = password
                    print('[Info] Changed successfully.')
                    self.save()
                else:
                    print('[Info] Change canceled.')
            else:
                chg_index = input('>>> Please input correct [number]: ')
                self.chg_index(chg_index, matches)
        else:
            chg_index = input('>>> Please input correct [number]: ')
            self.chg_index(chg_index, matches)

    def save(self):
        with pd.ExcelWriter(self.fileDir) as wt:
            self.df.to_excel(wt, 'Sheet1')
        print('[Info] Saved successfully.')


def first_ues():
    if os.path.exists('config.ini'):
        return False
    else:
        if getattr(sys, 'frozen', False):
            fileDir = os.path.join(os.path.dirname(sys.executable), 'myPassword.xlsx')
        else:
            fileDir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'myPassword.xlsx')
        print("[Info] It's your first time to run this app. A <config.ini> file will be created, where you can change the password file directory.")
        with open('config.ini', 'w', encoding='utf-8') as file:
            file.write(f'[config]\nfile_dir = {fileDir}\n')
        print("[Info] And now you can go to set the file directory.")
        input('>>> (Press Enter to exit...)')
        return True


def main():
    if not first_ues():
        config = ConfigParser()
        config.read('config.ini', encoding='utf-8')
        fileDir = config['config']['file_dir']
        print(f'[Info] fileDir: {fileDir}')
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