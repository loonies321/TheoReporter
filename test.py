'''
   Copyright 2022 Maksim Trushin  PET-Technology Podolsk
   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at
       http://www.apache.org/licenses/LICENSE-2.0
   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
'''

import os
from os import startfile
from psutil import process_iter
from time import sleep
import time
import glob
import numpy as np
import pandas as pd


def is_completed():
    # function to know dose the program complits
    # cheking for report file in current folder, True or False
    report_files = glob.glob('report.xlsx')
    if len(report_files) > 0:
        return True
    else:
        return False


def main():
    reporter_name = 'reporter.exe'
    df_sum = []
    path = os.getcwd()
    print(path + '\\' + reporter_name)
    # 1st we'll make list with folders
    folders = [name for name in os.listdir() if os.path.isdir(name)]
    for ind, i in enumerate(folders):
        # copy reporter file into folders with pdf reports
        shell_copy = f'copy {path}\\{reporter_name} {path}\\{i}'
        os.system(shell_copy)
        # cd to the dir with reporter and reports
        os.chdir(f'{path}\\{i}')
        # lets start reporter.exe
        startfile(reporter_name)
        # wait for his work and check report.xlsx in folder
        while not is_completed():
            sleep(1)
        # kill the reporter
        for proc in process_iter():
            if proc.name() == reporter_name.split('\\')[-1]:
                proc.kill()
        # del reporter from folder
        shell_del = f'del {path}\\{i}\\{reporter_name}'
        os.system(shell_del)
        # add report to pandas
        # tmp = []
        # read report exel file
        tmp = pd.read_excel(f'{path}\\{i}\\report.xlsx')
        # add clear row in report
        tmp = tmp.shift(periods=1)
        if ind == 0:
            df_sum = tmp
        else:
            df_sum = pd.concat([df_sum, tmp])
        shell_del = f'del {path}\\{i}\\report.xlsx'
        os.system(shell_del)
        os.chdir(path)
    # os.chdir('..')
    df_sum.replace(np.NaN, '', inplace=True)
    df_sum.to_excel('report.xlsx', index=False)


if __name__ == '__main__':

    start_time = time.time()
    main()
    end_time = time.time()
    os.system('cls')
    print(f'время выполнения = {end_time - start_time} секунд')
    print('------------------SUCCES!-------------------', '\n\n\nPress Enter to exit')
    input()
