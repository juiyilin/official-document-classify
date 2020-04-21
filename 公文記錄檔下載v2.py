import time
import os
import pandas as pd
import xlrd
import xlwt


# 民國年月
date = time.localtime()
roc_year = str(int(time.strftime('%Y', date))-1911)
month = time.strftime('%m', date)
day = time.strftime('%d', date)
print('%s/%s/%s' % (roc_year, month, day))

# 找到檔案
source = os.listdir('C://Users//d33703//Downloads')
source_file = [f for f in source if time.strftime('%Y', date)+month+day in f]
print(*source_file, end=', ')

# 資料分類
for doc in source_file:
    data = pd.read_excel('C://Users//d33703//Downloads/'+doc)
    filter_data = pd.DataFrame()
    doc_names = list(data['性質'].unique())
    doc_names.sort()
    for doc_name in doc_names:
        try:
            print('處理', doc_name, '中', sep='')
            f = data[data['性質'].str.contains(doc_name)]
        except TypeError:
            error_list = list(data['歸檔編號'][data['性質'].isna()])
            print(*a, sep=', ', end=' ')
            print('沒有打勾')
        except:
            print('在', doc_name, '時遇到不明錯誤', sep='')
        else:
            # print(f)
            filter_data = filter_data.append(f)

        # 移除舊檔
        path = '//box/國際處/部門資料夾/國際處收發公文登錄暨發文電子檔/歸檔公文匯出/'
        try:
            os.chdir(path+roc_year+'年')
        except FileNotFoundError:
            os.mkdir(path+roc_year+'年')
            print('建立', roc_year, '資料夾')
            os.mkdir(path+roc_year+'年'+'/'+doc_name[:-2])
            print('建立', doc_name[:-2], '資料夾')
        except:
            print('不明錯誤')
        else:
            try:
                os.chdir(path+roc_year+'年'+'/'+doc_name[:-2])
            except FileNotFoundError:
                os.mkdir(path+roc_year+'年'+'/'+doc_name[:-2])
                print('建立', doc_name[:-2], '資料夾')
            else:
                for xls in os.listdir():
                    if doc_name in xls:
                        print('刪除', xls)
                        os.remove(xls)

        # 存檔
        name = doc_name + '-' + roc_year + '年1-' + str(int(month)) + '月.xls'
        filter_data.to_excel(path + roc_year + '年/' +
                             doc_name[:-2] + '/' + name, index=False)
        print(doc_name+'已存檔完成')
        filter_data = pd.DataFrame()
