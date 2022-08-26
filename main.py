import pandas as pd
import openpyxl as opxl
import numpy as np
import xlsxwriter
#Загружаем файл, в который хотим получить реузльтат
result = pd.read_excel('Тестовое задание.xlsx','Решение')
#загружаем данные
data1 = pd.read_excel('Данные ЭС обезличен.xlsx')
data2 = pd.read_excel('ЭС_июнь обезличен.xlsx')


def account_number_create(string):
    '''
    Функция преобразования лицивого счета в формат метчинга
    '''
    if type(string) is str:     
        string = string.split('_')[-1]
        string= 'Л/С: ' + string
    return string
#Проебразуеим лицевой счет в необходимый формат 
data1['Лицевой счет'] = data1['Код (О)'].apply(account_number_create)

def addres_edit(string):
    '''
    Удаляет индекс из адреса
    '''
    if type(string) is str:
        string = string.split(',')
        new_string = ''
        for i in range(1,len(string)):
            new_string += string[i].strip() +', '

        new_string = new_string[:-2].strip()
        return new_string
    else:
        return string

#Преобразуем адрес
data1['Адрес юридический (ФИАС) наименование* (К)'] = data1['Адрес юридический (ФИАС) наименование* (К)'].apply(addres_edit)


def complement_list(list1,list2):
    '''
    Дополняет лист1 значениями листа2, если значение листа1 пусто или имеет Unnamed
    Подходит только для одинаковых листов
    '''
    for i in range(len(list1)):
        if (list1[i] is np.nan) or ('Unnamed' in str(list1[i])):
            list1[i]=list2[i]
    
    return list1
#Называем колонки датафремов из строк в них, а затем сокращаем df'ы до значений
data2.columns = complement_list(list(data2.iloc[2]),list(data2.iloc[1]))
data1.columns = complement_list(list(data1.columns),list(data1.iloc[0]))
data2 = data2[3:]
data1 = data1[1:]

#Мержим наши df по лицевому счету
n_data = pd.merge(data1,data2,how='right',on='Лицевой счет')

#Оставляем только нужные значения dtaframe n_data  и помещаем их в ndf dataframe
ndf = n_data[['Лицевой счет','Адрес юридический (ФИАС) наименование* (К)','Код (СД)','Код (К)',
              'Дата статуса','Сумма оплачено','Портфель для группировки']]
#Удаляем колонки датафрема, в которых нужено внести информаци, мечим и получаем нужный датафрейм
result = result.drop(['Адрес','СД','К ','Дата статуса','Сумма оплачено'],axis='columns')
result = pd.merge(result,ndf,how='left',on='Лицевой счет')
#Проставляем номерацию
result['№'] = result.index+1
#Находим сумму колонки 'сумма оплачено' по портфелям группировки и помещаем в df 
sum_group = result['Сумма оплачено'].groupby(result['Портфель для группировки']).sum().reset_index()
#Можем взглянуть на таблицу и построить график
print(sum_group)
sum_group.plot()
#Выводим таблицу в excel file
with pd.ExcelWriter('result.xlsx',engine='xlsxwriter') as writer:  
    result.to_excel(writer, sheet_name ='Решение',index=False) 
    sum_group.to_excel(writer, sheet_name ='Портфель для группировки',index=False) 
