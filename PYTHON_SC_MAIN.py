###############################################################################

### Алгоритм расчета скоринговой модели

### Авторы:
## Богданов В.
## Масленков А.
## Федосеев Д.

###############################################################################

### Загрузка библиотек

import os
import gc
import copy
import pandas as pd
import numpy as np
import datetime as dtime
import warnings
import openpyxl
import sklearn.metrics as metrics
import matplotlib.pyplot as plt
from concurrent import futures as pool
from scipy.stats import norm
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegressionCV, \
                                 LogisticRegression, \
                                 SGDClassifier
from sklearn import svm
from sklearn.calibration import CalibratedClassifierCV
from sklearn.neighbors import KNeighborsClassifier, \
                              RadiusNeighborsClassifier
from sklearn.gaussian_process import GaussianProcessClassifier
from sklearn.gaussian_process.kernels import RBF
from sklearn.naive_bayes import GaussianNB, \
                                MultinomialNB, \
                                BernoulliNB                               
from sklearn.tree import DecisionTreeClassifier
from sklearn.ensemble import RandomForestClassifier, \
                             ExtraTreesClassifier, \
                             AdaBoostClassifier, \
                             BaggingClassifier, \
                             GradientBoostingClassifier, \
                             VotingClassifier
from sklearn.neural_network import MLPClassifier
from sklearn.preprocessing import StandardScaler
from sc_functions import download_csv_file, \
                         download_excel_file, \
                         binning_by_param_loop, \
                         binning_result_unification, \
                         binning_get_shortlist_after_iv, \
                         binning_get_data_for_trend, \
                         test_stat, \
                         trend_flg, \
                         trend, \
                         get_pv_sequences, \
                         get_attr_type, \
                         get_detail_check, \
                         get_Trend_detail, \
                         get_Trend_Unique_detail, \
                         get_Correl_detail, \
                         get_detail_from_excel, \
                         long_list_final_check, \
                         long_list_final_detail, \
                         get_gini_gap, \
                         get_gini_gap_flg, \
                         cramers_corrected_stat, \
                         append_df_to_excel, \
                         bin_df, \
                         run_randomsearch, \
                         roc_auc, \
                         predict, \
                         cat_variable_corr, \
                         corrmtrx_to_excel                        

#from oracle_load import load_data_from_oracle
                         
#from Config import Config
#from DataAccess import DataAccess
#import MyLog
                        
###############################################################################
                         
### Изменяемые переменные, проверить их перед запуском алгоритма!
                         
## Данные по идентификации модели
                         
model_name = 'Модель оттока для HomeCreditBank'

## Список ключевых переменных, если нет ставить ''

person_id_attr = 'IDCLIENT'  # ID Клиента/телефон
time_id_attr   = 'DATE_DATA' # Дата события
target_attr    = 'EVENT'     # Бинарная целевая функция 0/1  

## Уровень cutoff по проверкам переменных

IV_left_cutoff    = 0.05 # нижний уровень отсечения IV по переменной
IV_right_cutoff   = 1.00 # верхний уровень отсечения IV по переменной
PSI_bin_cutoff    = False # флаг проверки PSI по всей переменной (False) или по каждому бину (True)
PSI_cutoff        = 0.25 # верхний уровень отсечения PSI
VOL_bin_cutoff    = False # флаг проверки волатильности по всей переменной (False) или по каждому бину (True)
VOL_cutoff        = 2.00 # верхний уровень отсечения волатильности
break_trd_cutoff  = 1 # максимально допустимое кол-во нарушенных бакетов трендовости для бинирования
bin_v_cutoff      = 0.005 # минимальный уровень объема одного бина от общего объема выборки
PVALUE_cutoff     = 0.05 # верхний уровень отсечения PValue по переменной
PVALUE_thresholds = [0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3, 0.2, 0.1, 0.05] # поэтапные уровни отсечения для PValue
CORREL_cutoff     = 0.30 # верхний уровень отсечения по парной корреляции для переменной

## Настройки для вывода валидационных статистик

ROCAUC_n_bucket  = 20 # кол-во бакетов в графике построения ROCAUC
PREDICT_n_bucket = 15 # кол-во бакетов в графике построения PREDICT
TREND_long = False    # флаг вывода графиков трендовости по всем переменным из шорт листа с иными вариантами бинирования
undersampling_rate = 0 # Кол-во раз, во сколько уменьшали нецелевые события в выборке (если не уменьшали, то 0)

## Таблицы Хранилища для загрузки данных

use_oracle = False                  # флаг использования базы oracle для загрузки данных
use_oracle_1 = False                # флаг использования таблицы 1: True/False
ora_select_1 = ''  # запрос к таблице 1
use_oracle_2 = False                # флаг использования таблицы 2: True/False
ora_select_2 = ''  # запрос к таблице 2

## Файлы CSV для загрузки данных

use_file = True                        # флаг использования файлов csv для загрузки данных
use_file_1 = True                      # флаг использования файла 1: True/False
filepath_or_buffer_1 = 'Данные\\Churn_Modelling_2.csv' # путь к файлу 1
use_file_2 = False                      # флаг использования файла 2: True/False
filepath_or_buffer_2 = '' # путь к файлу 2
use_file_3 = False                      # флаг использования файла 3 (доп категориальные переменные) :True/False
filepath_or_buffer_3 = '' # путь к файлу 3

## Данные для рестарта процесса с определенного этапа

use_restart_file = False # флаг использования файла с данными для рестарта: True/False
filepath_restart = r'C:\Users\Виталий\Desktop\Python_actual\RSB\Model # 2021-08-16 22_04_09\Model # 2021-08-16 22_04_09 # 3_backup.xlsx' # путь к файлу с данными для рестарта

use_restart_binning = False        # флаг использования посчитанного бинирования: True/False
sheetname_binning = 'bin_restart' # название листа с результатами бинирования

use_restart_trend = False         # флаг использования посчитанного тренда: True/False
sheetname_trend = 'trend_restart' # название листа с результатами тренда

use_restart_pvalue = False          # флаг использования посчитанного pvalue: True/False
sheetname_pvalue = 'pvalue_restart' # название листа с результатами pvalue

use_restart_correl = False          # флаг использования посчитанной корреляции: True/False
sheetname_correl = 'correl_restart' # название листа с результатами корреляции

## Предобработка данных

change_types_data = True  # флаг преобразования типов данных на основе их наименований: True/False
add_missing_data  = True  # флаг заполнения пустых значений: True/False
round_data        = True  # флаг округления числовых значений: True/False
strip_data        = False  # флаг удалений лишних пробелов в строковых значениях: True/False
int_to_cat_data   = False  # флаг создания категориальных переменных на основе интервальных: True/False
normalize_data    = False # флаг нормализации числовых данных: True/False

## Разбиение выборки на тренировочную и тестовую

use_file_train_test = False   # флаг ручного разбиения переменных на трейн/тест
filepath_or_buffer_train = '' # флаг использования файла трейн
filepath_or_buffer_test = ''  # флаг использования файла тест
test_size_split    = 0.30     # размер тестовой выборки от общей (доля)
random_state_split = 1        # фиксированное состояние разбиения

## Заполнение переменных для расчета модели

stop_list_attr = ['Surname'                 
                  ] # стоп-лист переменных: не используются в модели

my_long_list = [] # ручное заполнение лонг-листа переменных: только они будут использоваться в модели (attr_true_name)

my_short_list = [] # ручное заполнение забинированных переменных: только они будут использоваться в модели (attr_name)

## параметры бинирования

min_perc_total_list = [0.01, 0.02, 0.04, 0.06, 0.08] # [0.03]
min_perc_class_list = [0.001, 0.002, 0.004, 0.008, 0.01] # [0.001]
stop_limit_list     = [0.02, 0.04, 0.06, 0.08, 0.10] # [0.01]

## виды классификаторов для обучения

model_type_0  = True # 'Логистическая регрессия "LogisticRegression"'
model_type_1  = False# 'Логистическая регрессия "LogisticRegressionCV"'
model_type_2  = False#True # 'Метод опорных векторов "SVC"'
model_type_3  = False # НЕ РАБОТАЕТ # 'Метод опорных векторов "NuSVC"'
model_type_4  = False#True # 'Метод опорных векторов "LinearSVC"'
model_type_5  = False#True # 'Стохастический градиентный спуск "SGDClassifier"'
model_type_6  = False#True # 'Классификатор ближайших соседей "KNeighborsClassifier"'
model_type_7  = False # НЕ РАБОТАЕТ # 'Классификатор ближайших соседей "RadiusNeighborsClassifier"'
model_type_8  = False # НЕ РАБОТАЕТ # 'Гаусовский процесс "GaussianProcessClassifier"'
model_type_9  = False#True # 'Наивный байесовский классификатор "GaussianNB"'
model_type_10 = False#True # 'Наивный байесовский классификатор "MultinomialNB"'
model_type_11 = False # НЕ РАБОТАЕТ # 'Наивный байесовский классификатор "ComplementNB"'
model_type_12 = False # НЕ РАБОТАЕТ # 'Наивный байесовский классификатор "CategoricalNB"'
model_type_13 = False # 'Наивный байесовский классификатор "BernoulliNB"'
model_type_14 = False#True # 'Деревья решений "DecisionTreeClassifier"'
model_type_15 = False#True # 'Ансамбль "BaggingClassifier"'
model_type_16 = False#True # 'Ансамбль "RandomForestClassifier"'
model_type_17 = False #True # 'Ансамбль "ExtraTreesClassifier"'
model_type_18 = False #True # 'Ансамбль "AdaBoostClassifier"'
model_type_19 = False#True # 'Ансамбль "GradientBoostingClassifier"'
model_type_20 = False#True # 'Ансамбль "VotingClassifier"'
model_type_21 = False # НЕ РАБОТАЕТ # 'Ансамбль "StackingClassifier"'
model_type_22 = False#True # 'Нейронные сети "MLPClassifier"'

###############################################################################

### Неизменяемые переменные, задаются логикой алгоритма, не трогать!

## Общие переменные

t = '---'                         # переменная для вывода в консоли
warnings.filterwarnings('ignore') # игнорируем warning-и, для показа поставить default
bin_v_cutoff = bin_v_cutoff * 100 # в алгоритме данная переменная нужна в процентах, а не в долях

## Настройка наименований директории и файлов EXCEL с отчетом по модели

datetime_str = str(dtime.datetime.now()).replace(':','_') # текущая дата строкой
point_index = datetime_str.find('.')                      # позиция точки
datefilename = 'Model # ' + datetime_str[0:point_index]   # наименование для папки/файлов
os.mkdir(datefilename)                                    # создание папки, где будут храниться все файлы по модели

excelFileName  = datefilename + '/' + datefilename + ' # 1_report.xlsx'     # основной файл с отчетом по модели
excelFileNameS = datefilename + '/' + datefilename + ' # 2_statistics.xlsx' # статистика по переменным из лонг-листа
excelFileNameB = datefilename + '/' + datefilename + ' # 3_backup.xlsx'     # вспомогательный файл с данными для перезапуска

filepath_trend_stat_restart = filepath_restart[0:filepath_restart.find('3_backup.xls')] + '2_statistics.xlsx' # путь к файлу с данными по статистике

## Создаем отчет по будущей модели

modelInfo = pd.DataFrame(columns = ['parameter_name', 'parameter_value', 'commentary'])

modelInfo.loc[0] = ['Данные по идентификации модели',None,None]
modelInfo.loc[1] = ['model_name',model_name,'Наименование модели']
modelInfo.loc[2] = ['Список ключевых переменных',None,None]
modelInfo.loc[3] = ['person_id_attr',person_id_attr,'ID Клиента/телефон']
modelInfo.loc[4] = ['time_id_attr',time_id_attr,'Дата события']
modelInfo.loc[5] = ['target_attr',target_attr,'Бинарная целевая функция 0/1']
modelInfo.loc[6] = ['Уровень cutoff по проверкам переменных',None,None]
modelInfo.loc[7] = ['IV_left_cutoff',IV_left_cutoff,'нижний уровень отсечения IV по переменной']
modelInfo.loc[8] = ['IV_right_cutoff',IV_right_cutoff,'верхний уровень отсечения IV по переменной']
modelInfo.loc[9] = ['PSI_bin_cutoff',PSI_bin_cutoff,'флаг проверки PSI по всей переменной (False) или по каждому бину (True)']
modelInfo.loc[10] = ['PSI_cutoff',PSI_cutoff,'верхний уровень отсечения PSI']
modelInfo.loc[11] = ['VOL_bin_cutoff',VOL_bin_cutoff,'флаг проверки волатильности по всей переменной (False) или по каждому бину (True)']
modelInfo.loc[12] = ['VOL_cutoff',VOL_cutoff,'верхний уровень отсечения волатильности']
modelInfo.loc[13] = ['break_trd_cutoff',break_trd_cutoff,'максимально допустимое кол-во нарушенных бакетов трендовости для бинирования']
modelInfo.loc[14] = ['bin_v_cutoff',bin_v_cutoff,'минимальный уровень объема одного бина от общего объема выборки']
modelInfo.loc[15] = ['PVALUE_cutoff',PVALUE_cutoff,'верхний уровень отсечения PValue по переменной']
modelInfo.loc[16] = ['PVALUE_thresholds',PVALUE_thresholds,'поэтапные уровни отсечения для PValue']
modelInfo.loc[17] = ['CORREL_cutoff',CORREL_cutoff,'верхний уровень отсечения по парной корреляции для переменной']
modelInfo.loc[18] = ['Настройки для вывода валидационных статистик',None,None]
modelInfo.loc[19] = ['ROCAUC_n_bucket',ROCAUC_n_bucket,'кол-во бакетов в графике построения ROCAUC']
modelInfo.loc[20] = ['PREDICT_n_bucket',PREDICT_n_bucket,'кол-во бакетов в графике построения PREDICT']
modelInfo.loc[21] = ['TREND_long',TREND_long,'флаг вывода графиков трендовости по всем переменным из шорт листа с иными вариантами бинирования']
modelInfo.loc[22] = ['undersampling_rate',undersampling_rate,'Кол-во раз, во сколько уменьшали нецелевые события в выборке (если не уменьшали, то 0)']
modelInfo.loc[23] = ['Таблицы Хранилища для загрузки данных',None,None]
modelInfo.loc[24] = ['use_oracle',use_oracle,'флаг использования базы oracle для загрузки данных']
modelInfo.loc[25] = ['use_oracle_1',use_oracle_1,'флаг использования таблицы 1: True/False']
modelInfo.loc[26] = ['ora_select_1',ora_select_1,'запрос к таблице 1']
modelInfo.loc[27] = ['use_oracle_2',use_oracle_2,'флаг использования таблицы 2: True/False']
modelInfo.loc[28] = ['ora_select_2',ora_select_2,'запрос к таблице 2']
modelInfo.loc[29] = ['Файлы CSV для загрузки данных',None,None]
modelInfo.loc[30] = ['use_file',use_file,'флаг использования файлов csv для загрузки данных']
modelInfo.loc[31] = ['use_file_1',use_file_1,'флаг использования файла 1: True/False']
modelInfo.loc[32] = ['filepath_or_buffer_1',filepath_or_buffer_1,'путь к файлу 1']
modelInfo.loc[33] = ['use_file_2',use_file_2,'флаг использования файла 2: True/False']
modelInfo.loc[34] = ['filepath_or_buffer_2',filepath_or_buffer_2,'путь к файлу 2']
modelInfo.loc[35] = ['use_file_3',use_file_3,'флаг использования файла 3 (доп категориальные переменные) :True/False']
modelInfo.loc[36] = ['filepath_or_buffer_3',filepath_or_buffer_3,'путь к файлу 3']
modelInfo.loc[37] = ['Данные для рестарта процесса с определенного этапа',None,None]
modelInfo.loc[38] = ['use_restart_file',use_restart_file,'# флаг использования файла с данными для рестарта: True/False']
modelInfo.loc[39] = ['filepath_restart',filepath_restart,'путь к файлу с данными для рестарта']
modelInfo.loc[40] = ['use_restart_binning',use_restart_binning,'флаг использования посчитанного бинирования: True/False']
modelInfo.loc[41] = ['sheetname_binning',sheetname_binning,'название листа с результатами бинирования']
modelInfo.loc[42] = ['use_restart_trend',use_restart_trend,'флаг использования посчитанного тренда: True/False']
modelInfo.loc[43] = ['sheetname_trend',sheetname_trend,'название листа с результатами тренда']
modelInfo.loc[44] = ['use_restart_pvalue',use_restart_pvalue,'флаг использования посчитанного pvalue: True/False']
modelInfo.loc[45] = ['sheetname_pvalue',sheetname_pvalue,'название листа с результатами pvalue']
modelInfo.loc[46] = ['use_restart_correl',use_restart_correl,'флаг использования посчитанной корреляции: True/False']
modelInfo.loc[47] = ['sheetname_correl',sheetname_correl,'название листа с результатами корреляции']
modelInfo.loc[48] = ['Предобработка данных',None,None]
modelInfo.loc[49] = ['change_types_data',change_types_data,'флаг преобразования типов данных на основе их наименований: True/False']
modelInfo.loc[50] = ['add_missing_data',add_missing_data,'флаг заполнения пустых значений: True/False']
modelInfo.loc[51] = ['round_data',round_data,'флаг округления числовых значений: True/False']
modelInfo.loc[52] = ['strip_data',strip_data,'флаг удалений лишних пробелов в строковых значениях: True/False']
modelInfo.loc[53] = ['int_to_cat_data',int_to_cat_data,'флаг создания категориальных переменных на основе интервальных: True/False']
modelInfo.loc[54] = ['normalize_data',normalize_data,'флаг нормализации числовых данных: True/False']
modelInfo.loc[55] = ['Разбиение выборки на тренировочную и тестовую',None,None]
modelInfo.loc[56] = ['use_file_train_test',use_file_train_test,'флаг ручного разбиения переменных на трейн/тест']
modelInfo.loc[57] = ['filepath_or_buffer_train',filepath_or_buffer_train,'флаг использования файла трейн']
modelInfo.loc[58] = ['filepath_or_buffer_test',filepath_or_buffer_test,'флаг использования файла тест']
modelInfo.loc[59] = ['test_size_split',test_size_split,'размер тестовой выборки от общей (доля)']
modelInfo.loc[60] = ['random_state_split',random_state_split,'фиксированное состояние разбиения']
modelInfo.loc[61] = ['Заполнение переменных для расчета модели',None,None]
modelInfo.loc[62] = ['stop_list_attr',stop_list_attr,'стоп-лист переменных: не используются в модели']
modelInfo.loc[63] = ['my_long_list',my_long_list,'ручное заполнение лонг-листа переменных: только они будут использоваться в модели (attr_true_name)']
modelInfo.loc[64] = ['my_short_list',my_short_list,'ручное заполнение забинированных переменных: только они будут использоваться в модели (attr_name)']
modelInfo.loc[65] = ['параметры бинирования',None,None]
modelInfo.loc[66] = ['min_perc_total_list',min_perc_total_list,'Параметр: Кол-во изначальных бинов для количественных переменных']
modelInfo.loc[67] = ['min_perc_class_list',min_perc_class_list,'Параметр: Обязательно склеивает бин, если его WOE меньше данного значения']
modelInfo.loc[68] = ['stop_limit_list',stop_limit_list,'Параметр: Не склеивает бин дальше, если его WOE больше данного значения']
modelInfo.loc[69] = ['виды классификаторов для обучения',None,None]
modelInfo.loc[70] = ['model_type_0',model_type_0,'Логистическая регрессия "LogisticRegression"']
modelInfo.loc[71] = ['model_type_1',model_type_1,'Логистическая регрессия "LogisticRegressionCV"']
modelInfo.loc[72] = ['model_type_2',model_type_2,'Метод опорных векторов "SVC"']
modelInfo.loc[73] = ['model_type_3',model_type_3,'Метод опорных векторов "NuSVC"']
modelInfo.loc[74] = ['model_type_4',model_type_4,'Метод опорных векторов "LinearSVC"']
modelInfo.loc[75] = ['model_type_5',model_type_5,'Стохастический градиентный спуск "SGDClassifier"']
modelInfo.loc[76] = ['model_type_6',model_type_6,'Классификатор ближайших соседей "KNeighborsClassifier"']
modelInfo.loc[77] = ['model_type_7',model_type_7,'Классификатор ближайших соседей "RadiusNeighborsClassifier"']
modelInfo.loc[78] = ['model_type_8',model_type_8,'Гаусовский процесс "GaussianProcessClassifier"']
modelInfo.loc[79] = ['model_type_9',model_type_9,'Наивный байесовский классификатор "GaussianNB"']
modelInfo.loc[80] = ['model_type_10',model_type_10,'Наивный байесовский классификатор "MultinomialNB"']
modelInfo.loc[81] = ['model_type_11',model_type_11,'Наивный байесовский классификатор "ComplementNB"']
modelInfo.loc[82] = ['model_type_12',model_type_12,'Наивный байесовский классификатор "CategoricalNB"']
modelInfo.loc[83] = ['model_type_13',model_type_13,'Наивный байесовский классификатор "BernoulliNB"']
modelInfo.loc[84] = ['model_type_14',model_type_14,'Деревья решений "DecisionTreeClassifier"']
modelInfo.loc[85] = ['model_type_15',model_type_15,'Ансамбль "BaggingClassifier"']
modelInfo.loc[86] = ['model_type_16',model_type_16,'Ансамбль "RandomForestClassifier"']
modelInfo.loc[87] = ['model_type_17',model_type_17,'Ансамбль "ExtraTreesClassifier"']
modelInfo.loc[88] = ['model_type_18',model_type_18,'Ансамбль "AdaBoostClassifier"']
modelInfo.loc[89] = ['model_type_19',model_type_19,'Ансамбль "GradientBoostingClassifier"']
modelInfo.loc[90] = ['model_type_20',model_type_20,'Ансамбль "VotingClassifier"']
modelInfo.loc[91] = ['model_type_21',model_type_21,'Ансамбль "StackingClassifier"']
modelInfo.loc[92] = ['model_type_22',model_type_22,'Нейронные сети "MLPClassifier"']

modelInfo = modelInfo[['parameter_name','commentary','parameter_value']]

append_df_to_excel(excelFileName, modelInfo, 'modelInfo', 0, True)

## Заполняем лист ключевых атрибутов

id_attr = [] # лист ключевых атрибутов без целевой функции
if person_id_attr != '':
    id_attr.append(person_id_attr)
if time_id_attr != '':
    id_attr.append(time_id_attr)

key_attr = [] # лист ключевых атрибутов с целевой функцией
if person_id_attr != '':
    key_attr.append(person_id_attr)
if time_id_attr != '':
    key_attr.append(time_id_attr)
if target_attr != '':
    key_attr.append(target_attr)
    
## Словарь переменных для стратификации выборок для train/test

stratify_dict = [target_attr]
    
## Проверка для рестартовых данных
    
if use_restart_file == False:
    if use_restart_binning == True:
        use_restart_binning = False
        print('Бинирование не может быть перезапущено! Выключен use_restart_file!')
    if use_restart_trend == True:
        use_restart_trend = False
        print('Трендовость не может быть перезапущена! Выключен use_restart_file!')
    if use_restart_pvalue == True:
        use_restart_pvalue = False
        print('PValue не может быть перезапущено! Выключен use_restart_file!')
    if use_restart_correl == True:
        use_restart_correl = False
        print('Корреляция не может быть перезапущена! Выключен use_restart_file!')
elif use_restart_binning == False:
    if use_restart_trend == True:
        use_restart_trend = False
        print('Трендовость не может быть перезапущена! Выключен use_restart_binning!')
    if use_restart_pvalue == True:
        use_restart_pvalue = False
        print('PValue не может быть перезапущено! Выключен use_restart_binning!')
    if use_restart_correl == True:
        use_restart_correl = False
        print('Корреляция не может быть перезапущена! Выключен use_restart_binning!')
elif use_restart_trend == False:
    if use_restart_pvalue == True:
        use_restart_pvalue = False
        print('PValue не может быть перезапущено! Выключен use_restart_trend!')
    if use_restart_correl == True:
        use_restart_correl = False
        print('Корреляция не может быть перезапущена! Выключен use_restart_trend!')
elif use_restart_pvalue == False:
    if use_restart_correl == True:
        use_restart_correl = False
        print('Корреляция не может быть перезапущена! Выключен use_restart_pvalue!')
    
###############################################################################       

### Часть 1. Загрузка данных в питон

print(t)
print('1.1. Алгоритм запущен: ' + str(dtime.datetime.now()))

if use_oracle == True: ## Загрузка из хранилища Oracle

    try:       
        if use_oracle_1 == True:
            df1 = load_data_from_oracle(tablename=ora_select_1)
            print('1.2. Файл №1 загружен: ' + str(dtime.datetime.now()))        
        if use_oracle_2 == True:
            df2 = load_data_from_oracle(tablename=ora_select_2)
            print('1.3. Файл №2 загружен: ' + str(dtime.datetime.now()))       
    except Exception as e: 
        print('Error of DB connection (%s)' % e)
        raise SystemExit('Error of DB connection')
      
elif use_file == True: ## Загрузка из файлов CSV
    
    if use_file_1 == True: ## загрузка 1 файла с переменными
        df1 = download_csv_file(filepath_or_buffer_1)
        print('1.2.1. Файл №1 загружен: ' + str(dtime.datetime.now()))
           
    if use_file_2 == True: ## загрузка 2 файла с переменными
        df2 = download_csv_file(filepath_or_buffer_2)    
        print('1.2.2. Файл №2 загружен: ' + str(dtime.datetime.now()))
        
    if use_file_3 == True: ## загрузка 3 файла с переменными
        df3 = download_csv_file(filepath_or_buffer_3, myindex_col = 'INDEX')    
        print('1.2.3. Файл №3 загружен: ' + str(dtime.datetime.now()))
        
    if use_file_train_test == True: ## загрузка файлов с train и test выборками
        dftr = download_csv_file(filepath_or_buffer_train)    
        print('1.2.4. Файл train загружен: ' + str(dtime.datetime.now()))
        dfts = download_csv_file(filepath_or_buffer_test)    
        print('1.2.5. Файл train загружен: ' + str(dtime.datetime.now()))
       
else:
    raise Exception('Не выбран ни один способ загрузки исходных данных')
        
## Составление итогового файла
    
if (use_file_1 == True and use_file_2 == True) or (use_oracle_1 == True and use_oracle_2 == True):
    df1 = df1.sort_values(by = key_attr).reset_index(drop = True)
    df2 = df2.sort_values(by = key_attr).reset_index(drop = True)
    df2 = df2.drop(key_attr, axis = 1)
    df = df1.join(df2, how = 'inner')  
elif (use_file_1 == True and use_file_2 == False): #or (use_oracle_1 == True and use_oracle_2 == False):
    df = df1
elif (use_file_1 == False and use_file_2 == True): #or (use_oracle_1 == False and use_oracle_2 == False):
    df = df2   
elif (use_file_1 == False and use_file_2 == False and use_file_3 == True):    
    int_to_cat_attr_list = list(df3)
    for key_attr_name in key_attr:
        if key_attr_name in int_to_cat_attr_list:
            int_to_cat_attr_list.remove(key_attr_name)
    df = df3
    df[int_to_cat_attr_list] = df[int_to_cat_attr_list].astype(str)
elif use_file_train_test == True:
    df = dftr
     
if df.shape[0] == 0 or df.shape[1] == 0:
    raise Exception('Датафрейм df пустой, ошибка при загрузке данных в питон')

print('1.3. Составлен итоговый файл: ' + str(dtime.datetime.now())) 
print(t)

###############################################################################

### Часть 2. Предобработка значений таблицы

print('2.1. Начало предобработки данных: ' + str(dtime.datetime.now())) 

## Список атрибутов для преобразования

all_columns      = list(df)
digit_attr_list  = list(df.select_dtypes(include = ['number']))
string_attr_list = list(df.select_dtypes(include = ['object', 'category']))

for key_attr_name in key_attr:
    if key_attr_name in all_columns:
        all_columns.remove(key_attr_name)
    if key_attr_name in digit_attr_list:
        digit_attr_list.remove(key_attr_name)
    if key_attr_name in string_attr_list:
        string_attr_list.remove(key_attr_name)
   
## Преобразование типов (по названию атрибутов)

if change_types_data == True and (use_file_1 == True or use_file_2 == True or use_file_train_test == True):
    for column_name in all_columns:
        if (column_name.lower().find('str_') != -1 or \
            column_name.lower().find('_str') != -1) \
           and column_name.lower().find('restr_') == -1:  # преобразуем строковые переменные к строке
            df[column_name] = df[column_name].astype(str)
            if use_file_train_test == True:
                dfts[column_name] = dfts[column_name].astype(str)
        elif column_name.lower().find('cnt_') != -1:      # преобразуем числовые переменные к числу
            df[column_name] = df[column_name].astype(float)
            if use_file_train_test == True:
                dfts[column_name] = dfts[column_name].astype(float)
        elif column_name.lower().find('_min_') != -1:      # преобразуем числовые переменные к числу
            df[column_name] = df[column_name].astype(float)
            if use_file_train_test == True:
                dfts[column_name] = dfts[column_name].astype(float)
        elif column_name.lower().find('_avg_') != -1:      # преобразуем числовые переменные к числу
            df[column_name] = df[column_name].astype(float)
            if use_file_train_test == True:
                dfts[column_name] = dfts[column_name].astype(float)
        elif column_name.lower().find('_max_') != -1:      # преобразуем числовые переменные к числу
            df[column_name] = df[column_name].astype(float)
            if use_file_train_test == True:
                dfts[column_name] = dfts[column_name].astype(float)
        elif column_name in ('EstimatedSalary'):      # преобразуем числовые переменные к числу
            df[column_name] = df[column_name].astype(float)
            if use_file_train_test == True:
                dfts[column_name] = dfts[column_name].astype(float)        
    
    gc.collect()    
    print('2.2.1. Преобразованы типы данных: ' + str(dtime.datetime.now())) 
    
## Список атрибутов после преобразования

all_columns      = list(df)
digit_attr_list  = list(df.select_dtypes(include = ['number']))
string_attr_list = list(df.select_dtypes(include = ['object', 'category']))
        
## Заполнение пустых значений
    
if add_missing_data == True:  
    for column_name in digit_attr_list:   
        df[column_name] = df[column_name].fillna(-1)   
        if use_file_train_test == True:
            dfts[column_name] = dfts[column_name].fillna(-1)   
    for column_name in string_attr_list: 
        df[column_name] = df[column_name].fillna('EMPTY')
        if use_file_train_test == True:
            dfts[column_name] = dfts[column_name].fillna('EMPTY') 
    
    gc.collect()    
    print('2.2.2. Заполнены пустые значения: ' + str(dtime.datetime.now())) 

## Округление значений
    
if round_data == True:
    for column_name in digit_attr_list:
        df[column_name] = df[column_name].round(2)
        if use_file_train_test == True:
            dfts[column_name] = dfts[column_name].round(2)
    
    gc.collect()        
    print('2.2.3. Числовые значения округлены: ' + str(dtime.datetime.now())) 

## Удаление лишних пробелов в строках

if strip_data == True:
    for column_name in string_attr_list:
        df[column_name] = df[column_name].str.strip()
        if use_file_train_test == True:
            dfts[column_name] = dfts[column_name].str.strip()
    
    gc.collect()        
    print('2.2.4. В строках убраны лишние пробелы: ' + str(dtime.datetime.now())) 

## Нормализация данных

if normalize_data == True:
    digit_attr_list = list(df.select_dtypes(include = ['number']))
    for i in key_attr:
        if i in digit_attr_list:
            digit_attr_list.remove(i)
    for digit_attr_name in digit_attr_list:
        if np.std(df[digit_attr_name]) != 0:
            df_mean = np.mean(df[digit_attr_name])
            df_std = np.std(df[digit_attr_name])
            df[digit_attr_name] = (df[digit_attr_name] - df_mean)/df_std
            if use_file_train_test == True:
                dfts[digit_attr_name] = (dfts[digit_attr_name] - df_mean)/df_std                
        else:
            df[digit_attr_name] = 0
            if use_file_train_test == True:
                dfts[digit_attr_name] = 0

    gc.collect()    
    print('2.2.5. Данные нормализованы: ' + str(dtime.datetime.now())) 
 
## Преобразование интервальных переменных в категориальные
# Не работает с use_file_train_test == True
    
if int_to_cat_data == True:
    if use_file_3 == False: # рассчитывает датафрейм с доп переменными и выходим из расчета
        df_int_to_cat = pd.DataFrame()
        for key_attr_name in key_attr:
            df_int_to_cat[key_attr_name] = df[key_attr_name]    
        for column_name in digit_attr_list:
            print(column_name, str(dtime.datetime.now()))        
            if df[column_name].nunique() <= 100:
                df_int_to_cat = df_int_to_cat.assign(new_col_name_1 = df[column_name].astype('str'))
                df_int_to_cat.rename(columns = {'new_col_name_1': column_name + '_STR'}, inplace = True)
            elif df[column_name].nunique() > 100:
                cnt_bins = int(df[column_name].nunique() / 10)
                df_int_to_cat = df_int_to_cat.assign(new_col_name_2 = pd.cut (df[column_name], bins=cnt_bins, precision=0).astype('str'), \
                                                     new_col_name_3 = pd.qcut(df[column_name], q=cnt_bins,    precision=0, duplicates = 'drop').astype('str'))
                df_int_to_cat.rename(columns = {'new_col_name_2': column_name + '_STR_CUT', \
                                     'new_col_name_3': column_name + '_STR_QCUT'}, inplace = True)
        
        df_int_to_cat.index.names = ['INDEX']                   
        df_int_to_cat.to_csv(datefilename + ' # int_to_cat.csv', sep=';', encoding ='cp1251')            
        print('2.2.6. Рассчитаны категориальные переменные из интервальных: ' + str(dtime.datetime.now()))
        raise Exception('Выгружены доп переменные!')
        
    elif use_file_3 == True and (use_file_1 == True or use_file_2 == True): # подгружаем уже рассчитанные признаки к основному датафрейму
        df3 = df3.sort_values(by = key_attr).reset_index(drop = True)
        df3 = df3.drop(key_attr, axis = 1)
        df3 = df3.astype(str)
        df = df.join(df3, how = 'inner')
        print('2.2.6. Загружены категориальные переменные из интервальных: ' + str(dtime.datetime.now()))
    elif use_file_3 == True:
        print('2.2.6. Загружены категориальные переменные из интервальных: ' + str(dtime.datetime.now()))
    
## TODO!
# клонирование переменных

print('2.3. Таблица предобработана: ' + str(dtime.datetime.now())) 
print(t)

############################################################################### 

### Часть 3. Разбинение выборки на train и test, подготовка лонг листа переменных

## Заполняем дополнительные переменные для стратификации выборок для train/test

if person_id_attr == 'IDCLIENT':
    if 'Age' in all_columns:
        df['Age_stratify'] = pd.qcut(df['Age'], q=5, precision=0, duplicates = 'drop').astype('str')
        stratify_dict.append('Age_stratify')
    if 'NumOfProducts' in all_columns:
        df['NumOfProducts_stratify'] = pd.qcut(df['NumOfProducts'], q=5, precision=0, duplicates = 'drop').astype('str')
        stratify_dict.append('NumOfProducts_stratify')
    if 'WORK_SALARY' in all_columns:
        df['WORK_SALARY_stratify'] = pd.qcut(df['WORK_SALARY'], q=5, precision=0, duplicates = 'drop').astype('str')
        stratify_dict.append('WORK_SALARY_stratify')
elif person_id_attr == 'PHONE':    
    #TODO: выбрать переменные по телефонам для стратификации
    pass

## Таблица с типом данных

df_types = pd.DataFrame(df.dtypes)
df_types.columns = ['type']
df_types['type'] = df_types['type'].astype('str')
df_types['type'] = df_types.apply(lambda x: get_attr_type(x.type), axis=1)

## Разбиение выборки на тренировочную и тестовую

if use_file_train_test == True:
    df_train = df
    df_test = dfts
else:
    df_train, df_test = train_test_split(df,                                # исходная таблица данных
                                         test_size = test_size_split,       # размер тестовой выборки
                                         shuffle = True,                    # перемешивание данных
                                         random_state = random_state_split, # фиксирование перемешивания
                                         stratify = df[stratify_dict]       # переменные для стратификации (разделения на однородные группы)
                                         )

print("Train - Проникновение ЦФ в %: " + str(round(np.bincount(df_train[target_attr])[1]/ \
            (np.bincount(df_train[target_attr])[0] + np.bincount(df_train[target_attr])[1])*100,4)))
print("Test  - Проникновение ЦФ в %: " + str(round(np.bincount(df_test[target_attr])[1]/ \
            (np.bincount(df_test[target_attr])[0] + np.bincount(df_test[target_attr])[1])*100,4)))

df_train_initial = copy.deepcopy(df_train)
df_test_initial = copy.deepcopy(df_test)

## Расчет надежности построенной модели по тренировочной выборке

obj_qty = df_train.shape[0]
event_1_qty = df_train.query(target_attr + ' == 1').shape[0]
event_rate = event_1_qty / obj_qty

stability_df = pd.DataFrame(columns = ['ERROR','STABILITY'])
for i in range (1, 11, 1):
    error_rate = i/100
    z_val = np.sqrt(obj_qty*np.power(event_rate*error_rate,2)/(event_rate*(1-event_rate)))
    stab = np.around(2*norm.cdf(z_val) - 1,4)
    stability_df.loc[i-1] = [error_rate,stab]

## Статистика по тренировочной выборке по интервальным переменным

df_train_descr = df_train.describe().transpose()
df_train_descr.index.names = ['attr_true_name']
df_train_descr.reset_index()
                         
## Отбрасываем ключевые поля, так как это не атрибуты модели

model_attr = df_train.drop(key_attr, axis = 1)

## Отбрасываем переменные из стоп-листа переменных

for column_name in all_columns:
    if column_name in stop_list_attr:
        model_attr = model_attr.drop(column_name, axis = 1)
        
## Получаем лонг-лист переменных

long_list = list(model_attr)

## Если лонг лист определен вручную, то он становится лонг-листом

if len(my_long_list) > 0:
    long_list = my_long_list

gc.collect()
print('3. Подготовлены выборки train/test: ' + str(dtime.datetime.now())) 
print(t)

############################################################################### 

### Часть 4. Бинирование

### Часть 4.1. Осуществление бинирования

if use_restart_binning == False: # бинируем в первый раз самостоятельно

    binninglist = [] # Название листа, куда складируются результаты бинирования
    i = 0            # Шаг работы алгоритма
    
    ## алгоритм распараллеливания параметров по многопоточности
    
    if __name__ == "__main__":  
        
        with pool.ThreadPoolExecutor() as executor:   
                   
            for attr_in_long_set in long_list: # Цикл по переменным лонг-листа
                i = i + 1                      # Номер переменной для просчета
                future = executor.submit(binning_by_param_loop, 
                                         attr_in_long_set, 
                                         df_train, 
                                         min_perc_total_list, 
                                         stop_limit_list, 
                                         min_perc_class_list,
                                         i)
                binninglist.append(future.result())
    
    ## Распарсиваем результат бинирования в датафрейм
    
    binningDataframe = pd.DataFrame()
    for each_list in binninglist:
        if len(binningDataframe) == 0:
            binningDataframe = each_list
        else:
            binningDataframe = pd.concat([binningDataframe,each_list], axis = 0)   
            
    ## Унификация выходной таблицы из функции бинирования
    
    binningInfoLong = binning_result_unification(binningDataframe)
    
else: # загружаем результаты бинирования из файла    
    binningInfoLong = download_excel_file(filepath_restart, sheetname_binning, \
                                          converters = {'Group_1':str}) 
    print('4.1. Бинирование загружено из restart файла: ' + str(dtime.datetime.now()))
        
append_df_to_excel(excelFileNameB,binningInfoLong,'bin_restart', 0, True)

print('4.2. Бинирование окончено: ' + str(dtime.datetime.now()))
             
### Часть 4.2. Результаты бинирования по переменным, прошедшим проверку на IV

## Сокращаем выборку после бинирования при заданных листах (стоп-лист, лонг-лист, шорт-лист)

binningInfoLong = binningInfoLong.query('attr_true_name in @all_columns')
if len(my_long_list) > 0:
    binningInfoLong = binningInfoLong.query('attr_true_name in @my_long_list')
if len(my_short_list) > 0:
    binningInfoLong = binningInfoLong.query('attr_name in @my_short_list')
binningInfoLong = binningInfoLong.query('attr_true_name not in @stop_list_attr')

## Получение агрегированных данных по переменным с проверками по IV    

binningInfoShort = binning_get_shortlist_after_iv(binningInfoLong,IV_left_cutoff,IV_right_cutoff)
binningInfoShortList = list(binningInfoShort.query('iv_check == 1').attr_name)

## Получение данных бинирования только для переменных, прошедших IV

binningInfoLong = binningInfoLong.query('attr_name in @binningInfoShortList')
binningInfoLong = binningInfoLong.query('good + bad > 0')  

### Часть 4.3. Подготовка данных к расчету трендовости

## Получение преобразованной таблицы для расчета трендовости

binningInfoLongFin = binning_get_data_for_trend(binningInfoLong) #################################

if not use_restart_file:
    # самый первый автоматический расчет
    binningInfoLongFin = test_stat(binningInfoLongFin, df_test, person_id_attr, target_attr)
    
else:
    # используем результаты прошлого автоматического расчета
    binningInfoLongFin = download_excel_file(filepath_trend_stat_restart, 'binningInfoLongFin')
    
append_df_to_excel(excelFileNameS, binningInfoLongFin, 'binningInfoLongFin', 0, True)

## Создание лонг листа с проверками по IV, выгрузка в эксель

# джойн данных статистики по тренировочной выборке

binningInfoShortIndex = binningInfoShort.join(df_train_descr[['mean','min','max']], \
                                              on='attr_true_name', \
                                              how = 'left')

# джойн типов переменных

binningInfoShortIndex = binningInfoShortIndex.join(df_types['type'], \
                                                   on='attr_true_name', \
                                                   how = 'left')

# создание лонг-листа с индексом по модифицированному имени переменной

binningInfoShortIndex = binningInfoShortIndex.set_index('attr_name', drop=True)
long_list = binningInfoShortIndex
append_df_to_excel(excelFileName, long_list, 'long_list', 0, True)

print('4.3. Подготовлены таблицы для расчета трендовости: ' + str(dtime.datetime.now()))
print(t)

###############################################################################

## Часть 5. Проверка на трендовость

# Смотрим только переменные, прошедшие проверку на IV

trendInfo = binningInfoShort.query('iv_check == 1').reset_index(drop=True)

print('5. Проверка на трендовость: ' + str(dtime.datetime.now()))

if use_restart_trend == False: # проверка на тренд в первый раз самостоятельно

    # Обрабатываем интервальные переменные
    
    trendInfoDetail = copy.deepcopy(trendInfo)
    
    trend_flg(binningInfoLongFin, trendInfoDetail, PSI_cutoff, PSI_bin_cutoff, VOL_cutoff, VOL_bin_cutoff, bin_v_cutoff, break_trd_cutoff)
    
    trendInfoDetail['trend_detail'] = trendInfoDetail.apply(lambda x: get_Trend_detail(x['has trend'], \
                                                                                       x['PSI'], PSI_cutoff, PSI_bin_cutoff, \
                                                                                       x['VOLATILITY'], VOL_cutoff, VOL_bin_cutoff, \
                                                                                       x['bin-value <'+str(bin_v_cutoff)+'%'], bin_v_cutoff, \
                                                                                       x['no category'], \
                                                                                       x['no_events_in_bin']), axis=1)
    
    trendInfoDetail['attr_max_iv'] = trendInfoDetail.query('trend_detail == ""') \
                                                    .sort_values(by='iv_total_final',ascending = False) \
                                                    .groupby(['attr_true_name'])['iv_total_final'] \
                                                    .cumcount() + 1
                                                    
    trendInfoDetail['attr_max_iv'] = trendInfoDetail['attr_max_iv'].fillna(0)
    
    trendInfoDetail['trend_detail'] = trendInfoDetail.apply(lambda x: get_Trend_Unique_detail(x['trend_detail'], \
                                                                                              x['attr_max_iv']), axis=1)                                                                                                                 
    
    trendInfoDetail = trendInfoDetail[['attr_name','trend_detail','VOLATILITY VALUE']].set_index('attr_name')
    trendInfo = trendInfo.set_index('attr_name')
    trendInfo = pd.concat([trendInfo,trendInfoDetail], axis = 1)
    
    trendInfo = trendInfo.reset_index()
    
    trendInfo['trend_check'] = trendInfo.apply(lambda x: get_detail_check(x.trend_detail), axis=1)
else:
    trendInfo = download_excel_file(filepath_restart, sheetname_trend) 
    trendInfo['trend_detail'] = trendInfo.apply(lambda x: get_detail_from_excel(x['trend_detail']), axis=1)    
    print('5.1. Трендовость загружена из restart файла: ' + str(dtime.datetime.now()))

append_df_to_excel(excelFileNameB,trendInfo,'trend_restart', 0, True)

trendInfoIndex = trendInfo.set_index('attr_name', drop=True)
cols_to_merge = trendInfoIndex.columns.difference(long_list.columns)
long_list = pd.merge(long_list, trendInfoIndex[cols_to_merge], left_index = True, right_index = True, how = 'left')
append_df_to_excel(excelFileName, long_list, 'long_list', 0, True)

print('5.5. Отобраны переменные по трендовости: ' + str(dtime.datetime.now()))
print(t)

###############################################################################

### Часть 6. Проверка на значимость переменных (p-value)

print('6.1. Начинаем отбор переменных по значимости: ' + str(dtime.datetime.now()))

## Получаем набор переменных, прошедших проверку трендовости

pValueInfo = trendInfo.query('trend_check == 1').reset_index(drop=True)

## Переводим значения атрибутов в номера бинов

remain_attr_list = list(pValueInfo['attr_name'])
remain_true_attr_list = list(pValueInfo['attr_true_name'])
df_train = bin_df(df_train,binningInfoLongFin,remain_attr_list, var_bin_type = 'woe')
df_test  = bin_df(df_test,binningInfoLongFin,remain_attr_list, var_bin_type = 'woe')

if use_restart_pvalue == False: # проверка на pvalue в первый раз самостоятельно
    ## Проверяем переменные на проверку p-value    
    #pValueInfo = pvalue_get_check(pValueInfo,df_train,PVALUE_cutoff,target_attr) - старый расчет P-Value
    pValueRes = get_pv_sequences(df_train, remain_true_attr_list, PVALUE_thresholds, target_attr)
    pValueInfo = pValueInfo.join(pValueRes[['pvalue_detail', 'pvalue_check']], on='attr_true_name', how='left')
else:
    pValueInfo = download_excel_file(filepath_restart, sheetname_pvalue)
    pValueInfo['pvalue_detail'] = pValueInfo.apply(lambda x: get_detail_from_excel(x['pvalue_detail']), axis=1)
    print('6.2. PValue загружено из restart файла: ' + str(dtime.datetime.now()))

append_df_to_excel(excelFileNameB,pValueInfo,'pvalue_restart', 0, True)
    
## Обновляем флаги проверок в лонг-листе

pValueInfoIndex = pValueInfo.set_index('attr_name', drop=True)
cols_to_merge = pValueInfoIndex.columns.difference(long_list.columns)
long_list = pd.merge(long_list, pValueInfoIndex[cols_to_merge], left_index = True, right_index = True, how = 'left')
append_df_to_excel(excelFileName, long_list, 'long_list', 0, True)

print('6.3. Отобраны переменные по значимости: ' + str(dtime.datetime.now()))
print(t)

###############################################################################

### Часть 7. Проверка на корреляцию

## Сбор данных для проверок

correlInfo = pValueInfo.query('pvalue_check == 1').reset_index(drop=True)
correlInfo['correl_detail'] = ''

if use_restart_correl == False: # проверка на корреляцию в первый раз самостоятельно

    print('7.1. Начата проверка на корреляцию: ' + str(dtime.datetime.now()))
    
    # Решили что забинированные данные должны проверяться на корреляцию целиком как категориальные
    
    correlDataIntList = list(long_list.query('(pvalue_check == 1) & (type == "Нет")')['attr_true_name'])
    correlDataInt = df_train[correlDataIntList]
    correlDataCatList = list(long_list.query('(pvalue_check == 1)')['attr_true_name'])
    correlDataCat = df_train[correlDataCatList]
    
    correlBadAttrListInt = []
    correlBadAttrListCat = []
    
    ## Проверка интервальных переменных
    
    if correlDataInt.shape[1] > 1:
        corrIntMtrx = correlDataInt.corr(method='pearson', min_periods = 1)
        for i in correlDataIntList:
            for j in correlDataIntList:
                if i != j \
                    and i not in correlBadAttrListInt \
                    and j not in correlBadAttrListInt \
                    and abs(corrIntMtrx.loc[i][j]) > CORREL_cutoff:
                    correlBadAttrListInt.append(j)
                    correlInfo['correl_detail'] = correlInfo.apply(lambda x: get_Correl_detail(x.attr_true_name, \
                                                                                               j, \
                                                                                               x.correl_detail, \
                                                                                               i, \
                                                                                               'int', \
                                                                                               CORREL_cutoff), axis=1) 
                    
    ## Проверка категориальных переменных
    
    if correlDataCat.shape[1] > 1: 
        for i in correlDataCatList:
            for j in correlDataCatList:
                if i != j \
                    and i not in correlBadAttrListCat \
                    and j not in correlBadAttrListCat:
                        confusion_mx = pd.crosstab(correlDataCat[i],correlDataCat[j])
                        corrCatValue = cramers_corrected_stat(confusion_mx)
                        if abs(corrCatValue) > CORREL_cutoff:
                            correlBadAttrListCat.append(j)
                            correlInfo['correl_detail'] = correlInfo.apply(lambda x: get_Correl_detail(x.attr_true_name, \
                                                                                                       j, \
                                                                                                       x.correl_detail, \
                                                                                                       i, \
                                                                                                       'cat', \
                                                                                                       CORREL_cutoff), axis=1)
               
    correlInfo['correl_check'] = correlInfo.apply(lambda x: get_detail_check(x.correl_detail), axis=1)  
else:
    correlInfo = download_excel_file(filepath_restart, sheetname_correl) 
    correlInfo['correl_detail'] = correlInfo.apply(lambda x: get_detail_from_excel(x['correl_detail']), axis=1)
    print('7.1. Корреляция загружена из restart файла: ' + str(dtime.datetime.now()))

append_df_to_excel(excelFileNameB,correlInfo,'correl_restart', 0, True)

correlInfoIndex = correlInfo.set_index('attr_name', drop=True)
cols_to_merge = correlInfoIndex.columns.difference(long_list.columns)
long_list = pd.merge(long_list, correlInfoIndex[cols_to_merge], left_index = True, right_index = True, how = 'left')

append_df_to_excel(excelFileName, long_list, 'long_list', 0, True)

print('7.2. Отобраны переменные по корреляции: ' + str(dtime.datetime.now()))
print(t)

###############################################################################

### Часть 8. Отбор сбалансированных переменных, подготовка данных перед обучением

## Часть 8.1. Отбор сбалансированных переменных

long_list['balance_detail'] = ''
long_list_short = long_list.query('correl_check == 1')['iv_total_final'].reset_index()
long_list_short_2 = copy.deepcopy(long_list_short)
long_list_short_2.index -= 1
long_list_short_3 = copy.deepcopy(long_list_short)
long_list_short_3.index += 1
long_list_short = long_list_short.join(long_list_short_2['iv_total_final'], rsuffix = '_pred')
long_list_short = long_list_short.join(long_list_short_3['iv_total_final'], rsuffix = '_next')
long_list_short['iv_gap_pred'] = long_list_short['iv_total_final'] - long_list_short['iv_total_final_pred']
long_list_short['iv_gap_next'] = long_list_short['iv_total_final_next'] - long_list_short['iv_total_final']
long_list_short['iv_int'] = 0
long_list_short = long_list_short.fillna(100)

short_list_intervals = pd.DataFrame({'left_interval': [0.05, 0.2, 0.4, 0.6, 0.8], \
                                     'right_interval': [0.2, 0.4, 0.6, 0.8, 1]})
short_list_intervals['cnt'] = 0

for i, m in long_list_short.iterrows():
    for j, k in short_list_intervals.iterrows():       
        if m.iv_total_final >= k.left_interval and m.iv_total_final < k.right_interval:
            short_list_intervals.at[j, 'cnt'] += 1 
            long_list_short.at[i,'iv_int'] = j
            break
            
long_list_short = long_list_short.join(short_list_intervals['cnt'], on = 'iv_int')

for i, j in long_list_short.iterrows():
    if j.cnt == 1 and j.iv_gap_pred > 0.25 and j.iv_gap_next > 0.25:
        long_list.at[j.attr_name, 'balance_detail'] = '5. Переменная ' \
        + j.attr_name + ' несбалансирована по IV.'
        
long_list['balance_check'] = long_list.apply(lambda x: get_detail_check(x.balance_detail, \
                                                                        x.correl_check), axis=1)
append_df_to_excel(excelFileName, long_list, 'long_list', 0, True)

print('8.1. Отобраны переменные по сбалансированности: ' + str(dtime.datetime.now()))

## Часть 8.2. Подготовка данных перед обучением модели

## обработка лонг листа переменных, агрегированные лонг-листы

long_list['final_check'] = long_list.apply(lambda x: long_list_final_check(x.iv_check, \
                                                                           x.trend_check, \
                                                                           x.pvalue_check, \
                                                                           x.correl_check, \
                                                                           x.balance_check), axis=1)  

long_list['detalization'] = long_list.apply(lambda x: long_list_final_detail(x.iv_detail, \
                                                                             x.trend_detail, \
                                                                             x.pvalue_detail, \
                                                                             x.correl_detail, \
                                                                             x.balance_detail), axis=1)

long_list = long_list[['attr_true_name',
                       'type',
                       'min_perc_total',
                       'min_perc_class',
                       'stop_limit',
                       'par_loop',
                       #'mean',
                       #'min',
                       #'max',
                       'iv_total_final', 
                       'VOLATILITY VALUE',
                       'iv_check',                       
                       'trend_check', 
                       'pvalue_check',
                       'correl_check',
                       'balance_check',
                       'final_check',
                       'detalization'
                     ]].fillna(0)

long_list.rename(columns = {'iv_total_final':'iv_value', \
                            'VOLATILITY VALUE':'volatility_value'}, inplace = True)

long_list = long_list.sort_values(by = ['final_check','iv_value', 'par_loop'], ascending = [False, False, True])

append_df_to_excel(excelFileName,long_list,'long_list', 0, True)

## Статистика по лонг листу в разрезе детализации

long_list_agr_attr = long_list['detalization'].value_counts() \
                                              .reset_index()                                                                                           
long_list_agr_attr.columns = ['detalization','Количество']  
long_list_agr_attr = long_list_agr_attr.sort_values(by = 'detalization', ascending = True) \
                                       .reset_index(drop = True)
  
## Статистика по лонг листу в разрезе детализации и параметров                                
  
long_list_agr_params = long_list[['min_perc_total', 'min_perc_class', 'stop_limit', 'par_loop', 'detalization']] \
                                .groupby(['min_perc_total', 'min_perc_class', 'stop_limit', 'par_loop','detalization']) \
                                .size() \
                                .reset_index()                                
long_list_agr_params.columns = ['min_perc_total', 'min_perc_class', 'stop_limit', 'par_loop', 'detalization', 'Количество']  
long_list_agr_params = long_list_agr_params.sort_values(by = ['min_perc_total', 'min_perc_class', 'stop_limit', 'par_loop'], ascending = [True, True, True, True]) \
                                           .reset_index(drop = True)
                                           
## Подготовка датасетов только с финальными атрибутами

short_list_attrs = list(long_list.query('final_check == 1') \
                                 .index \
                                 .unique())

short_list_true_attrs = list(long_list.query('final_check == 1') \
                                      .attr_true_name \
                                      .unique())

short_list_true_attrs_target = short_list_true_attrs
short_list_true_attrs_target.append(target_attr)

df_train = df_train[short_list_true_attrs_target]
df_test = df_test[short_list_true_attrs_target]

## Матрицы корреляции

# Решили что забинированные данные должны проверяться на корреляцию целиком как категориальные

correlDataIntListF = list(long_list.query('(final_check == 1) & (type == "Нет")')['attr_true_name'])
correlDataCatListF = list(long_list.query('(final_check == 1)')['attr_true_name'])

# Тренировочная

corrTrainMtrx = cat_variable_corr(correlDataCatListF, df_train, CORREL_cutoff)

'''
# Тренировочная интервальная

correlTrainInt = df_train[correlDataIntListF]

if correlTrainInt.shape[1] > 0: 
    corrTrainMtrxInt = correlTrainInt.corr(method='pearson', min_periods = 1)
    append_df_to_excel(excelFileName,corrTrainMtrxInt,'corr_mtrx_int_train', 0, True) 
    
# Тестовая интервальная
    
correlTestInt = df_test[correlDataIntListF]
    
if correlTestInt.shape[1] > 0: 
    corrTestMtrxInt = correlTestInt.corr(method='pearson', min_periods = 1)
    # append_df_to_excel(excelFileName,corrTestMtrxInt,'corr_mtrx_int_test', 0, True)
    
# Тренировочная категориальная
    
correlTrainCat = df_train[correlDataCatListF]
    
if correlTrainCat.shape[1] > 0:
    correlTrainCatList = list(correlTrainCat)
    corrTrainMtrxCat = pd.DataFrame(correlTrainCatList, columns=['attr'])
    for attr_name in correlTrainCatList:
        corrTrainMtrxCat[attr_name] = 0.0000
    corrTrainMtrxCat = corrTrainMtrxCat.set_index('attr', drop=True)
    for i in correlTrainCatList:
        for j in correlTrainCatList:
            confusion_mx = pd.crosstab(correlTrainCat[i],correlTrainCat[j])
            corrCatValue = cramers_corrected_stat(confusion_mx)
            corrTrainMtrxCat.at[i,j] = corrCatValue
    append_df_to_excel(excelFileName,corrTrainMtrxCat,'corr_mtrx_cat_train', 0, True)
    
# Тестовая категориальная
    
correlTestCat = df_test[correlDataCatListF]
    
if correlTestCat.shape[1] > 0:
    correlTestCatList = list(correlTestCat)
    corrTestMtrxCat = pd.DataFrame(correlTestCatList, columns=['attr'])
    for attr_name in correlTestCatList:
        corrTestMtrxCat[attr_name] = 0.0000
    corrTestMtrxCat = corrTestMtrxCat.set_index('attr', drop=True)
    for i in correlTestCatList:
        for j in correlTestCatList:
            confusion_mx = pd.crosstab(correlTestCat[i],correlTestCat[j])
            corrCatValue = cramers_corrected_stat(confusion_mx)
            corrTestMtrxCat.at[i,j] = corrCatValue
    # append_df_to_excel(excelFileName,corrTestMtrxCat,'corr_mtrx_cat_test', 0, True)
'''

print('8.2. Подготовлены данные для обучения модели: ' + str(dtime.datetime.now()))
print(t)

###############################################################################

### Часть 9. Обучение модели
    
## Составляем вектора с переменными и ЦФ для обучения и тестирования
    
X_train = df_train.drop(target_attr, axis = 1)
y_train = df_train[target_attr]
X_test = df_test.drop(target_attr, axis = 1)
y_test = df_test[target_attr]

## Создаем таблицу с результатами обучения

validation_results = pd.DataFrame(columns=['model_name', 'gini_train', 'gini_test'])

## Обучение:

print('9. Модель начала обучение: ' + str(dtime.datetime.now()))

## Наименования моделей

name_0  = 'Логистическая регрессия "LogisticRegression"'
name_1  = 'Логистическая регрессия "LogisticRegressionCV"'
name_2  = 'Метод опорных векторов "SVC"'
name_3  = 'Метод опорных векторов "NuSVC"'
name_4  = 'Метод опорных векторов "LinearSVC"'
name_5  = 'Стохастический градиентный спуск "SGDClassifier"'
name_6  = 'Классификатор ближайших соседей "KNeighborsClassifier"'
name_7  = 'Классификатор ближайших соседей "RadiusNeighborsClassifier"'
name_8  = 'Гаусовский процесс "GaussianProcessClassifier"'
name_9  = 'Наивный байесовский классификатор "GaussianNB"'
name_10 = 'Наивный байесовский классификатор "MultinomialNB"'
name_11 = 'Наивный байесовский классификатор "ComplementNB"'
name_12 = 'Наивный байесовский классификатор "CategoricalNB"'
name_13 = 'Наивный байесовский классификатор "BernoulliNB"'
name_14 = 'Деревья решений "DecisionTreeClassifier"'
name_15 = 'Ансамбль "BaggingClassifier"'
name_16 = 'Ансамбль "RandomForestClassifier"'
name_17 = 'Ансамбль "ExtraTreesClassifier"'
name_18 = 'Ансамбль "AdaBoostClassifier"'
name_19 = 'Ансамбль "GradientBoostingClassifier"'
name_20 = 'Ансамбль "VotingClassifier"'
name_21 = 'Ансамбль "StackingClassifier"'
name_22 = 'Нейронные сети "MLPClassifier"'

## 9.0.  Логистическая регрессия "LogisticRegression"

if model_type_0 == True:

    print('9.0.1. Обучение # ' + name_0 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_0 = LogisticRegression() # LogisticRegressionCV
    param_0 = {'solver'        : ['lbfgs'], #['newton-cg', 'lbfgs', 'liblinear', 'sag', 'saga'],
               'tol'           : [0.0001], #[0.000001, 0.00001, 0.0001, 0.001, 0.01],
               'max_iter'      : [100] #[10, 100, 1000]
               }
    model_0 = run_randomsearch(X_train,y_train,clf_0,param_0,5,1,1)
    
    model_0_list_X_train = list(X_train)
    model_0_intercept = model_0.best_estimator_.intercept_
    model_0_coef = model_0.best_estimator_.coef_
    
    y_train_predict_proba_0 = model_0.predict_proba(X_train)[:, 1]
    y_test_predict_proba_0 = model_0.predict_proba(X_test)[:, 1]
    gini_train_0 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_0) - 1,4)
    gini_test_0 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_0) - 1,4)
    validation_results.loc[0] = [name_0,gini_train_0,gini_test_0]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.0.2. Обучение # ' + name_0 + ' # Конец: ' + str(dtime.datetime.now()))

## 9.1.  Логистическая регрессия "LogisticRegressionCV"

if model_type_1 == True:

    print('9.1.1. Обучение # ' + name_1 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_1 = LogisticRegressionCV()
    param_1 = {'solver'        : ['lbfgs'], #['newton-cg', 'lbfgs', 'liblinear', 'sag', 'saga'],
               'tol'           : [0.0001], #[0.000001, 0.00001, 0.0001, 0.001, 0.01],
               'max_iter'      : [100] #[10, 100, 1000]
               }
    model_1 = run_randomsearch(X_train,y_train,clf_1,param_1,5,1,1)
    
    model_1_list_X_train = list(X_train)
    model_1_intercept = model_1.best_estimator_.intercept_
    model_1_coef = model_1.best_estimator_.coef_
    
    y_train_predict_proba_1 = model_1.predict_proba(X_train)[:, 1]
    y_test_predict_proba_1 = model_1.predict_proba(X_test)[:, 1]
    gini_train_1 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_1) - 1,4)
    gini_test_1 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_1) - 1,4)
    validation_results.loc[1] = [name_1,gini_train_1,gini_test_1]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.1.2. Обучение # ' + name_1 + ' # Конец: ' + str(dtime.datetime.now()))
      
## 9.2.  Метод опорных векторов "SVC"
    
if model_type_2 == True:

    print('9.2.1. Обучение # ' + name_2 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_2 = svm.SVC()
    param_2 = {'kernel'        : ['linear'],
               'probability'   : [True],         
               'tol'           : [0.0001, 0.001, 0.01, 0.1, 1],
               'C'             : [0.98, 0.99, 1.0, 1.01, 1.02]           
               }
    model_2 = run_randomsearch(X_train,y_train,clf_2,param_2,5,25,1)
    
    # print(model_2.best_estimator_.intercept_) # свободный член
    # print(model_2.best_estimator_.coef_) # коэффициенты
    
    y_train_predict_proba_2 = model_2.predict_proba(X_train)[:, 1]
    y_test_predict_proba_2 = model_2.predict_proba(X_test)[:, 1]
    gini_train_2 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_2) - 1,4)
    gini_test_2 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_2) - 1,4)
    validation_results.loc[2] = [name_2,gini_train_2,gini_test_2]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.2.2. Обучение # ' + name_2 + ' # Конец: ' + str(dtime.datetime.now()))
      
## 9.3.  Метод опорных векторов "NuSVC"
# Пока не работает - жалуется что Nu Is Infeasible 
# https://scikit-learn.org/stable/modules/generated/sklearn.svm.NuSVC.html

if model_type_3 == True:

    print('9.3.1. Обучение # ' + name_3 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_3 = svm.NuSVC()
    param_3 = {'kernel'        : ['linear'],
               'probability'   : [True],         
               'tol'           : [0.0001, 0.001, 0.01, 0.1, 1],
               'nu'            : [0.3, 0.4, 0.5, 0.6, 0.7]           
               }
    # model_3 = run_randomsearch(X_train,y_train,clf_3,param_3,5,25,1)
    
    # print(model_3.best_estimator_.intercept_) # свободный член
    # print(model_3.best_estimator_.coef_) # коэффициенты
    
    # y_train_predict_proba_3 = model_3.predict_proba(X_train)[:, 1]
    # y_test_predict_proba_3 = model_3.predict_proba(X_test)[:, 1]
    gini_train_3 = -100 #np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_3) - 1,4)
    gini_test_3 = -100 #np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_3) - 1,4)
    validation_results.loc[3] = [name_3,gini_train_3,gini_test_3]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.3.2. Обучение # ' + name_3 + ' # Конец: ' + str(dtime.datetime.now()))

## 9.4.  Метод опорных векторов "LinearSVC"
    
if model_type_4 == True:

    print('9.4.1. Обучение # ' + name_4 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_4 = svm.LinearSVC()
    param_4 = {'penalty'       : ['l2'],
               'loss'          : ['hinge', 'squared_hinge'],
               'tol'           : [0.0001, 0.001, 0.01, 0.1, 1],
               'C'             : [0.98, 0.99, 1.0, 1.01, 1.02]           
               }
    model_4_base = run_randomsearch(X_train,y_train,clf_4,param_4,5,25,1)
            
    clf_4 = CalibratedClassifierCV(base_estimator = model_4_base.best_estimator_,
                                   cv = 'prefit')
    
    model_4 = clf_4.fit(X_train, y_train)
    
    # print(model_4_base.best_estimator_.intercept_) # свободный член
    # print(model_4_base.best_estimator_.coef_) # коэффициенты
    
    y_train_predict_proba_4 = model_4.predict_proba(X_train)[:, 1]
    y_test_predict_proba_4 = model_4.predict_proba(X_test)[:, 1]
    gini_train_4 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_4) - 1,4)
    gini_test_4 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_4) - 1,4)
    validation_results.loc[4] = [name_4,gini_train_4,gini_test_4]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.4.2. Обучение # ' + name_4 + ' # Конец: ' + str(dtime.datetime.now()))
  
## 9.5.  Стохастический градиентный спуск "SGDClassifier"

if model_type_5 == True:

    print('9.5.1. Обучение # ' + name_5 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_5 = SGDClassifier()
    
    scalerI = StandardScaler()
    X_train_scale = scalerI.fit_transform(X_train)
    X_test_scale = scalerI.transform(X_test)
    
    param_5 = {'loss'        : ['hinge','log', 'modified_huber', 'perceptron'],
               'penalty'     : ['l2', 'l1', 'elasticnet'],         
               'tol'         : [0.00001, 0.0001, 0.001],
               'alpha'       : [0.00001, 0.0001, 0.001]           
               }
    model_5 = run_randomsearch(X_train_scale,y_train,clf_5,param_5,5,25,1)
    
    # print(model_5.best_estimator_.intercept_) # свободный член
    # print(model_5.best_estimator_.coef_) # коэффициенты
    
    y_train_predict_proba_5 = model_5.predict_proba(X_train_scale)[:, 1]
    y_test_predict_proba_5 = model_5.predict_proba(X_test_scale)[:, 1]
    gini_train_5 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_5) - 1,4)
    gini_test_5 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_5) - 1,4)
    validation_results.loc[5] = [name_5,gini_train_5,gini_test_5]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.5.2. Обучение # ' + name_5 + ' # Конец: ' + str(dtime.datetime.now()))
     
## 9.6.  Классификатор ближайших соседей "KNeighborsClassifier"

if model_type_6 == True:

    print('9.6.1. Обучение # ' + name_6 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_6 = KNeighborsClassifier()
    param_6 = {'n_neighbors'   : [2,3,4,5,6,7,8],
               'weights'       : ['uniform','distance'],         
               'algorithm'     : ['auto','ball_tree','kd_tree','brute'],
               'leaf_size'     : [25,30,35]           
               }
    model_6 = run_randomsearch(X_train,y_train,clf_6,param_6,5,5,1)
    
    # print(model_6.best_estimator_.intercept_) # свободный член
    # print(model_6.best_estimator_.coef_) # коэффициенты
    
    y_train_predict_proba_6 = model_6.predict_proba(X_train)[:, 1]
    y_test_predict_proba_6 = model_6.predict_proba(X_test)[:, 1]
    gini_train_6 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_6) - 1,4)
    gini_test_6 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_6) - 1,4)
    validation_results.loc[6] = [name_6,gini_train_6,gini_test_6]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.6.2. Обучение # ' + name_6 + ' # Конец: ' + str(dtime.datetime.now()))

## 9.7.  Классификатор ближайших соседей "RadiusNeighborsClassifier"
# Пока не работает - жалуется на нехватку predict_proba
# https://scikit-learn.org/stable/modules/generated/sklearn.neighbors.RadiusNeighborsClassifier.html#sklearn.neighbors.RadiusNeighborsClassifier

if model_type_7 == True:
      
    print('9.7.1. Обучение # ' + name_7 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_7 = RadiusNeighborsClassifier()
    param_7 = {'radius'        : [0.8, 0.9, 1, 1.1, 1.2],
               'weights'       : ['uniform','distance'],         
               'algorithm'     : ['auto','ball_tree','kd_tree','brute'],
               'leaf_size'     : [25,30,35]           
               }
    # model_7 = run_randomsearch(X_train,y_train,clf_7,param_7,5,25,1)
    
    # print(model_7.best_estimator_.intercept_) # свободный член
    # print(model_7.best_estimator_.coef_) # коэффициенты
    
    # y_train_predict_proba_7 = model_7.predict_proba(X_train)[:, 1]
    # y_test_predict_proba_7 = model_7.predict_proba(X_test)[:, 1]
    gini_train_7 = -100 # np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_7) - 1,4)
    gini_test_7 = -100 # np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_7) - 1,4)
    validation_results.loc[7] = [name_7,gini_train_7,gini_test_7]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.7.2. Обучение # ' + name_7 + ' # Конец: ' + str(dtime.datetime.now()))
    
## 9.8.  Гаусовский процесс "GaussianProcessClassifier"
# Пока не работает - memory error
    
if model_type_8 == True:          
          
    print('9.8.1. Обучение # ' + name_8 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_8 = GaussianProcessClassifier()
    kernel = 1.0 * RBF(1.0)
    param_8 = {'kernel'               : [kernel],
               'max_iter_predict'     : [50,100,150],         
               'n_restarts_optimizer' : [0]          
               }
    # model_8 = run_randomsearch(X_train,y_train,clf_8,param_8,5,3,1)
    
    # print(model_8.best_estimator_.intercept_) # свободный член
    # print(model_8.best_estimator_.coef_) # коэффициенты
    
    # y_train_predict_proba_8 = model_8.predict_proba(X_train)[:, 1]
    # y_test_predict_proba_8 = model_8.predict_proba(X_test)[:, 1]
    gini_train_8 = -100 #np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_8) - 1,4)
    gini_test_8 = -100 #np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_8) - 1,4)
    validation_results.loc[8] = [name_8,gini_train_8,gini_test_8]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.8.2. Обучение # ' + name_8 + ' # Конец: ' + str(dtime.datetime.now()))
   
## 9.9.  Наивный байесовский классификатор "GaussianNB"
    
if model_type_9 == True:          

    print('9.9.1. Обучение # ' + name_9 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_9 = GaussianNB()
    param_9 = {}
    model_9 = run_randomsearch(X_train,y_train,clf_9,param_9,5,1,1)
    
    # print(model_9.best_estimator_.intercept_) # свободный член
    # print(model_9.best_estimator_.coef_) # коэффициенты
    
    y_train_predict_proba_9 = model_9.predict_proba(X_train)[:, 1]
    y_test_predict_proba_9 = model_9.predict_proba(X_test)[:, 1]
    gini_train_9 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_9) - 1,4)
    gini_test_9 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_9) - 1,4) 
    validation_results.loc[9] = [name_9,gini_train_9,gini_test_9]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.9.2. Обучение # ' + name_9 + ' # Конец: ' + str(dtime.datetime.now()))    

## 9.10. Наивный байесовский классификатор "MultinomialNB"
    
if model_type_10 == True:

    print('9.10.1. Обучение # ' + name_10 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_10 = MultinomialNB()
    param_10 = {'alpha' : [0, 0.5, 1, 1.5, 2]          
               }
    model_10 = run_randomsearch(X_train,y_train,clf_10,param_10,5,5,1)
    
    # print(model_10.best_estimator_.intercept_) # свободный член
    # print(model_10.best_estimator_.coef_) # коэффициенты
    
    y_train_predict_proba_10 = model_10.predict_proba(X_train)[:, 1]
    y_test_predict_proba_10 = model_10.predict_proba(X_test)[:, 1]
    gini_train_10 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_10) - 1,4)
    gini_test_10 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_10) - 1,4)
    validation_results.loc[10] = [name_10,gini_train_10,gini_test_10]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.10.2. Обучение # ' + name_10 + ' # Конец: ' + str(dtime.datetime.now()))      

## 9.11. Наивный байесовский классификатор "ComplementNB" 
# Пока не работает
# Отсутствует в версии питона 3.6
      
# https://scikit-learn.org/stable/modules/generated/sklearn.naive_bayes.ComplementNB.html#sklearn.naive_bayes.ComplementNB 
# from sklearn.naive_bayes import ComplementNB 
    
if model_type_11 == True:          

    print('9.11.1. Обучение # ' + name_11 + ' # Начало: ' + str(dtime.datetime.now()))
    
    # clf_11 = ComplementNB()
    param_11 = {'alpha' : [0, 0.5, 1, 1.5, 2],
                'norm'  : [True, False]
               }
    # model_11 = run_randomsearch(X_train,y_train,clf_11,param_11,5,10,1)
    
    # print(model_11.best_estimator_.intercept_) # свободный член
    # print(model_11.best_estimator_.coef_) # коэффициенты
    
    # y_train_predict_proba_11 = model_11.predict_proba(X_train)[:, 1]
    # y_test_predict_proba_11 = model_11.predict_proba(X_test)[:, 1]
    gini_train_11 = -100 # np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_11) - 1,4)
    gini_test_11 = -100 # np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_11) - 1,4)
    validation_results.loc[11] = [name_11,gini_train_11,gini_test_11]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.11.2. Обучение # ' + name_11 + ' # Конец: ' + str(dtime.datetime.now())) 

## 9.12. Наивный байесовский классификатор "CategoricalNB"    
# Пока не работает
# Отсутствует в версии питона 3.6
      
# https://scikit-learn.org/stable/modules/generated/sklearn.naive_bayes.CategoricalNB.html#sklearn.naive_bayes.CategoricalNB
# from sklearn.naive_bayes import CategoricalNB
    
if model_type_12 == True:          

    print('9.12.1. Обучение # ' + name_12 + ' # Начало: ' + str(dtime.datetime.now()))
    
    # clf_12 = CategoricalNB()
    param_12 = {'alpha' : [0, 0.5, 1, 1.5, 2]
               }
    # model_12 = run_randomsearch(X_train,y_train,clf_12,param_12,5,5,1)
    
    # print(model_12.best_estimator_.intercept_) # свободный член
    # print(model_12.best_estimator_.coef_) # коэффициенты
    
    # y_train_predict_proba_12 = model_12.predict_proba(X_train)[:, 1]
    # y_test_predict_proba_12 = model_12.predict_proba(X_test)[:, 1]
    gini_train_12 = -100 # np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_12) - 1,4)
    gini_test_12 = -100 # np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_12) - 1,4)
    validation_results.loc[12] = [name_12,gini_train_12,gini_test_12]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.12.2. Обучение # ' + name_12 + ' # Конец: ' + str(dtime.datetime.now()))      
      
## 9.13. Наивный байесовский классификатор "BernoulliNB"
    
if model_type_13 == True:          

    print('9.13.1. Обучение # ' + name_13 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_13 = BernoulliNB()
    param_13 = {'alpha' : [0, 0.5, 1, 1.5, 2]
               }
    model_13 = run_randomsearch(X_train,y_train,clf_13,param_13,5,5,1)
    
    # print(model_13.best_estimator_.intercept_) # свободный член
    # print(model_13.best_estimator_.coef_) # коэффициенты
    
    y_train_predict_proba_13 = model_13.predict_proba(X_train)[:, 1]
    y_test_predict_proba_13 = model_13.predict_proba(X_test)[:, 1]
    gini_train_13 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_13) - 1,4)
    gini_test_13 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_13) - 1,4)
    validation_results.loc[13] = [name_13,gini_train_13,gini_test_13]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.13.2. Обучение # ' + name_13 + ' # Конец: ' + str(dtime.datetime.now()))

## 9.14. Деревья решений "DecisionTreeClassifier"
    
if model_type_14 == True:

    print('9.14.1. Обучение # ' + name_14 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_14 = DecisionTreeClassifier()
    param_14 = {'criterion'         : ['gini','entropy'],
                'max_depth'         : [3,5,7,9,11,13,15,17,19],
                'min_samples_split' : [2,3,4],
                'min_samples_leaf'  : [1,2,3]            
               }
    model_14 = run_randomsearch(X_train,y_train,clf_14,param_14,5,100,1)
    
    # print(model_14.best_estimator_.feature_importances_) # коэффициенты
    
    y_train_predict_proba_14 = model_14.predict_proba(X_train)[:, 1]
    y_test_predict_proba_14 = model_14.predict_proba(X_test)[:, 1]
    gini_train_14 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_14) - 1,4)
    gini_test_14 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_14) - 1,4)
    validation_results.loc[14] = [name_14,gini_train_14,gini_test_14]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.14.2. Обучение # ' + name_14 + ' # Конец: ' + str(dtime.datetime.now()))

## 9.15. Ансамбль "BaggingClassifier"

if model_type_15 == True:    

    print('9.15.1. Обучение # ' + name_15 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_15 = BaggingClassifier()
    param_15 = {'base_estimator' : [DecisionTreeClassifier(), LogisticRegressionCV()],
                'n_estimators'   : [5,10,15]            
               }
    model_15 = run_randomsearch(X_train,y_train,clf_15,param_15,5,6,1)
    
    # print(model_15.best_estimator_.intercept_) # свободный член
    # print(model_15.best_estimator_.coef_) # коэффициенты 
    
    y_train_predict_proba_15 = model_15.predict_proba(X_train)[:, 1]
    y_test_predict_proba_15 = model_15.predict_proba(X_test)[:, 1]
    gini_train_15 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_15) - 1,4)
    gini_test_15 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_15) - 1,4)
    validation_results.loc[15] = [name_15,gini_train_15,gini_test_15]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.15.2. Обучение # ' + name_15 + ' # Конец: ' + str(dtime.datetime.now())) 

## 9.16. Ансамбль "RandomForestClassifier"
    
if model_type_16 == True:          

    print('9.16.1. Обучение # ' + name_16 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_16 = RandomForestClassifier()
    param_16 = {'n_estimators'      : [50, 100, 150],
                'criterion'         : ['gini','entropy'],
                'max_depth'         : [3,5,7,9,11,13,15,17,19],
                'min_samples_split' : [2,3,4],
                'min_samples_leaf'  : [1,2,3]            
               }
    model_16 = run_randomsearch(X_train,y_train,clf_16,param_16,5,100,1)
    
    # print(model_16.best_estimator_.feature_importances_) # коэффициенты
    
    y_train_predict_proba_16 = model_16.predict_proba(X_train)[:, 1]
    y_test_predict_proba_16 = model_16.predict_proba(X_test)[:, 1]
    gini_train_16 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_16) - 1,4)
    gini_test_16 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_16) - 1,4)
    validation_results.loc[16] = [name_16,gini_train_16,gini_test_16]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.16.2. Обучение # ' + name_16 + ' # Конец: ' + str(dtime.datetime.now()))
      
## 9.17. Ансамбль "ExtraTreesClassifier"
    
if model_type_17 == True:          

    print('9.17.1. Обучение # ' + name_17 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_17 = ExtraTreesClassifier()
    param_17 = {'n_estimators'      : [50, 100, 150],
                'criterion'         : ['gini','entropy'],
                'max_depth'         : [3,5,7,9,11,13,15,17,19],
                'min_samples_split' : [2,3,4],
                'min_samples_leaf'  : [1,2,3]      
               }
    model_17 = run_randomsearch(X_train,y_train,clf_17,param_17,5,100,1)
    
    # print(model_17.best_estimator_.intercept_) # свободный член
    # print(model_17.best_estimator_.coef_) # коэффициенты 
    
    y_train_predict_proba_17 = model_17.predict_proba(X_train)[:, 1]
    y_test_predict_proba_17 = model_17.predict_proba(X_test)[:, 1]
    gini_train_17 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_17) - 1,4)
    gini_test_17 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_17) - 1,4)
    validation_results.loc[17] = [name_17,gini_train_17,gini_test_17]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.17.2. Обучение # ' + name_17 + ' # Конец: ' + str(dtime.datetime.now()))

## 9.18. Ансамбль "AdaBoostClassifier"
    
if model_type_18 == True:          

    print('9.18.1. Обучение # ' + name_18 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_18 = AdaBoostClassifier()
    param_18 = {'n_estimators' : [50, 100, 150, 200],
                'learning_rate' : [0.8, 0.9, 1, 1.1, 1.2]
               }
    model_18 = run_randomsearch(X_train,y_train,clf_18,param_18,5,20,1)
    
    # print(model_18.best_estimator_.intercept_) # свободный член
    # print(model_18.best_estimator_.coef_) # коэффициенты 
    
    y_train_predict_proba_18= model_18.predict_proba(X_train)[:, 1]
    y_test_predict_proba_18 = model_18.predict_proba(X_test)[:, 1]
    gini_train_18 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_18) - 1,4)
    gini_test_18 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_18) - 1,4)
    validation_results.loc[18] = [name_18,gini_train_18,gini_test_18]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.18.2. Обучение # ' + name_18 + ' # Конец: ' + str(dtime.datetime.now())) 

## 9.19. Ансамбль "GradientBoostingClassifier"
    
if model_type_19 == True:          

    print('9.19.1. Обучение # ' + name_19 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_19 = GradientBoostingClassifier()
    param_19 = {'learning_rate'     : [0.01, 0.1, 1],
                'n_estimators'      : [50, 100, 150],
                'max_depth'         : [3,5,7,9,11,13,15,17,19],
                'min_samples_split' : [2,3,4],
                'min_samples_leaf'  : [1,2,3]
               }
    model_19 = run_randomsearch(X_train,y_train,clf_19,param_19,5,50,1)
    
    # print(model_19.best_estimator_.intercept_) # свободный член
    # print(model_19.best_estimator_.coef_) # коэффициенты 
    
    y_train_predict_proba_19 = model_19.predict_proba(X_train)[:, 1]
    y_test_predict_proba_19 = model_19.predict_proba(X_test)[:, 1]
    gini_train_19 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_19) - 1,4)
    gini_test_19 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_19) - 1,4)
    validation_results.loc[19] = [name_19,gini_train_19,gini_test_19]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.19.2. Обучение # ' + name_19 + ' # Конец: ' + str(dtime.datetime.now()))

## 9.20. Ансамбль "VotingClassifier"
    
if model_type_20 == True:          

    print('9.20.1. Обучение # ' + name_20 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_20 = VotingClassifier(estimators = [('lr', DecisionTreeClassifier()), ('rf', LogisticRegressionCV())])
    param_20 = {'voting' : ['soft']
               }
    model_20 = run_randomsearch(X_train,y_train,clf_20,param_20,5,1,1)
    
    # print(model_20.best_estimator_.intercept_) # свободный член
    # print(model_20.best_estimator_.coef_) # коэффициенты 
    
    y_train_predict_proba_20 = model_20.predict_proba(X_train)[:, 1]
    y_test_predict_proba_20 = model_20.predict_proba(X_test)[:, 1]
    gini_train_20 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_20) - 1,4)
    gini_test_20 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_20) - 1,4)
    validation_results.loc[20] = [name_20,gini_train_20,gini_test_20]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.20.2. Обучение # ' + name_20 + ' # Конец: ' + str(dtime.datetime.now()))

## 9.21. Ансамбль "StackingClassifier" 
# Пока не работает
# Отсутствует в версии питона 3.6
      
# https://scikit-learn.org/stable/modules/generated/sklearn.ensemble.StackingClassifier.html#sklearn.ensemble.StackingClassifier
# from sklearn.ensemble import StackingClassifier
    
if model_type_21 == True:          

    print('9.21.1. Обучение # ' + name_21 + ' # Начало: ' + str(dtime.datetime.now()))
    
    # clf_21 = StackingClassifier()
    param_21 = {}
    # model_21 = run_randomsearch(X_train,y_train,clf_21,param_21,5,1,1)
    
    # print(model_21.best_estimator_.intercept_) # свободный член
    # print(model_21.best_estimator_.coef_) # коэффициенты
    
    # y_train_predict_proba_21 = model_21.predict_proba(X_train)[:, 1]
    # y_test_predict_proba_21 = model_21.predict_proba(X_test)[:, 1]
    gini_train_21 = -100 # np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_21) - 1,4)
    gini_test_21 = -100 # np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_21) - 1,4)
    validation_results.loc[21] = [name_21,gini_train_21,gini_test_21]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)  
    
    print('9.21.2. Обучение # ' + name_21 + ' # Конец: ' + str(dtime.datetime.now()))     
     
## 9.22. Нейронные сети "MLPClassifier"
    
if model_type_22 == True:          

    print('9.22.1. Обучение # ' + name_22 + ' # Начало: ' + str(dtime.datetime.now()))
    
    clf_22 = MLPClassifier()
    param_22 = {'activation'         : ['identity','logistic','tanh','relu'],
                'solver'             : ['lbfgs','sgd','adam'],
                'alpha'              : [0.00001, 0.0001, 0.001],
                'learning_rate_init' : [0.0001, 0.001, 0.01],
                'tol'                : [0.00001, 0.0001, 0.001]
               }
    model_22 = run_randomsearch(X_train,y_train,clf_22,param_22,5,50,1)
    
    # print(model_22.best_estimator_.intercept_) # свободный член
    # print(model_22.best_estimator_.coef_) # коэффициенты 
    
    y_train_predict_proba_22= model_22.predict_proba(X_train)[:, 1]
    y_test_predict_proba_22 = model_22.predict_proba(X_test)[:, 1]
    gini_train_22 = np.around(2 * metrics.roc_auc_score(y_train,y_train_predict_proba_22) - 1,4)
    gini_test_22 = np.around(2 * metrics.roc_auc_score(y_test,y_test_predict_proba_22) - 1,4)
    validation_results.loc[22] = [name_22,gini_train_22,gini_test_22]
    append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
    
    print('9.22.2. Обучение # ' + name_22 + ' # Конец: ' + str(dtime.datetime.now())) 
      
print('9. Модель обучена: ' + str(dtime.datetime.now()))
print(t)

###############################################################################

### Часть 10. Валидация модели

## Сортируем таблицу с результатами обучения

validation_results['gini_gap'] = validation_results.apply(lambda x: get_gini_gap(x.gini_train,
                                                                                 x.gini_test), axis=1)
validation_results['gini_gap_flg'] = validation_results.apply(lambda x: get_gini_gap_flg(x.gini_train,
                                                                                         x.gini_test), axis=1)
validation_results = validation_results.sort_values(by = ['gini_gap_flg','gini_test','gini_train'], ascending = [False,False,False]) \
                                       .reset_index(drop = True)

#print(validation_results)
#print(t)

## Выбор лучшей модели в качестве финального результата

if validation_results.at[0,'model_name'] == name_0:
    y_train_predict_proba = y_train_predict_proba_0
    y_test_predict_proba = y_test_predict_proba_0
    model_list_X_train = model_0_list_X_train
    model_intercept = model_0_intercept
    model_coef = model_0_coef
elif validation_results.at[0,'model_name'] == name_1:
    y_train_predict_proba = y_train_predict_proba_1
    y_test_predict_proba = y_test_predict_proba_1
    model_list_X_train = model_1_list_X_train
    model_intercept = model_1_intercept
    model_coef = model_1_coef
elif validation_results.at[0,'model_name'] == name_2:
    y_train_predict_proba = y_train_predict_proba_2
    y_test_predict_proba = y_test_predict_proba_2
elif validation_results.at[0,'model_name'] == name_3:
    y_train_predict_proba = y_train_predict_proba_3
    y_test_predict_proba = y_test_predict_proba_3
elif validation_results.at[0,'model_name'] == name_4:
    y_train_predict_proba = y_train_predict_proba_4
    y_test_predict_proba = y_test_predict_proba_4
elif validation_results.at[0,'model_name'] == name_5:
    y_train_predict_proba = y_train_predict_proba_5
    y_test_predict_proba = y_test_predict_proba_5
elif validation_results.at[0,'model_name'] == name_6:
    y_train_predict_proba = y_train_predict_proba_6
    y_test_predict_proba = y_test_predict_proba_6
elif validation_results.at[0,'model_name'] == name_7:
    y_train_predict_proba = y_train_predict_proba_7
    y_test_predict_proba = y_test_predict_proba_7
elif validation_results.at[0,'model_name'] == name_8:
    y_train_predict_proba = y_train_predict_proba_8
    y_test_predict_proba = y_test_predict_proba_8
elif validation_results.at[0,'model_name'] == name_9:
    y_train_predict_proba = y_train_predict_proba_9
    y_test_predict_proba = y_test_predict_proba_9
elif validation_results.at[0,'model_name'] == name_10:
    y_train_predict_proba = y_train_predict_proba_10
    y_test_predict_proba = y_test_predict_proba_10
elif validation_results.at[0,'model_name'] == name_11:
    y_train_predict_proba = y_train_predict_proba_11
    y_test_predict_proba = y_test_predict_proba_11
elif validation_results.at[0,'model_name'] == name_12:
    y_train_predict_proba = y_train_predict_proba_12
    y_test_predict_proba = y_test_predict_proba_12
elif validation_results.at[0,'model_name'] == name_13:
    y_train_predict_proba = y_train_predict_proba_13
    y_test_predict_proba = y_test_predict_proba_13
elif validation_results.at[0,'model_name'] == name_14:
    y_train_predict_proba = y_train_predict_proba_14
    y_test_predict_proba = y_test_predict_proba_14
elif validation_results.at[0,'model_name'] == name_15:
    y_train_predict_proba = y_train_predict_proba_15
    y_test_predict_proba = y_test_predict_proba_15
elif validation_results.at[0,'model_name'] == name_16:
    y_train_predict_proba = y_train_predict_proba_16
    y_test_predict_proba = y_test_predict_proba_16
elif validation_results.at[0,'model_name'] == name_17:
    y_train_predict_proba = y_train_predict_proba_17
    y_test_predict_proba = y_test_predict_proba_17
elif validation_results.at[0,'model_name'] == name_18:
    y_train_predict_proba = y_train_predict_proba_18
    y_test_predict_proba = y_test_predict_proba_18
elif validation_results.at[0,'model_name'] == name_19:
    y_train_predict_proba = y_train_predict_proba_19
    y_test_predict_proba = y_test_predict_proba_19
elif validation_results.at[0,'model_name'] == name_20:
    y_train_predict_proba = y_train_predict_proba_20
    y_test_predict_proba = y_test_predict_proba_20
elif validation_results.at[0,'model_name'] == name_21:
    y_train_predict_proba = y_train_predict_proba_21
    y_test_predict_proba = y_test_predict_proba_21
elif validation_results.at[0,'model_name'] == name_22:
    y_train_predict_proba = y_train_predict_proba_22
    y_test_predict_proba = y_test_predict_proba_22
    
## Вывод скоров по модели
    
#append_df_to_excel(excelFileNameB,pd.DataFrame(y_train_predict_proba),'train_scores', 0, True)
#append_df_to_excel(excelFileNameB,pd.DataFrame(y_test_predict_proba),'test_scores', 0, True)
    
## Пересчет скоров при undersampling
    
if undersampling_rate > 0:
    undersampling_rate = 1/undersampling_rate
    y_train_predict_proba = (y_train_predict_proba * undersampling_rate) / (y_train_predict_proba * undersampling_rate - y_train_predict_proba + 1)
    y_test_predict_proba  = (y_test_predict_proba * undersampling_rate) / (y_test_predict_proba * undersampling_rate - y_test_predict_proba + 1)
    
## Вывод скоров по модели после undersampling
    
#append_df_to_excel(excelFileNameB,pd.DataFrame(y_train_predict_proba),'train_scores_undrsmpl', 0, True)
#append_df_to_excel(excelFileNameB,pd.DataFrame(y_test_predict_proba),'test_scores_undrsmpl', 0, True)
    
## Составление датафрейма с коэффициентами модели
    
all_coef_list = ['INTERCEPT']
for X_train_attr in model_list_X_train:
    all_coef_list.append(X_train_attr)    
all_coef_df = pd.DataFrame(all_coef_list, columns = ['ATTR_TRUE_NAME'])
all_coef_df['COEF'] = 0.0
all_coef_df.at[0,'COEF'] = model_intercept[0]
X_train_coef_num = 0
for X_train_coef in model_coef[0]:
    X_train_coef_num += 1
    all_coef_df.at[X_train_coef_num,'COEF'] = X_train_coef
    
## Подготовка лучшей модели для проверок валидации
    
y_train = pd.Series(y_train).reset_index(drop=True)
y_train_predict_proba = pd.Series(y_train_predict_proba)
y_test = pd.Series(y_test).reset_index(drop=True)
y_test_predict_proba = pd.Series(y_test_predict_proba)
    
train_model_results = pd.concat([y_train,y_train_predict_proba], axis = 1)
train_model_results.columns = ['EVENT','SCORE'] 
test_model_results = pd.concat([y_test,y_test_predict_proba], axis = 1)
test_model_results.columns = ['EVENT','SCORE']
    
# Графики roc auc в консоли

fpr_train, tpr_train, threshold_train = metrics.roc_curve(y_train, y_train_predict_proba)
roc_auc_train = metrics.auc(fpr_train, tpr_train)
gini_train = round(2 * roc_auc_train - 1,4)
fpr_test, tpr_test, threshold_test = metrics.roc_curve(y_test, y_test_predict_proba)
roc_auc_test = metrics.auc(fpr_test, tpr_test)
gini_test = round(2 * roc_auc_test - 1,4)

plt.title('Receiver Operating Characteristic')
plt.plot(fpr_train, tpr_train, '--b', label = 'GINI_train = %0.4f' % gini_train)
plt.plot(fpr_test, tpr_test, '--r', label = 'GINI_test = %0.4f' % gini_test)
plt.legend(loc = 'lower right')
plt.plot([0, 1], [0, 1],'--k')
plt.xlim([0, 1])
plt.ylim([0, 1])
plt.ylabel('True Positive Rate')
plt.xlabel('False Positive Rate')
plt.show()
print(t)

print('10. Модель провалидирована: ' + str(dtime.datetime.now()))
print(t)

###############################################################################

### Часть 11. Вывод данных в эксель

## Построение графиков roc auc, predict, trend в основном отчете
    
wb = openpyxl.load_workbook(excelFileName)
corrmtrx_to_excel(corrTrainMtrx, wb)
roc_auc(train_model_results, test_model_results, wb, ROCAUC_n_bucket)
predict(train_model_results, test_model_results, wb, PREDICT_n_bucket)
trend(binningInfoLongFin, df_test_initial, all_coef_df, wb, short_list_attrs, list_size='short')
wb.save(excelFileName)

## Построение различных статистик во вспомогательный отчет
  
append_df_to_excel(excelFileNameS,long_list_agr_attr,'long_list_agr_attr', 0, True)
append_df_to_excel(excelFileNameS,long_list_agr_params,'long_list_agr_params', 0, True) 
append_df_to_excel(excelFileNameS,validation_results,'validation_results', 0, True)
append_df_to_excel(excelFileNameS,all_coef_df,'model_coefficients', 0, True)
append_df_to_excel(excelFileNameS,stability_df,'model_stability', 0, True)

if TREND_long == True:
    wb = openpyxl.load_workbook(excelFileNameS)
    trend(binningInfoLongFin, df_test_initial, wb, short_list_attrs, list_size='long')
    wb.save(excelFileNameS)  

print('11. Выгружены данные в EXCEL: ' + str(dtime.datetime.now()))
print(t)

###############################################################################
