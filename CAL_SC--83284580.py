# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandapower as pp
import pandapower.plotting as plot
import pandapower.shortcircuit as sc

import codecs
import pandas as pd

import time
import logging
import sys
import webbrowser


def create_logger(name, silent=False, to_disk=False, log_file=None):
    # setup logger
    # print('create logger')
    log = logging.getLogger(name)
    if not log.handlers:
        log.setLevel(logging.DEBUG)
        formatter = logging.Formatter(fmt='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S')
        if not silent:
            ch = logging.StreamHandler(sys.stdout)
            ch.setLevel(logging.INFO)
            ch.setFormatter(formatter)
            log.addHandler(ch)
        if to_disk:
            log_file = log_file if log_file is not None else strftime("%Y-%m-%d-%H-%M-%S.log", gmtime())
            fh = logging.FileHandler(log_file, encoding='utf-8',  mode='w')
            fh.setLevel(logging.DEBUG)
            fh.setFormatter(formatter)
            log.addHandler(fh)
    log.propagate = False
    return log 

panda_json= 'panda.json'
panda_ansi_json = 'panda_ansi.json'
output_json='output.json'
output_excel='Result.xlsx'
output_html='Result.html'
output_html_css='Result_css.html'
scheme_html='scheme.html'
log='encalc_sc.log'
#logging.basicConfig(filename="encalc_sc.log", filemode='w', encoding='utf-8',format='%(asctime)s - %(message)s', level= logging.INFO)

logger = create_logger('ENCALC',to_disk=True,log_file= log)


logger.info('Выполнен старт программы расчета КЗ')

try:

    with codecs.open(panda_json, 'r', encoding= "utf-8") as file:
        lines = file.read()

except OSError as err:
    logger.critical('Системная ошибка: {0}'.format(err))
    #logger.error('Нет расчетного файла по адресу: ' + panda_json)  
except BaseException as err:
    logger.debug(f'Нежданчик {err=}, {type(err)=}')
    raise

    
logger.info('Файл модели успешно загружен по пути: '+ panda_json)

 
try:

    with codecs.open(panda_ansi_json, 'w', encoding = 'ansi') as file:
        file.write(lines)

except OSError as err:
    logger.critical('Системная ошибка: {0}'.format(err))
    #logger.error('Нет расчетного файла по адресу: ' + panda_ansi_json)
except BaseException as err:
    logger.debug(f'Нежданчик {err=}, {type(err)=}')
    raise


try:
    net = pp.from_json(panda_ansi_json) #Загрузка рабочей модели (и текущего состояния схемы)
except BaseException as err:
    print(f"Ошибка импорта {err=}, {type(err)=}")
    logger.critical(f"Ошибка импорта {err=}, {type(err)=}")
    raise

try:
    sc.calc_sc(net,ip=True,return_all_currents=True) #Запуск функции расчёта КЗ
except BaseException as err:
    print(f"Ошибка расчета КЗ {err=}, {type(err)=}")
    logger.error(f'Ошибка расчета КЗ {err=}, {type(err)=}')
    raise

pp.to_excel(net, output_excel) # Для возможности визуальной проверки полученных результатов 

    
array=net.res_bus_sc
bus = net.bus


frames = pd.merge(bus, array, how= "inner",right_index=True, left_index=True)

my_data= frames.iloc[:,[0,6,8]]

my_data.rename(columns = {'name': 'Название шины', 'ikss_ka': 'Ток КЗ, кА', 'ip_ka': 'Пиковый ток КЗ, кА'}, inplace = True)

pd.set_option('colheader_justify', 'center')

#html = my_data.to_html(output_html, index= False , encoding= "utf-8", classes='mystyle')

#file = codecs.open(output_html, "r", "utf-8")
#html_file = file.read()


html_string = '''
<html>
  <head><title>Рачет токов КЗ на шинах</title></head>
  <link rel="stylesheet" type="text/css" href="new.css"/>
  <body>
    {table}
  </body>
</html>.
'''


with open(output_html_css, 'w') as result_html:
    result_html.write(html_string.format(table=my_data.to_html(index= False, classes='mystyle')))
    

#result_html.close() 

#file = codecs.open("style.css", "r")
#css_file = file.read()
#print(css_file)
#window = webview.create_window('Расчет токов КЗ на шинах', html= result_html)
#window = webview.create_window('Расчет токов КЗ на шинах', html= output_html)
#window = webview.create_window('Расчет токов КЗ на шинах', html= output_html_css)


#window = webview.create_window('Расчет токов КЗ на шинах', html= html_file, text_select=True)
#plot.to_html(net,scheme_html,respect_switches=True,include_lines=True,include_trafos=True, show_tables=True)    
#pp.plotting.simple_plot(net, respect_switches=True, line_width=1.0, bus_size=1.0, trafo_size=0.5, plot_loads=True, plot_sgens=True, load_size=1.0, sgen_size=1.0, switch_size=2.0, switch_distance=1.0, plot_line_switches=True, scale_size=True, bus_color='b', line_color='grey', trafo_color='k', ext_grid_color='y', switch_color='k', library='igraph', show_plot=True, ax=None)
#webbrowser.open(scheme_html)
webbrowser.open('file://Result_css.html', new= 2)

logger.info('Конец расчета КЗ')
#logger.disabled= True


