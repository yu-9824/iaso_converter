import PySimpleGUI as sg
import json
import os


APP_NAME = 'IASO converter'
path_settings = os.path.join(os.path.dirname(__file__), 'settings.json')


# フォント
font_default = 'Hiragino Sans CNS'
font_size_default = 20
option_text_default = {
    'font': (font_default, font_size_default),
}

option_popup = dict(**option_text_default, modal = False, keep_on_top=True)


class gui:
    def __init__(self):
        '''
        theme: 'SystemDefault' or 'Black'
        '''
        # load setting
        try:
            self.settings = json.load(open(path_settings, mode = 'r', encoding = 'utf_8'))
        except Exception as e:
            pass
        else:
            if self.settings['theme'] == 'light':
                theme = 'LightGray1'
            elif self.settings['theme'] == 'dark':
                theme = 'Black'
            sg.theme(theme)

        lang = self.settings['lang']

    def run(self):
        main_menu = Menu(layout=[
            [sg.Text('1. Select input file'), sg.InputText(visible=False, key='-INPUT-FILE-PATH-'), sg.FileBrowse()],
            [sg.Text('2. Select output folder'), sg.InputText(visible=False, key='-OUTPUT-FILE-PATH-'), sg.FileSaveAs('SaveAs', default_extension = '.xlsx', file_types=(('Excel file', '.xlsx'),))],
            [sg.Text('_'*50)],
            [sg.Submit('Run')],
        ])

        main_menu.make_window()

        while True:
            main_menu.read()
            # print(main_menu.values)
            if main_menu.event is None:
                break
            elif main_menu.event == 'Run':
                fpath_input = main_menu.values['-INPUT-FILE-PATH-']
                fpath_output = main_menu.values['-OUTPUT-FILE-PATH-']

                if fpath_input == '':
                    sg.PopupError('You have to select input file.', **option_popup)
                    continue
                elif fpath_output == '':
                    sg.PopupError('You have to select output directory.', **option_popup)
                    continue
                else:
                    try:
                        cv = Converter(fpath_input=fpath_input, fpath_output=fpath_output)
                        cv.main()
                    except Exception as e:
                        sg.PopupError(e, **option_popup)
                        continue
                    else:
                        ok_or_cancel = sg.PopupYesNo('Completed!\nWould you like to convert another one?', **option_popup)
                        if ok_or_cancel in ('Yes', None):
                            continue
                        elif ok_or_cancel == 'No':
                            break




class Menu:
    def __init__(self, layout = ((sg.Text('You have to add something to show.')),)):
        self.layout = layout
        self.menu_def = [
            ['Menu', ['About {}'.format(APP_NAME), '---', 'Setting']],
        ]

    def make_window(self, **options):
        default_options = {
            'size' : (800, 450),
            'element_justification' : 'center',
            'resizable': True
        }
        default_options.update(option_text_default)
        default_options.update(**options)
        # Add MenuBar
        self.layout.insert(0, [sg.Menu(self.menu_def, font = sg.DEFAULT_FONT)])
        self.window = sg.Window(APP_NAME, layout = self.layout, **default_options, finalize = True)

    def read(self):
        self.event, self.values = self.window.read()
        if self.event == 'Setting':
            if _change_setting():   # 設定に反映させるために閉じる場合
                self.window.close()
        elif self.event == 'About {}'.format(APP_NAME):
            with open('about.txt', mode = 'r', encoding = 'utf_8') as f:
                lcns = f.read()
            sg.PopupOK(lcns, modal= False, keep_on_top= True, title = 'About {}'.format(APP_NAME))

def _change_setting():
    settings = json.load(open(path_settings, mode = 'r', encoding = 'utf_8'))

    corr_lang = {
        '日本語': 'ja',
        'English': 'en',
        'ja': '日本語',
        'en': 'English'
    }
    corr_theme = {
        'Light': 'light',
        'Dark': 'dark',
        'light': 'Light',
        'dark': 'Dark',
    }

    setting_menu = Menu(layout = [
        [sg.Text('Setting')],
        [sg.Text('')],
        # [sg.Text('Langage'), sg.Combo(['日本語', 'English'], default_value = corr_lang[settings['lang']], key = 'lang')],
        [sg.Text('Theme'), sg.Combo(['Light', 'Dark'], default_value = corr_theme[settings['theme']], key = 'theme')],
        [sg.Text('')],
        [sg.Cancel(), sg.OK()]
    ])

    setting_menu.make_window(size = (None, None))
    while True:
        setting_menu.read()
        do_close = False    # 再起動するかどうか
        if setting_menu.event == 'OK':
            settings = {
                # 'lang': corr_lang[setting_menu.values['lang']],
                'lang': 'ja',
                'theme': corr_theme[setting_menu.values['theme']]
                }
            event = sg.PopupYesNo('You will need to reboot to apply the configuration changes.\nCan I close it to apply the settings?', modal = False, keep_on_top = True, **option_text_default)
            if event == 'Yes':
                json.dump(settings, open(path_settings, mode = 'w', encoding='utf_8'), indent = 4)
                do_close = True
                break
        else:
            break
    setting_menu.window.close()
    return do_close

import pandas as pd
import openpyxl as xl
import codecs as cd
import numpy as np
from openpyxl.styles.borders import Border, Side
import os


rstrip = np.frompyfunc(lambda x:x.rstrip('\r\n'), 1, 1)

class Converter:
    def __init__(self, fpath_input, fpath_output, delimiter = ','):
        self.fpath_input = fpath_input
        self.fpath_output = fpath_output
        self.delimiter = delimiter
        
        fpath_cols = 'cols.txt'
        with open(fpath_cols, mode='r', encoding='utf_8_sig') as f:
            self.cols = rstrip(f.readlines())
        # self.cols = ['薬品名','メーカー','規格','CAS No.','内容量','単位名','未開封','開封','法規 日本語','シンボル 日本語','形状','純度規格']

        self.options_codecs = ["r", 'shift_jis', "ignore"]

    def main(self):
        skiprows = self._get_skiprows()
        df = self._get_df(skiprows=skiprows)
        df_info = self._get_df_info(skiprows=skiprows)
        df_info_add = self._extract_info(df_info=df_info)
        self._output_excel(df=df, df_info_add=df_info_add)
        # レイアウトを整える．
        self._arrange_layout(df_info_add=df_info_add)
        

    def _get_skiprows(self):
        with cd.open(self.fpath_input, *self.options_codecs) as csv_file:
            row = ['']
            skiprows = -3
            while not (set(self.cols) <= set(row)):
                row = csv_file.readline().split(self.delimiter)
                skiprows += 1
        return skiprows

    def _get_df(self, skiprows):
        with cd.open(self.fpath_input, *self.options_codecs) as csv_file:
            df = pd.read_table(csv_file, sep=',', header=skiprows)
        return df



    def _get_df_info(self, skiprows):
        usecols = [0, 1]
        with cd.open(self.fpath_input, *self.options_codecs) as csv_file:
            df_info = pd.read_table(csv_file, sep=self.delimiter, nrows=skiprows-1, usecols=usecols, header=None)
        return df_info


    def _extract_info(self, df_info):
        n_userows = 2
        df_info_add = rstrip(df_info[-n_userows:])
        df_info_add.columns = range(df_info_add.shape[1])
        df_info_add = pd.concat([df_info_add, pd.DataFrame(np.array([np.nan]*df_info_add.shape[1])).transpose()], axis=0, ignore_index=True)    # 1行あける
        return df_info_add


    def _output_excel(self, df, df_info_add):
        df_output = pd.concat([df_info_add, df[self.cols].transpose().reset_index().transpose()], axis = 0, ignore_index=True)

        # 書き出す．
        df_output.to_excel(self.fpath_output, index=False, header=False)


    def _arrange_layout(self, df_info_add):
        # read input xlsx
        wb1 = xl.load_workbook(filename=self.fpath_output)
        ws1 = wb1.worksheets[0]

        # set border (black thin line)
        side = Side(style='thin', color='000000')
        border = Border(top=side, bottom=side)

        # set column width
        for col in ws1.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            adjusted_width = (max_length + 2) * 1.2
            # print(ws1.column_dimensions[column].width)
            ws1.column_dimensions[column].width = adjusted_width

        # set height
        for row in ws1.rows:
            n_lines = 1
            r = row[0].row
            for cell in row:
                if cell.value is not None:
                    n_lines = max(n_lines, str(cell.value).count('\n')+1) # intになる可能性があるのでstrを使ってる．
                if cell.row > df_info_add.shape[0]:
                    ws1[cell.coordinate].border = border
            adjusted_height = 11 * 1.25 * n_lines   # 11はフォントサイズ．
            ws1.row_dimensions[r].height = adjusted_height
        # save xlsx file
        wb1.save(self.fpath_output)


if __name__ == '__main__':
    gui().run()

    # fpath_input = 'input/stock_list_20210610145730.csv'

    # dir_output = os.path.dirname(__file__)
    # fname_output = '{}.xlsx'.format(os.path.splitext(os.path.basename(fpath_input))[0])
    # fpath_output = os.path.join(dir_output, fname_output)

    # cv = Converter(fpath_input=fpath_input, fpath_output=fpath_output)
    # cv.main()
    

