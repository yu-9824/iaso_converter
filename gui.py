from src import Converter
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
        finally:
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

if __name__ == '__main__':
    gui().run()