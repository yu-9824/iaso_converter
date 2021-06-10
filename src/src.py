import pandas as pd
import openpyxl as xl
import codecs as cd
import numpy as np
from openpyxl.styles.borders import Border, Side
import os


rstrip = np.frompyfunc(lambda x:x.rstrip('\r\n'), 1, 1)


fpath_input = 'input/stock_list_20210610145730.csv'

dir_output = os.path.dirname(__file__)
fname_output = '{}.xlsx'.format(os.path.splitext(os.path.basename(fpath_input))[0])
fpath_output = os.path.join(dir_output, fname_output)



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
        wb1 = xl.load_workbook(filename=fpath_output)
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
    cv = Converter(fpath_input=fpath_input, fpath_output=fpath_output)
    cv.main()
    

