import pandas as pd
import numpy as np
from numpy import linalg as LA


class SearchNameObj():
    '''
    Return the VLOOKUP range of the selected name
    '''
    def __init__(self, filename, sheet_name):
        self.fn = filename
        self.sn = sheet_name
        self.search_name = None
        self.df = pd.read_excel(self.fn, sheet_name = self.sn)
        
        self.sample = np.array([28265, 20385, 4350])
    
    def correctName(self, correctfn, correctsht):
        df_corr = pd.read_excel(correctfn, sheet_name = correctsht)

        for i in df_corr.index:
            wrongname_d = self.df[self.df['姓名 '] == df_corr.loc[i, '更正前']]
            for wn_idx in wrongname_d.index:
                self.df.loc[wn_idx, '姓名 '] = df_corr.loc[i, '更正後']
        self.df.to_excel(self.fn, sheet_name = self.sn, index = False)
    
    def getDefaultRangeStr(self):
        col_start = str(chr(65 + list(self.df.columns).index('姓名 ')))
        col_end = str(chr(65 + len(self.df.columns) - 1))
        row_start = str(2)
        row_end = str(self.df.shape[0] + 1)
        rng_str = '${}${}:${}${}'.format(col_start, row_start, col_end, row_end)
        
        return rng_str
    
    def getSearchColumnIndices(self):
        cidx_name = list(self.df.columns).index('姓名 ')
        cidx_pm1 = list(self.df.columns).index('新-本俸')
        cidx_pm2 = list(self.df.columns).index('新-專業')
        cidx_pm3 = list(self.df.columns).index('新-主管')
        return (cidx_pm1 - cidx_name + 1, 
                cidx_pm2 - cidx_name + 1, 
                cidx_pm3 - cidx_name + 1)
        
    
    def getDupNameRangeStr(self, name, payarray):
        df_names = self.df[self.df['姓名 '] == name]
        paynorm = np.array(payarray)/self.sample
        if df_names.shape[0] <= 1:
            # Duplicated names don't exist ==> no return
            rng_str = None
        else:
            # Duplicated names exist ==> return the string of single line range
            line_idx = df_names.index[0]
            right_num = 0
            diff = None
            for idx in df_names.index:
                _right_num  = 0
                _pay_array = self.df.loc[idx, ['新-本俸', '新-專業', '新-主管']].values
                _pay_norm = _pay_array/self.sample
                for i in range(len(_pay_array)):
                    if _pay_array[i] == payarray[i]:
                        _right_num += 1
                _diff = LA.norm(_pay_norm - paynorm)
                
                if _right_num > right_num:
                    line_idx = idx
                    right_num = _right_num
                    diff = _diff
                elif (_right_num == right_num) and \
                     ((diff == None) or (_diff <= diff)):
                    line_idx = idx
                    right_num = _right_num
                    diff = _diff
                    
            col_start = str(chr(65 + list(self.df.columns).index('姓名 ')))
            col_end = str(chr(65 + len(self.df.columns) - 1))
            row_start = str(line_idx + 2)
            row_end = row_start
            
            rng_str = '${}${}:${}${}'.format(col_start, row_start, col_end, row_end)
        
        return rng_str
    
    def quit(self):
        self.df.to_excel(self.fn, sheet_name = self.sn, index = False)