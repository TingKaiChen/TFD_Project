import xlwings as xw


class SearchNameObj():
    '''
    Return the ID of the selected name and unit
    '''
    def __init__(self, filename = None, sheet_name = None):
        self.fn = filename
        self.sn = sheet_name
        self.search_name = None
        self.unit = None
    def setFileName(self, filename):
        self.fn = filename
    def setSheetName(self, sheetname):
        self.sn = sheetname
    def execute(self):
        self.app = xw.App(add_book = False, visible = False)
        self.app.display_alerts = False
        self.app.books.api.Open(self.fn, UpdateLinks=False)
        self.wb = self.app.books[-1]
        self.sht = self.wb.sheets[self.sn]
        self.rng = self.sht.range('A1').current_region
        self.cidx_name = self.rng.rows[0].value.index('姓名 ')
        self.cidx_unit = self.rng.rows[0].value.index('實際服務單位')
        self.cidx_id = self.rng.rows[0].value.index('身分證號')
        self.search_rng = self.rng.columns[self.cidx_name]
    def correctName(self, correctfn, correctsht):
        self.app.books.api.Open(correctfn, UpdateLinks=False)
        wb_name = self.app.books[-1]
        sht_name = wb_name.sheets[correctsht]
        rng_name = sht_name.range('A1').current_region
        for i in range(1, rng_name.shape[0]):
            wrong_name = rng_name[i, 0].value
            right_name = rng_name[i, 1].value
            replace_cell = self.rng.api.Find(wrong_name)
            if replace_cell != None:
                replace_cell.value = right_name
        self.wb.save()
        wb_name.save()
        wb_name.close()
    def findID(self, search_name, unit):
        search_fml = '=COUNTIF(' + self.search_rng.address + ', "' + search_name + '")'
        rep_cell = self.sht.range((1, len(self.rng.columns) + 1))
        rep_cell.formula = search_fml
        rep_num = int(rep_cell.value)
        # Find the correct ID of the name
        name_cell = self.search_rng.api.Find(search_name)
        if name_cell == None:
            rep_cell.clear()
            return None
        unit_cell = self.rng[name_cell.row - 1, self.cidx_unit]
        for i in range(1, rep_num):
            # Danger: use '==' instead
            if unit not in unit_cell.value:
                name_cell = self.search_rng.api.FindNext(name_cell)
                unit_cell = self.rng[name_cell.row - 1, self.cidx_unit]
            else:
                break
        rep_cell.clear()
        return self.rng[name_cell.row - 1, self.cidx_id].value
    def quit(self):
        self.wb.save()
        self.wb.close()
        for wb in self.app.books:
            wb.save()
            wb.close()
        self.app.quit()
        self.app.kill()
