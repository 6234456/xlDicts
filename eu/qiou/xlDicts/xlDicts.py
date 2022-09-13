from openpyxl import load_workbook
from openpyxl.worksheet import worksheet

class xlDicts:
    def __init__(self, wb: str = "", data_only: bool = True):
        self.file = wb
        self.wb = load_workbook(wb,  data_only=data_only)

    def __getSht(self, sht: str = "") -> worksheet:
        if len(sht) > 0:
            return self.wb[sht]
        else:
            return self.wb.active

    def load(self, sht: str = "", keyCol: int = 1, valCol = 2, startRow : int = 1, endRow: int = None, ignoreNullVal: bool = True, setNullValTo = 0, reversed:bool = False):
        """
                   load the data on the spreadsheet to dict
                   @param
                           sht             sht name, active sheet by default
                           keyCol          can be a single value           -> the normal Dicts instance
                           valCol          can be a single value           -> the normal Dicts instance
                                           or tuple with two elements      -> (keyColFrom, keyColTo)
                                           or list with more than one,     -> [keyLvl1, keyLvl2, ...]
                                           which specified the order of the columns
                           startRow        from which row to start
                           endRow          ends at which row
                           ignoreNullVal   whether to ignore the key if the corresponding value is None, default True
                           setNullValTo    set the None value to the given value, valid only when ignoreNullVal is False
                           reversed        read from top to bottom. the duplicated value will be replaced by the one below
               """


        s = self.__getSht(sht)
        end = endRow or s.max_row

        if not type(valCol) is list:
            if type(valCol) is tuple:
                valCol = range(valCol[0], valCol[1] + 1)
            elif isinstance(valCol, int):
                valCol = [valCol]
            else:
                raise TypeError("the type of {0} is invalid!\n Integer, Tuple with two elements or List is required.")

        filterCol = [i - min(valCol) for i in valCol]
        k = list(list(s.iter_cols(keyCol, keyCol, startRow, end, True))[0])
        v = [[i[j] for j in filterCol] for i in s.iter_rows(startRow, endRow, min(valCol), max(valCol), True)]

        if reversed:
            k.reverse()
            v.reverse()

        self.d = {k: v for k, v in dict(zip(k, v)).items() if not k == None}

        if ignoreNullVal:
            self.d = {k: [i or setNullValTo for i in v] for k, v in self.d.items() if any(v)}

        if len(valCol) == 1:
            self.d = {k: v[0] for k, v in self.d.items()}

    def fromDict(self, d:dict):
        self.d = d

    def unload(self, sht: str = "", keyCol: int = 1, startCol = 2, startRow : int = 1, endRow: int = None):
        if self.d and len(self.d):
            s = self.__getSht(sht)
            l = list(list(self.d.values())[0])
            end = endRow or s.max_row
            ks = list(self.d.keys())
            for r in range(startRow, end + 1):
                k = s.cell(column=keyCol, row=r).value
                if k in ks:
                    tmp = self.d[k]
                    cnt = 0
                    for c in range(startCol, startCol + len(l)):
                        s.cell(column=c, row=r, value=tmp[cnt])
                        cnt = cnt + 1

            self.wb.save(self.file)








