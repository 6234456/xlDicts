import unittest
from eu.qiou.xlDicts.xlDicts import xlDicts

class MyTestCase(unittest.TestCase):
    def test_something(self):
        d = xlDicts("123.xlsx")
        d.load("1",1,[3, 2], reversed= False, asFormula=True).unload("123")
        print(d.data)
        d.loadStruct("mapping").feed()
        print(d.struct)
        print(d.structuredData)
        d.dumpStructuredData("123", 1, 20)
        d.aggregate(lambda dic: "=" +("+".join([v[0] for v in dic.values()]))).dump("123", 20, 20).save()
        # d.aggregate(lambda dic: sum(([v[0] for v in dic.values()]))).dump("123", 20, 20).save()


if __name__ == '__main__':
    unittest.main()
