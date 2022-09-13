import unittest
from eu.qiou.xlDicts.xlDicts import xlDicts

class MyTestCase(unittest.TestCase):
    def test_something(self):
        d = xlDicts("123.xlsx")
        d.load("1",1,[3, 2], reversed= False)
        print(d.d)
        d.unload("123", endRow=3)


if __name__ == '__main__':
    unittest.main()
