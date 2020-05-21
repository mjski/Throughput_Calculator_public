"""

    author: Morgan D
    module: fiscal_year_str.py
    python version: 3.7
    creation date: 24 Jan 2020
    purpose: contains class FiscalYearStr which takes an input year and produces a list of strings with all fiscal
             months in the format of 'Jan-2020'.

"""

from calendar import month_abbr


class FiscalYearStr:

    def __init__(self, year=2020):
        self.__year = year

    def get_cal(self):
        c = [month_abbr[i].title() for i in range(-3, 10)]
        del(c[3])
        c = [i + '-' for i in c]
        return c

    def get_fy(self, c):
        b = []
        for i in c:
            if i == 'Oct-' or i == 'Nov-' or i == 'Dec-':
                b.append(self.year - 1)
        d = [self.year] * 9
        [b.append(i) for i in d]
        fy = ['{}{}'.format(a, d) for a, d in zip(c, b)]
        return fy

    @property
    def year(self):
        return self.__year

    @year.setter
    def year(self, value):
        self.__year = value


def main():
    fiscal_year = FiscalYearStr(2020)
    c = fiscal_year.get_cal()
    fy_format = fiscal_year.get_fy(c)
    print(fy_format)
    return fy_format


if __name__ == '__main__':
    main()
