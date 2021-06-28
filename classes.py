import time
import datetime

from xlrd.formula import num2strg

# 时间戳类
class Timestamp():
    def __init__(self, date: str) -> None:
        strs = date.split('/')
        self.year = int(strs[0])
        self.month = int(strs[1])
        self.day = int(strs[2])

    def __add__(self, date: str):
        strs = date.split('/')
        year = int(strs[0])
        month = int(strs[1])
        day = int(strs[2])
        self.year = year + self.year
        if month + self.month > 12:
            self.year += 1
            self.month = (month + self.month) % 12
        else:
            self.month += month
        # 可能出现常识错误，如2月30日，不影响计算
        self.day += day
        return Timestamp(self.__str__())

    def __lt__(self, rhs: object) -> bool:  # 小于
        if isinstance(rhs, Timestamp):
            if self.year < rhs.year:
                return True
            elif self.year > rhs.year:
                return False
            if self.month < rhs.month:
                return True
            elif self.month > rhs.month:
                return False
            if self.day < rhs.day:
                return True
            else:
                return False
        else:
            raise Exception("The type of object must be 'Timestamp'!")

    def __eq__(self, rhs: object) -> bool:  # 等于
        if isinstance(rhs, Timestamp):
            return self.year == rhs.year and self.month == rhs.month and self.day == rhs.day
        else:
            raise Exception("The type of object must be 'Timestamp'!")    

    def __str__(self) -> str:               # 输出
        return str(self.year) + '/' + str(self.month) + '/' + str(self.day)

    __repr__ = __str__

# 浮动利率计算类
class FloatRateEx():
    def __init__(self) -> None:
        # 存储格式：(Timestamp, float)，代表（日期，利率）
        self.oneYear = []
        self.SHIBOR1Year = []
        self.SHIBOR3Month = []

    # 获得真实票面利率
    def GetRate(self, date: Timestamp, type: str) -> float:
        if type == '一年定存利率':
            for item in self.oneYear:
                if item[0] < date:
                    return item[1]
        elif type == '票面利率以3个月Shibor5日均值' :
            for i in range(len(self.SHIBOR3Month)):
                if self.SHIBOR3Month[i][0] < date:
                    rateSum = 0
                    for i in range(i, i + 5, 1):
                        rateSum += self.SHIBOR3Month[i][1]
                    return rateSum / 5
        elif type == '票面利率以1年期Shibor20日均值':
            for i in range(len(self.SHIBOR1Year)):
                if self.SHIBOR1Year[i][0] < date:
                    rateSum = 0
                    for i in range(i, i + 20, 1):
                        rateSum += self.SHIBOR1Year[i][1]
                    return rateSum / 20
        else:
            raise Exception("Unknown type!")

    # 根据输入的估值日期获得估计利率，返回列表
    def GetEstimateRate(self, testDate: Timestamp, timeStamps, type: str) -> list:
        # 找到夹着testDate（估值日期）的两个现金流发生日
        rates = []
        leftBound, rightBound = None, None
        for i in range(len(timeStamps)):
            if testDate < timeStamps[i]:
                rightBound = timeStamps[i]
                break
            rates.append(0) # 填充空位而已
        testDateRate = self.GetRate(testDate, type)
        if i != 0:
            leftBound = timeStamps[i - 1]
            rates.append(self.GetRate(leftBound, type))
            for i in range(len(timeStamps) - i - 1):
                rates.append(testDateRate)
        # 如果估值日期比第一个现金流发生日还早
        else:
            rates = [testDateRate for i in range(len(timeStamps))]
        return rates

    def Add(self, type: str, time: Timestamp, rate: float) -> None:
        if type == 'oneYear':
            self.oneYear.append((time, rate))
        elif type == 'SHIBOR3Month':
            self.SHIBOR3Month.append((time, rate))
        elif type == 'SHIBOR1Year':
            self.SHIBOR1Year.append((time, rate))
        else:
            raise Exception("Unknown type!")

# 国债收益率计算类
class RiskfreeRate():
    def __init__(self) -> None:
        # 格式：(float, float)，代表（期限，利率）
        self.rate = []

    # 找到第一个小于所给日期的期限，比如输入期限为1.5，就返回1年期国债收益率
    def GetRate(self, maturity):
        for i in range(len(self.rate)):
            if self.rate[i][0] > maturity:
                return self.rate[i - 1][1]

    def Add(self, maturity, rate):
        self.rate.append((maturity, rate))

# 债券类
class Secutiry():
    def __init__(
        self, 
        code, # 债券代码
        name, # 债券名称
        couponType, # 息票类型
        couponRate, # 票面利率
        rateInfo: str, # 利率说明
        rateType, # 利率类型
        frequency, # 付息频率
        carrayDate, # 起息日期
        maturityDate,  # 到期日期
        category, # 类别（风险因子）
        maturity, # 到期期限
        closingPrice, # 收盘全价
    ) -> None:
        self.code = code # 债券代码
        self.name = name # 债券名称
        self.couponType = couponType # 息票类型
        self.couponRate = couponRate # 票面利率
        self.rateInfo = rateInfo # 利率说明
        self.rateType = rateType # 利率类型
        self.frequency = frequency # 付息频率
        self.carrayDate = carrayDate # 起息日期
        self.maturityDate = maturityDate # 到期日期
        self.category = category
        self.maturity = maturity
        self.closingPrice = closingPrice

        # 分割时间点
        self.dateArray = []
        self.SplitHoldingTime()
        self.rateArray = []

    # 得到计算现金流的开始时间
    def GetNearestDateIndex(self, date: Timestamp) -> int:
        for i in range(len(self.dateArray)):
            if date < self.dateArray[i]:
                return i
        return -1

    # 把起息日期到到期日期分割为一个个需要付息的时间点
    def SplitHoldingTime(self) -> None:
        delta = '0/0/0'
        if self.frequency == -1:    # 到期一次还本付息
            self.dateArray.append(self.carrayDate)
            self.dateArray.append(self.maturityDate)
            return
        elif self.frequency == 1:
            delta = '1/0/0'
        else:
            deltaMonth = int(12 / self.frequency)
            delta = '0/' + str(deltaMonth) + '/0'
        startDate = Timestamp(self.carrayDate.__str__())
        while startDate < self.maturityDate:
            self.dateArray.append(startDate)
            startDate += delta
        # self.carrayDate += '-1/0/0'
    
    # 计算出每期现金流对应利率
    def GetRateArray(self, floatRate: FloatRateEx):
        if self.rateType == u'固定利率':
            for i in range(len(self.dateArray)):
                self.rateArray.append(self.couponRate)
        elif self.rateType == u'累进利率':
            info = self.rateInfo.split(';')
            # 此段有较大局限性，仅能支持两段式累进利率
            period, rateInfo = info[0].split(',')
            (beg, end) = tuple(map(int, period.split('-')))
            length = int((end - beg + 1) / 10000)
            rate = float(rateInfo.split(':')[1].split('%')[0])
            for i in range(length - 1):
                self.rateArray.append(rate / 100)

            period, rateInfo = info[1].split(',')
            (beg, end) = tuple(map(int, period.split('-')))
            length = int((end - beg + 1) / 10000)
            rate = float(rateInfo.split(':')[1].split('%')[0])
            for i in range(length + 1):
                self.rateArray.append(rate / 100)
            # for i in info:
            #     period, rateInfo = i.split(',')
            #     (beg, end) = tuple(map(int, period.split('-')))
            #     length = int((end - beg + 1) / 10000) + 1
            #     rate = float(rateInfo.split(':')[1].split('%')[0])
            #     for i in range(length):
            #         self.rateArray.append(rate / 100)
            # self.rateArray = self.rateArray[1:]
        elif self.rateType == u'浮动利率':
            info = self.rateInfo.split('+')
            _type = info[0]
            _rate = float(info[1].split('%')[0])
            tmp_dateArray = [self.carrayDate]
            tmp_dateArray.extend(self.dateArray[:len(self.dateArray) - 1])
            for date in tmp_dateArray:
                baseRate = floatRate.GetRate(date, _type)
                self.rateArray.append((baseRate + _rate) / 100)
        else:
            raise Exception('Unsupported rate type!')

    def EstimateRateArray(self, testDate: Timestamp, floatRate: FloatRateEx):
        self.rateArray = []
        if self.rateType == u'固定利率':
            for i in range(len(self.dateArray)):
                self.rateArray.append(self.couponRate)
        elif self.rateType == u'累进利率':
            info = self.rateInfo.split(';')
            # 此段有较大局限性，仅能支持两段式累进利率
            period, rateInfo = info[0].split(',')
            (beg, end) = tuple(map(int, period.split('-')))
            length = int((end - beg + 1) / 10000)
            rate = float(rateInfo.split(':')[1].split('%')[0])
            for i in range(length - 1):
                self.rateArray.append(rate / 100)

            period, rateInfo = info[1].split(',')
            (beg, end) = tuple(map(int, period.split('-')))
            length = int((end - beg + 1) / 10000)
            rate = float(rateInfo.split(':')[1].split('%')[0])
            for i in range(length + 1):
                self.rateArray.append(rate / 100)
        elif self.rateType == u'浮动利率':
            info = self.rateInfo.split('+')
            _type = info[0]
            _rate = float(info[1].split('%')[0])    # 利率增量
            rates = floatRate.GetEstimateRate(testDate, self.dateArray, _type)
            for rate in rates:
                self.rateArray.append((rate + _rate) / 100)
        else:
            raise Exception('Unsupported rate type!')

    def CalcDateGap(self, date: str) -> float:
        date1 = date
        date2 = self.maturityDate.__str__()
        date1=time.strptime(date1,"%Y/%m/%d")
        date2=time.strptime(date2,"%Y/%m/%d")
        date1=datetime.datetime(date1[0],date1[1],date1[2])
        date2=datetime.datetime(date2[0],date2[1],date2[2])
        #返回两个变量相差的值，就是相差天数
        return (date2 - date1).days / 365

    def UpdateMaturity(self, date):
        self.maturity = self.CalcDateGap(date)

    def CalcCashFlow(self, date: Timestamp):
        
        timeStamps = []
        cashFlow = []
        beg = self.GetNearestDateIndex(date)
        if beg == -1:
            return None, None
        # 根据证券类型处理
        if self.couponType == u'到期一次还本付息' or self.couponType == u'贴现':
            if beg == 1:
                cashFlow.append(100)
                timeStamps.append(self.dateArray[beg])
            elif beg == 0:
                cashFlow.append(100 * self.couponRate + 100)
                timeStamps.append(self.dateArray[beg + 1])
        elif self.couponType == u'附息':
            for i in range(beg, len(self.dateArray) - 1):
                actualRate = self.rateArray[i] / self.frequency
                timeStamps.append(self.dateArray[i])
                cashFlow.append(actualRate * 100)
            timeStamps.append(self.dateArray[-1])
            cashFlow.append(100 + 100 * self.rateArray[len(self.dateArray) - 1] / self.frequency)
        else:
            raise Exception('unknown coupon type!')
        return timeStamps, cashFlow