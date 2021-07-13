from WindPy import w
from pandas.core.frame import DataFrame
import matplotlib.pyplot as plt
import time
from datetime import datetime
import dateutil

def GetEDBData(code, beginTime, endTime, usedf=True) -> DataFrame:
    error_code, cpi = w.edb(code, beginTime, endTime, "Fill=Previous", usedf=usedf)    # CPI指标
    if error_code == 0:
        return cpi
    else:
        raise Exception('edb指标%s获取失败！' % code)

# 此函数为美林周期算法的具体实现
def _MLCycle(cpi: DataFrame, oecd: DataFrame, cpi_rolling=12, oecd_rolling=1, fig=False):
    cpi = cpi.rolling(cpi_rolling).mean().rename(columns={'CLOSE':'CPI'})
    oecd = oecd.rolling(oecd_rolling).mean().rename(columns={'CLOSE':'OECD'})
    data = cpi.join(oecd)
    data.index.name = 'date'
    data['deltaCPI'] = cpi['CPI'].diff()
    data['deltaOECD'] = oecd['OECD'].diff()
    data = data.dropna(axis=0, how='any')
    # print(data)
    result = []

    for index, row in data.iterrows():
        c = row['deltaCPI']
        o = row['deltaOECD']
        if c < 0 and o > 0:
            result.append([index.__str__(), 1])
        elif c > 0 and o > 0:
            result.append([index.__str__(), 2])
        elif c > 0 and o < 0:
            result.append([index.__str__(), 3])
        elif c < 0 and o < 0:
            result.append([index.__str__(), 4])
        else:   # 出现0沿用上一个周期
            result.append([index.__str__(), result[-1][1]])

    if fig:
        fig = plt.figure(figsize=(15,5))
        ax = fig.add_subplot(111)
        data['CPI'].plot(ax=ax, grid=True, label='CPI', marker='X')
        data['OECD'].plot(ax=ax, grid=True, secondary_y=True, label='OECD', marker='*')
        plt.title('ML cycle')
        plt.show()
    return result

# 此函数为货币信用周期算法的具体实现
def _CurrencyCreditCycle(loan: DataFrame, inteRate: DataFrame, loan_rolling=1, inteRate_rolling=1, fig=False):
    loan = loan.rename(columns={'CLOSE':'Loan'})
    inteRate = inteRate.rename(columns={'CLOSE':'Rate'})
    data = loan.join(inteRate, how='outer')
    data.index.name = 'date'
    data['Rate'] = data['Rate'].fillna(method='ffill')
    data = data.dropna(axis=0, how='any')
    # for index, row in data.iterrows():
    #     print(index, row['Loan'], row['Rate'])
    data['Loan'] = data['Loan'].rolling(loan_rolling).mean()
    data['Rate'] = data['Rate'].rolling(inteRate_rolling).mean()
    data['deltaLoan'] = data['Loan'].diff()
    data['deltaRate'] = data['Rate'].diff()

    data = data.dropna(axis=0, how='any')
    result = []

    for index, row in data.iterrows():
        l = row['deltaLoan']
        r = row['deltaRate']
        if l < 0 and r < 0:
            result.append([index.__str__(), 1])
        elif l > 0 and r < 0:
            result.append([index.__str__(), 2])
        elif l > 0 and r > 0:
            result.append([index.__str__(), 3])
        elif l < 0 and r > 0:
            result.append([index.__str__(), 4])
        else:   # 出现0沿用上一个周期
            result.append([index.__str__(), result[-1][1]])
    
    if fig:
        fig = plt.figure(figsize=(15,5))
        ax = fig.add_subplot(111)
        data['Loan'].plot(ax=ax, grid=True, label='Loan', marker='X')
        data['Rate'].plot(ax=ax, grid=True, secondary_y=True, label='Rate', marker='*')
        plt.title('currency credit cycle')
        plt.show()
    return result

'''
---美林周期的接口---
oecd_code: OECD指标的指标代码，仅支持索引为“月”的指标
cpi_code: CPi指标的指标代码，仅支持索引为“月”的指标
begin_time: 计算周期的开始时间，表示方式如：'2000-01-31'
end_time: 计算周期的结束时间，表示方式同上
oecd_rolling: OECD指标的移动平均数量，单位：月
cpi_rolling: CPI指标的移动平均数量，单位：月
fig: 是否展现简单可视化，若为True弹窗会阻塞程序
'''

def MLCycle(oecd_code: str, cpi_code: str, begin_time: str, end_time: str, oecd_rolling=1, cpi_rolling=1, fig=False):
    w.start()
    ## [数据准备start]
    # 根据rolling换算开始时间
    oecd_begin_time = time.strptime(begin_time,"%Y-%m-%d")
    oecd_begin_time = datetime(oecd_begin_time[0], oecd_begin_time[1], oecd_begin_time[2]) - \
        dateutil.relativedelta.relativedelta(months=oecd_rolling)
    cpi_begin_time = time.strptime(begin_time,"%Y-%m-%d")
    cpi_begin_time = datetime(cpi_begin_time[0], cpi_begin_time[1], cpi_begin_time[2]) - \
        dateutil.relativedelta.relativedelta(months=cpi_rolling)

    cpi = GetEDBData(cpi_code, beginTime=cpi_begin_time, endTime=end_time, usedf=True)    # CPI指标
    oecd = GetEDBData(oecd_code, beginTime=oecd_begin_time, endTime=end_time, usedf=True)    # OECD指标
    ## [数据准备end]

    ## [计算start]
    return _MLCycle(cpi, oecd, cpi_rolling=cpi_rolling, oecd_rolling=oecd_rolling, fig=fig)
    ## [计算end]

'''
---货币信用周期接口---
loan_code: 贷款指标的指标代码，DataFrame的索引会以此指标的日期为准
interestRate_code: 利率指标的指标代码
begin_time: 计算周期的开始时间，表示方式如：'2000-01-31'
end_time: 计算周期的结束时间，表示方式同上
loan_rolling: OECD指标的移动平均数量，单位：月
inteRate_rolling: CPI指标的移动平均数量，单位：月
fig: 是否展现简单可视化，若为True弹窗会阻塞程序
'''

def CurrencyCreditCycle(loan_code, interestRate_code, begin_time: str, end_time: str, loan_rolling=1, inteRate_rolling=1, fig=False):
    w.start()
    loan_begin_time = time.strptime(begin_time,"%Y-%m-%d")
    loan_begin_time = datetime(loan_begin_time[0], loan_begin_time[1], loan_begin_time[2]) - \
        dateutil.relativedelta.relativedelta(months=loan_rolling)
    rate_begin_time = time.strptime(begin_time,"%Y-%m-%d")
    rate_begin_time = datetime(rate_begin_time[0], rate_begin_time[1], rate_begin_time[2]) - \
        dateutil.relativedelta.relativedelta(months=inteRate_rolling, days=3)

    loan = GetEDBData(loan_code, beginTime=loan_begin_time, endTime=end_time, usedf=True)
    rate = GetEDBData(interestRate_code, beginTime=rate_begin_time, endTime=end_time, usedf=True)

    return _CurrencyCreditCycle(loan, rate, loan_rolling=loan_rolling, inteRate_rolling=inteRate_rolling, fig=fig)

'''
---周期对照表---
下方对照表按照《方法整理_68》文档
'''
# 美林周期对照表
MLDict = {1: 'Recovery', 2: 'Overheat', 3: 'Stagflation', 4: 'Reflation'}
# 货币信用周期对照表
CCDict = {1: '周期1', 2: '周期2', 3: '周期3', 4: '周期4'}

if __name__ == '__main__':
    # 示例一：美林周期
    ml_result = MLCycle(
        oecd_code='G1000116', 
        cpi_code='M0000562', 
        begin_time='2020-01-31', 
        end_time='2021-01-31', 
        oecd_rolling=1, 
        cpi_rolling=12, 
        fig=False
    )
    print(ml_result)
    
    ## 示例二：货币信用周期
    # cc_result = CurrencyCreditCycle(
    #     loan_code='M0009970',
    #     interestRate_code='M0041653',
    #     begin_time='2020-01-31',
    #     end_time='2021-01-31',
    #     loan_rolling=12,
    #     inteRate_rolling=12,
    #     fig=False
    # )
    # print(cc_result)
