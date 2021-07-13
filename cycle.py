from WindPy import w
from pandas.core.frame import DataFrame
import matplotlib.pyplot as plt
import time
from datetime import datetime
import dateutil
import numpy as np
import pandas as pd

def GetData(code, begin_time, end_time, source: str) -> dict:
    '''
    ---获取数据函数---
    code: 指标代码
    begin_time: 开始时间，格式只要对应数据源可识别即可
    end_time: 结束时间，格式同上
    source: 数据来源，可选项：['wind']
    '''
    sequence = {'key': [], 'value': []}
    if source == 'wind':    # 从wind获取数据
        error_code, data = w.edb(code, begin_time, end_time, "Fill=Previous", usedf=True)
        if error_code == 0:
            for index, row in data.iterrows():
                sequence['key'].append(index.__str__())
                sequence['value'].append(row[0])
            return sequence
        else:
            raise Exception('指标%s获取失败！' % code)
    # 加入其他数据源： elif source == ...
    else:
        raise Exception('未知的数据源%s！' % source)

def ToDataFrame(sequence: dict, index_name: str, value_name: str):
    # 将序列转为DataFrame并且更改索引名与列名
    df = pd.DataFrame([[v] for v in sequence['value']], index=sequence['key'], columns=[value_name])
    df.index.name = index_name
    return df

'''
---周期对照表---
对字典进行了变形，使得_CalcCycle函数[结果计算]部分的if条件与标识周期的数字的对应变为一样
'''
# 美林周期对照表
MLDict = {1: 'Reflation', 2: 'Recovery', 3: 'Overheat', 4: 'Stagflation'}
# 货币信用周期对照表
CCDict = {1: '周期1', 2: '周期2', 3: '周期3', 4: '周期4'}

def TransformResult(result, method, to_pair=True):
    '''
    ---结果转换函数，将用数字标识的周期转为用字符串标识的周期---
    result: period用数字标识
    method: 采用哪种对照表，可选项：['ML', 'CurrencyCredit', 'skip']
    pair: 是否用pair的方式转换数据，若为True结果格式变为[[日期1, 周期1], [日期2, 周期2]...]
    '''
    pair = []
    if method == 'ML':
        for i in range(len(result['date'])):
            result['period'][i] = MLDict[result['period'][i]]
            pair.append([result['date'][i], result['period'][i]])
    elif method == 'CurrencyCredit':
        for i in range(len(result['date'])):
            result['period'][i] = CCDict[result['period'][i]]
            pair.append([result['date'][i], result['period'][i]])
    # 添加其他周期类型：elif method == ...
    elif method == 'skip':  # 若采用“skip”方法则跳过“period”的转换，直接返回原结果的pair形式
        return pair
    else:
        raise Exception('未知的参数“method”！')

    if to_pair:
        return pair
    else:
        return result


def _CalcCycle(target1, target2, rolling1, rolling2, method, shape=1, to_pair=True, fig=False):
    '''
    ---计算周期的具体实现---
    target1: 用来计算的第一个指标数据，类型为dict，两个维度为'key'和'value'
    target2: 用来计算的第二个指标数据，类型维度同上
    rolling1: 指标一的移动平均数量，单位：月
    rolling2: 指标二的移动平均数量，单位：月
    method: 计算哪个周期，可选项['ML', 'CurrencyCredit']
    shape: 以哪个指标的数据形状为准，可选项[1, 2]
    to_pair: 输出结果是否转成pair的格式，为True转，为False时格式与输入相同，两个维度为'date'和'period'
    fig: 是否展现简单可视化，为True会阻塞程序
    '''
    target1 = ToDataFrame(target1, 'date', 't1')
    target2 = ToDataFrame(target2, 'date', 't2')
    data = target1.join(target2, how='outer')

    ## [数据变形begin]
    # 把两种指标变成相同的格式
    if shape == 1:
        # 给target2的数据填上与target1相比的空缺
        data['t2'] = data['t2'].fillna(method='ffill')
        # 把data['t2']变成data['t1']的形状
        data = data.drop(list(data[np.isnan(data['t1'])].index))
    elif shape == 2:
        data['t1'] = data['t1'].fillna(method='ffill')
        data = data.drop(list(data[np.isnan(data['t2'])].index))
    else:
        raise Exception('未知的参数“shape”！')
    ## [数据变形end]

    data['t1'] = data['t1'].rolling(rolling1).mean().diff()
    data['t2'] = data['t2'].rolling(rolling2).mean().diff()
    data = data.dropna(axis=0, how='any')

    # [结果计算begin]
    result = {'date': [], 'period': []}
    for index, row in data.iterrows():
        l = row['t1']
        r = row['t2']
        result['date'].append(index.__str__())
        if l < 0 and r < 0:
            result['period'].append(1)
        elif l > 0 and r < 0:
            result['period'].append(2)
        elif l > 0 and r > 0:
            result['period'].append(3)
        elif l < 0 and r > 0:
            result['period'].append(4)
        else:   # 出现0沿用上一个周期
            result['period'].append(result['period'][-1])
    # [结果计算end]

    # [可视化begin]
    if fig:
        fig = plt.figure(figsize=(15,5))
        ax = fig.add_subplot(111)
        data['t1'].plot(ax=ax, grid=True, label='target1', marker='o')
        data['t2'].plot(ax=ax, grid=True, secondary_y=True, label='target2', marker='*')
        plt.title('figure')
        plt.show()
    # [可视化end]
    
    # 算完得到的result的period是用数字标识的，如果仅需要转换成pair形式可这样调用
    # return TransformResult(result, 'skip', to_pair)
    return TransformResult(result, method, to_pair) # 转换输出格式

def CalcCycle(code1: str, code2: str, begin_time: str, end_time:str, rolling1=1, rolling2=1, method='ML', shape=1, to_pair=True, fig=False):
    '''
    ---周期计算函数---
    code1: 指标一代码，在使用美林周期时代表OECD指标，在使用货币信用周期时代表贷款指标
    code2: 指标二代码，在使用美林周期时代表CPI指标，在使用货币信用周期时代表利率指标
    begin_time: 要得到的结果的开始时间，形式如：'2020-1-31'（年月日写全）
    end_time: 要得到的结果的结束时间，形式同上
    rolling1: 指标一的移动平均数量，单位：月
    rolling2: 指标二的移动平均数量，单位：月
    method: 计算哪个周期，可选项['ML', 'CurrencyCredit']
    shape: 以哪个指标的数据形状为准，可选项[1, 2]
    to_pair: 输出结果是否转成pair的格式，为True转，为False时格式与GetData返回值相同，两个维度为'date'和'period'
    fig: 是否展现简单可视化，为True会阻塞程序
    '''
    ## [开始时间计算begin]
    # 由于要进行rolling，所以往往需要向前多拿几个月数据，为方便直接按大的来
    bigger = rolling1 if rolling1 > rolling2 else rolling2
    begin_time = time.strptime(begin_time,"%Y-%m-%d")
    begin_time = datetime(begin_time[0], begin_time[1], begin_time[2]) - \
        dateutil.relativedelta.relativedelta(months=bigger)
    ## [开始时间计算end]

    ## [数据获取begin]
    target1 = GetData(code1, begin_time, end_time, source='wind')
    target2 = GetData(code2, begin_time, end_time, source='wind')
    ## [数据获取end]

    return _CalcCycle(target1, target2, rolling1, rolling2, method, shape, to_pair, fig)

if __name__ == '__main__':
    w.start()

    # 示例一：美林周期
    print('示例一：美林周期')
    result = CalcCycle(
        code1='G1000116', 
        code2='M0000562', 
        begin_time='2020-01-31', 
        end_time='2021-01-31', 
        rolling1=1, 
        rolling2=12, 
        method='ML'
    )
    print(result)

    # 示例二：货币信用周期
    print('示例二：货币信用周期')
    result = CalcCycle(
        code1='M0009970', 
        code2='M0041653', 
        begin_time='2020-01-31', 
        end_time='2021-01-31', 
        rolling1=12, 
        rolling2=12, 
        method='CurrencyCredit',
    )
    print(result)

