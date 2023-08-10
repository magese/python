import datetime
import time


# 组装日志
def msg(k, *p):
    k = '{} - ' + k
    k = k.replace('%', '%%').replace('{}', '%s')
    return k % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), *p)


# 日志打印
def info(k, *p):
    print(msg(k, *p))


# 循环日志
def loop_msg(current, total, pre):
    holder = '%s/%s %s%% cost:%ss'
    return holder % (current, total, format(current / total * 100, '.2f'), format((time.perf_counter() - pre), '.2f'))
