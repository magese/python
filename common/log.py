import datetime
import time


# 日志打印
def info(k, *p):
    k = '{} - ' + k
    k = k.replace('%', '%%').replace('{}', '%s')
    print(k % (datetime.datetime.now(), *p))


# 循环日志
def loop_msg(current, total, pre):
    holder = 'loop progress: %s/%s %s%% cost:%ss'
    return holder % (current, total, format(current / total * 100, '.2f'), format((time.perf_counter() - pre), '.2f'))
