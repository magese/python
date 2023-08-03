import time

T1 = time.perf_counter()
time.sleep(1)
T2 = time.perf_counter()
print((T2 - T1))
print(format((T2 - T1), '.2f'))
print('%.2f'%((T2 - T1) / 1000 / 1000))
