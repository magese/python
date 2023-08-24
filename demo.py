def tuple_test(first, *second):
    if second:
        print(True)
        print(float(second[0]))
    else:
        print(False)
    print(first)
    print(second)


tuple_test('magese', 3.14)
print('-------------')
tuple_test('magese', 3.14, 0.003)
print('-------------')
tuple_test('zora')

