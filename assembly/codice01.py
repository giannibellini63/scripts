__author__ = 'bellini'

'''
https://docs.python.org/2.7/library/dis.html
'''

import dis

def fn_expressive(upper = 10000000):
    total = 0
    for i in range(upper):
        total += i
    return total

def fn_terse(upper = 10000000):
    return sum(range(upper))

if __name__ == "__main__":
    print ("Function return same result: ", fn_expressive())
    dis.dis(fn_expressive)
    print ("Function return same result: ", fn_terse())
    dis.dis(fn_terse)