import sys


n=len(sys.argv)
print("total number of arguments: ",n)

for argument in sys.argv:
    print(argument)

if __name__ == '__main__':
    print('executed')