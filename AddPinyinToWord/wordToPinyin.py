import sys, getopt
import pypinyin

word = sys.argv[1:]

s = ''
for i in pypinyin.pinyin(word):
    s = s + ''.join(i) + " "

print(s)

