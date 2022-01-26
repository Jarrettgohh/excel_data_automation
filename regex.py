import re

string = '100Hz_1.2V'
pattern = '100Hz'

# match = re.search(re.compile(f'[-|_]{pattern}[-|_]'), string)
match = re.search(re.compile(f'(^|\-|\_|\.){pattern}(\-|\_|\.|$)'), string)
# match = re.search(re.compile(f'(^|(\-|\_|\.)){pattern}'), string)

if match: print('match')
else: print('no match')