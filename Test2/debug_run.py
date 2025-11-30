import os
from Test2 import process_leaderboard

src = os.path.join(os.getcwd(), 'leaderboard.xlsx')
out = 'Test2/output_debug'
print('SRC', src)
try:
    res = process_leaderboard.run(src, out)
    print('RESULT:', res)
except Exception as e:
    import traceback
    traceback.print_exc()
    raise
