import subprocess
import glob

tests = glob.glob('[A-Za-z]*.py')
excludes = ['runtests.py']

for test in tests:
    if test not in excludes:
        print '%s ...' % test
        subprocess.call(['python','./%s' % test])

print 'Done.'
