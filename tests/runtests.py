import subprocess
import glob
import six

tests = sorted(glob.glob('[A-Za-z]*.py'))
excludes = ['runtests.py']

for test in tests:
    if test not in excludes and test.startswith('unescape'):
        six.print_('%s ...' % test)
        subprocess.call(['python','./%s' % test])

six.print_('Done.')
