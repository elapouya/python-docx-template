import glob
import os
import subprocess

# Change work dir if not in tests/
tests_dir = os.path.dirname(__file__)
if os.getcwd() != tests_dir:
    os.chdir(tests_dir)

tests = sorted(glob.glob("[A-Za-z]*.py"))
excludes = ["runtests.py"]

output_dir = os.path.join(os.path.dirname(__file__), "output")
if not os.path.exists(output_dir):
    os.mkdir(output_dir)

for test in tests:
    if test not in excludes:
        print("%s ..." % test)
        subprocess.check_call(["python", "./%s" % test])

print("Done.")
