import subprocess
import glob
import os
import sys

tests_dir = os.path.dirname(__file__)
tests = sorted(glob.glob(os.path.join(tests_dir, "[A-Za-z]*.py")))
excludes = ["runtests.py"]

output_dir = os.path.join(tests_dir, "output")
if not os.path.exists(output_dir):
    os.mkdir(output_dir)

failed = []

for test in tests:
    test_name = os.path.basename(test)
    if test_name not in excludes:
        print("%s ..." % test_name)
        completed = subprocess.run([sys.executable, "./%s" % test_name], cwd=tests_dir)
        if completed.returncode:
            failed.append(test_name)

if failed:
    print("Failed: %s" % ", ".join(failed))
    sys.exit(1)

print("Done.")
