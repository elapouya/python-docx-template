import glob
import os
import subprocess
from pathlib import Path

# Change work dir if not in tests/
tests_dir = Path(__file__).parent.resolve()
if Path.cwd() != tests_dir:
    os.chdir(tests_dir)

tests = sorted(glob.glob("[A-Za-z]*.py"))
excludes = ["runtests.py"]

output_dir = tests_dir / "output"
if not output_dir.exists():
    os.mkdir(output_dir)

for test in tests:
    if test not in excludes:
        print("%s ..." % test)
        subprocess.check_call(["python", "./%s" % test])

print("Done.")
