import glob
import os
import subprocess

tests = sorted(glob.glob("[A-Za-z]*.py"))
excludes = ["runtests.py"]

output_dir = os.path.join(os.path.dirname(__file__), "output")
if not os.path.exists(output_dir):
    os.mkdir(output_dir)

for test in tests:
    if test not in excludes:
        print(f"{test} ...")
        subprocess.call(["python", f"./{test}"])

print("Done.")
