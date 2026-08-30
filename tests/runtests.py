from hashlib import sha256
import json
from pathlib import Path
import subprocess
import sys


def file_sha256(path):
    digest = sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


tests_dir = Path(__file__).resolve().parent
repo_root = tests_dir.parent
launcher_path = Path(__file__).resolve()
tests = sorted(
    path
    for path in tests_dir.glob("[A-Za-z]*.py")
    if path.name != launcher_path.name
)

(tests_dir / "output").mkdir(exist_ok=True)
records = []
for test in tests:
    relative_path = test.relative_to(repo_root).as_posix()
    test_sha256 = file_sha256(test)
    print("RUN %s %s" % (relative_path, test_sha256), flush=True)
    subprocess.run(
        [sys.executable, str(test)],
        cwd=str(tests_dir),
        check=True,
    )
    print("PASS %s %s" % (relative_path, test_sha256), flush=True)
    records.append({"path": relative_path, "sha256": test_sha256})

completion = {
    "status": "passed",
    "count": len(records),
    "launcher": {
        "path": launcher_path.relative_to(repo_root).as_posix(),
        "sha256": file_sha256(launcher_path),
    },
    "tests": records,
}
print(json.dumps(completion, sort_keys=True, separators=(",", ":")), flush=True)
