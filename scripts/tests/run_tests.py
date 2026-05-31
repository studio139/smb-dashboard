# -*- coding: utf-8 -*-
"""Dependency-free test runner — builds the fixtures once, runs every test_* function,
prints an ASCII PASS/FAIL table, and exits non-zero if anything fails (blocks done).

    C:\\Python314\\python.exe scripts\\tests\\run_tests.py
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)
sys.stdout.reconfigure(encoding="utf-8")

import harness  # noqa: E402
import test_numbers  # noqa: E402
import test_parity  # noqa: E402
import test_reconciliation  # noqa: E402
import test_expense_sheet  # noqa: E402
import test_determinism  # noqa: E402
import test_structure  # noqa: E402
import test_business_logic  # noqa: E402

MODULES = [test_numbers, test_parity, test_reconciliation, test_expense_sheet,
           test_determinism, test_structure, test_business_logic]


def main():
    print("building test fixtures (load -> compute -> xlsx/html into outputs/_test) ...")
    ctx = harness.get_ctx()
    results = []
    for mod in MODULES:
        short = mod.__name__
        for name in sorted(vars(mod)):
            fn = getattr(mod, name)
            if not (name.startswith("test_") and callable(fn)):
                continue
            qual = "%s.%s" % (short, name)
            try:
                fn(ctx)
                results.append((qual, "PASS", ""))
            except AssertionError as e:
                results.append((qual, "FAIL", str(e)[:150]))
            except Exception as e:  # noqa: BLE001
                results.append((qual, "FAIL", "ERR " + repr(e)[:150]))

    width = max([len(q) for q, _, _ in results] + [12])
    bar = "+-" + "-" * width + "-+------+"
    print("")
    print(bar)
    for qual, status, detail in results:
        line = "| %-*s | %-4s |" % (width, qual, status)
        if detail:
            line += " " + detail.encode("ascii", "replace").decode()
        print(line)
    print(bar)
    n_fail = sum(1 for _, s, _ in results if s == "FAIL")
    print("TESTS: %d PASS / %d FAIL (of %d)" % (len(results) - n_fail, n_fail, len(results)))
    return 1 if n_fail else 0


if __name__ == "__main__":
    sys.exit(main())
