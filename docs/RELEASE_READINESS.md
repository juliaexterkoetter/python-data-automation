# Version 1 Release Readiness

## Validated Environment

Version 1 was validated on 2026-08-21 with:

- Python 3.14.4;
- openpyxl 3.1.5;
- defusedxml 0.7.1;
- pytest 9.1.1.

The exact direct dependencies are pinned in `requirements.txt` and `requirements-dev.txt`. The development requirements include the runtime requirements so one installation reproduces the direct dependency set used by the test suite.

pytest declares Python 3.14 support. openpyxl 3.1.5 and defusedxml 0.7.1 have Python version metadata compatible with this interpreter, but their published classifiers do not specifically verify Python 3.14. Formal dependency-specific Python 3.14 support is therefore **NOT VERIFIED**. Actual compatibility in this project is demonstrated by the complete suite running successfully on Python 3.14.4.

## Installation

```bash
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements-dev.txt
.venv/bin/python -m pip check
.venv/bin/python -m pytest
```

Exact direct pins improve reproducibility but are not a hash-locked supply-chain guarantee. Transitive packages selected by the installer may change while remaining compatible with the direct pins.

## Dependency Vulnerability Audit

The release-readiness audit used PyPA pip-audit 2.10.1 from a temporary tool environment:

```bash
python3 -m venv /tmp/python-data-audit-venv
/tmp/python-data-audit-venv/bin/python -m pip install pip-audit==2.10.1
/tmp/python-data-audit-venv/bin/python -m pip_audit -r requirements.txt
/tmp/python-data-audit-venv/bin/python -m pip_audit -r requirements-dev.txt
```

When pip-audit is already installed in an isolated tool environment, the equivalent project-directory commands are:

```bash
python -m pip_audit -r requirements.txt
python -m pip_audit -r requirements-dev.txt
```

Both commands returned `No known vulnerabilities found` on 2026-08-21. Repeat the audit before a later release because advisory databases and dependency resolution change over time.

pip-audit checks resolved Python distributions against known-vulnerability services. It is not a static code analyzer, does not establish that packages are trustworthy, may not cover vulnerabilities in non-Python components, and cannot prove the absence of unknown or unreported vulnerabilities.

Primary references:

- pip requirements files: <https://pip.pypa.io/en/stable/reference/requirements-file-format/>
- pip-audit usage and security model: <https://pypi.org/project/pip-audit/>
- openpyxl package metadata: <https://pypi.org/project/openpyxl/>
- defusedxml package metadata: <https://pypi.org/project/defusedxml/>
- pytest package metadata: <https://pypi.org/project/pytest/>

## Atomicity and Durability

The exporter creates the report in a temporary file in the destination directory, closes and validates it, and then publishes it with `os.replace`. Normal construction, conversion, save, close, validation, and pre-replacement publication failures preserve an existing report. A successful replace is atomic when the filesystem provides the documented rename/replace contract.

This is not a power-loss durability guarantee. Version 1 does not call `fsync`, does not promise persistence after a power loss, kernel crash, or physical storage failure, and does not automatically recover temporary files after abrupt process termination.

Primary reference: <https://docs.python.org/3/library/os.html#os.replace>

## Trusted Input Directory and Residual TOCTOU Risk

`data/input/` is a local directory controlled by the trusted operator. The application rejects symlinks that are visible during normal discovery. Because discovery, inspection, and file opening are separate path-based operations, a local process with write access could replace an entry during a run.

Version 1 accepts this residual time-of-check/time-of-use risk. Operators must not allow untrusted local writers to modify the input directory during processing. Defending a hostile shared directory with descriptor-relative traversal, `O_NOFOLLOW`, or private input copies is outside the approved Version 1 architecture.

Reference: <https://owasp.org/www-community/pages/vulnerabilities/race_conditions>

## Manual Release Check

Automated tests validate workbook structure and logical content programmatically. They do not replace opening the result in Microsoft Excel. Use `EXCEL_SMOKE_TEST.md` for the required manual Windows Excel check; do not mark that check complete without performing it in the real application.
