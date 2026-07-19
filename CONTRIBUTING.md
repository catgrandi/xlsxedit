# Contributing to xlsxedit

Thanks for your interest in improving xlsxedit! Contributions of all kinds are
welcome: bug reports, documentation, tests, and code.

## Contributor License Agreement

Before your code can be merged, you must agree to the
[Contributor License Agreement](CLA.md). This keeps the project legally clean
and preserves the maintainer's ability to license the project. To sign, add
this line to your pull request description:

> I have read the CLA document and I hereby sign the CLA.

(along with your full name and the date).

## Development setup

```bash
git clone https://github.com/<your-username>/xlsxedit.git
cd xlsxedit
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev,pandas]"
```

## Running the tests

```bash
pytest
```

Please make sure the full test suite passes before opening a pull request, and
add tests for any new behavior.

## Guidelines

- Keep changes focused; one logical change per pull request.
- Match the existing code style (type hints, no unnecessary comments).
- Preserve template fidelity: edits should change only what the API targets and
  leave unrelated package parts untouched.
- If you touch the OPC layer (`src/xlsxedit/opc/**`), remember parts of it are
  adapted from python-docx / python-pptx (MIT); keep the attribution intact.

## Reporting bugs

Open an issue with a minimal reproduction, the xlsxedit version, your Python
version, and (if possible) a sample `.xlsx` that triggers the problem.
