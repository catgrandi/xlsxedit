"""A private Excel instance driven through COM, for authoring fixtures.

Windows only: needs Microsoft Excel and pywin32. ``Excel()`` starts a new
Excel process with ``DispatchEx``, never attaches to one the user is working
in, and refuses to continue if the instance it got already has a workbook
open. Every save goes through Excel with personal information removed
(``filterPrivacy``), and the saved package is scanned for the user's name and
local paths before the script moves on.
"""

from __future__ import annotations

import os
import platform
import subprocess
import time
import winreg
import zipfile
from pathlib import Path

import pythoncom
import win32api
import win32con
import win32process
from win32com.client import DispatchEx

REPO = Path(__file__).resolve().parents[2]
FIXTURES = REPO / "fixtures"

XLSX = 51  # xlOpenXMLWorkbook
XLSM = 52  # xlOpenXMLWorkbookMacroEnabled


class xl:
    """The Excel enumeration values the builders use."""

    SrcRange = 1
    Yes = 1
    Database = 1
    # XlTotalsCalculation
    TotalsNone = 0
    TotalsSum = 1
    TotalsAverage = 2
    # XlSortOrder / XlSortOn
    Ascending = 1
    Descending = 2
    SortOnValues = 0
    # XlDVType / XlDVAlertStyle / XlFormatConditionOperator
    ValidateWholeNumber = 1
    ValidateDecimal = 2
    ValidateList = 3
    ValidateDate = 4
    ValidateCustom = 7
    ValidAlertStop = 1
    ValidAlertWarning = 2
    Between = 1
    Greater = 5
    GreaterEqual = 7
    # XlFormatConditionType
    CellValue = 1
    Expression = 2
    # XlIconSet / XlIcon
    IconSet3TrafficLights1 = 4
    IconRedDownArrow = 3
    IconYellowFlag = 8
    IconGreenCheck = 22
    # XlPlacement
    MoveAndSize = 1
    Move = 2
    FreeFloating = 3
    # MsoAutoShapeType / MsoTextOrientation
    ShapeRectangle = 1
    ShapeOval = 9
    TextHorizontal = 1
    # XlChartType
    ColumnClustered = 51
    # XlPivotFieldOrientation / XlConsolidationFunction
    RowField = 1
    PageField = 3
    Sum = -4157
    # XlSparkType
    SparkLine = 1
    SparkColumn = 2
    # XlDirection
    Down = -4121
    # MsoAutomationSecurity
    ForceDisable = 3


def excel_exe_version(app) -> str:
    """The file version of the running ``EXCEL.EXE``, such as ``16.0.20430.20000``."""
    info = win32api.GetFileVersionInfo(os.path.join(app.Path, "EXCEL.EXE"), "\\")
    ms, ls = info["FileVersionMS"], info["FileVersionLS"]
    return f"{ms >> 16}.{ms & 0xFFFF}.{ls >> 16}.{ls & 0xFFFF}"


class Excel:
    """Context manager around one private, invisible Excel instance."""

    def __enter__(self) -> Excel:
        pythoncom.CoInitialize()
        self.app = DispatchEx("Excel.Application")
        self.pid = win32process.GetWindowThreadProcessId(self.app.Hwnd)[1]
        if self.app.Workbooks.Count:
            self.app = None
            pythoncom.CoUninitialize()
            raise RuntimeError(
                "COM handed back an Excel instance with workbooks open; refusing to touch it"
            )
        self._mine: list = []  # the workbooks this script created or opened
        try:
            self._configure()
        except BaseException:
            self.__exit__(None, None, None)
            raise
        return self

    def _configure(self) -> None:
        self.app.Visible = False
        self.app.DisplayAlerts = False
        self.app.AutomationSecurity = xl.ForceDisable
        bits = "32-bit" if "(x86)" in self.app.Path else "64-bit"
        self.version = f"{excel_exe_version(self.app)} ({bits})"
        self.platform = f"{platform.system()} {platform.release()} {platform.version()} ({platform.machine()})"
        self._private = personal_strings(self.app.UserName)

    def __exit__(self, *exc) -> None:
        """Close this script's workbooks and end Excel, unless someone else's is open in it.

        Windows can route a workbook the user opens to a hidden automation
        instance. Such a workbook is left open, the instance is shown and handed
        to the user, and Excel keeps running.
        """
        app = self.app
        if app is None:
            return
        others = []
        try:
            mine = set()
            for wb in self._mine:
                try:
                    mine.add(wb.FullName)
                except pythoncom.com_error:  # already closed
                    pass
            for wb in list(app.Workbooks):
                if wb.FullName in mine:
                    wb.Close(False)
                else:
                    others.append(wb.FullName)
            if others:
                app.Visible = True
                app.UserControl = True
                print(f"left Excel running: it also holds {', '.join(others)}")
            else:
                app.Quit()
        finally:
            self.app = app = None
            pythoncom.CoUninitialize()
            if not others:
                self._wait_for_exit()

    def _wait_for_exit(self, timeout: float = 30.0) -> None:
        """Wait for this instance's own process to end; terminate it if it hangs."""
        try:
            handle = win32api.OpenProcess(
                win32con.SYNCHRONIZE | win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_TERMINATE,
                False,
                self.pid,
            )
        except win32api.error:  # already gone
            return
        try:
            deadline = time.monotonic() + timeout
            while time.monotonic() < deadline:
                if win32process.GetExitCodeProcess(handle) != 259:  # STILL_ACTIVE
                    return
                time.sleep(0.25)
            win32process.TerminateProcess(handle, 1)
        finally:
            win32api.CloseHandle(handle)

    def new_workbook(self, *sheets: str):
        """A new workbook whose worksheets are named ``sheets``, in order."""
        wb = self.app.Workbooks.Add()
        self._mine.append(wb)
        # Excel's "Include this many sheets" option decides how many a new workbook has
        while wb.Worksheets.Count > len(sheets):
            wb.Worksheets(wb.Worksheets.Count).Delete()
        while wb.Worksheets.Count < len(sheets):
            # positional: pywin32's late binding drops the After= keyword
            wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
        for index, name in enumerate(sheets, start=1):
            wb.Worksheets(index).Name = name
        wb.Worksheets(1).Activate()
        return wb

    def open(self, path: Path):
        wb = self.app.Workbooks.Open(str(Path(path).resolve()))
        self._mine.append(wb)
        return wb

    def save(self, wb, path: Path, *, file_format: int | None = None) -> Path:
        """Save ``wb`` through Excel at ``path``, close it, and check it for personal data."""
        path = Path(path).resolve()  # Excel resolves relative paths against its own folder
        if file_format is None:
            file_format = XLSM if path.suffix.lower() == ".xlsm" else XLSX
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists():
            path.unlink()
        wb.Worksheets(1).Activate()
        wb.Worksheets(1).Range("A1").Select()
        wb.RemovePersonalInformation = True
        wb.SaveAs(str(path), file_format)
        wb.Close(False)
        self.check_private(path)
        return path

    def check_private(self, path: Path) -> None:
        """Fail when any member of ``path`` names the user or a local path."""
        leaks = personal_data(path, self._private)
        if leaks:
            raise RuntimeError(f"{path} names the user or a local path in: {', '.join(leaks)}")


def personal_strings(*extra: str) -> list[str]:
    """What no fixture may contain: user names, git identity, home folder, repository path."""
    found = {
        os.environ.get("USERNAME", ""),
        str(Path.home()),
        str(REPO),
        *_git_identity(),
        *_office_user(),
        *extra,
    }
    return sorted(s for s in found if len(s) >= 3)


def personal_data(path: Path, needles: list[str]) -> list[str]:
    """The members of the package at ``path`` that contain one of ``needles``, ignoring case."""
    leaks = []
    with zipfile.ZipFile(path) as z:
        for member in z.namelist():
            blob = z.read(member).lower()
            for needle in needles:
                for encoding in ("utf-8", "utf-16-le"):
                    if needle.lower().encode(encoding) in blob:
                        leaks.append(f"{member} ({encoding})")
    return sorted(set(leaks))


def _office_user() -> list[str]:
    """The user name Office stamps on documents (File > Options > General)."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\Common\UserInfo") as key:
            return [str(winreg.QueryValueEx(key, "UserName")[0])]
    except OSError:
        return []


def _git_identity() -> list[str]:
    """The global git user name and email, which the scan treats as personal data too."""
    found = []
    for key in ("user.name", "user.email"):
        try:
            out = subprocess.run(
                ["git", "config", "--global", key], capture_output=True, text=True, check=False
            )
        except OSError:
            continue
        if out.stdout.strip():
            found.append(out.stdout.strip())
    return found
