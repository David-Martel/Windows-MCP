"""
Shared constants for the UIAutomation wrapper modules.

Extracted from core.py, controls.py, patterns.py, enums.py to eliminate
duplication and prevent constant drift.
"""

import os
import sys
import time
from typing import Any

METRO_WINDOW_CLASS_NAME = "Windows.UI.Core.CoreWindow"  # for Windows 8 and 8.1
SEARCH_INTERVAL = 0.5  # search control interval seconds
MAX_MOVE_SECOND = 1  # simulate mouse move or drag max seconds
TIME_OUT_SECOND = 10
OPERATION_WAIT_TIME = 0.5
MAX_PATH = 260
DEBUG_SEARCH_TIME = False
DEBUG_EXIST_DISAPPEAR = False
S_OK = 0

IsNT6orHigher = os.sys.getwindowsversion().major >= 6
CurrentProcessIs64Bit = sys.maxsize > 0xFFFFFFFF
ProcessTime = time.perf_counter  # this returns nearly 0 when first call it if python version <= 3.6
ProcessTime()  # need to call it once if python version <= 3.6
TreeNode = Any
