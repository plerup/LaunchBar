/*
Copyright (C) 2012-2014 Peter Lerup

This file is part of LaunchBar.

LaunchBar is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>
*/

#include <atlstr.h>

#define CURR_INSTANCE GetModuleHandle(NULL)

#define EMPTY_STR _T("")
#define EMPTY_CSTR (CString)_T("")
#define BS _T("\\")
#define APPEND_BS(s) if (!s.IsEmpty() && s.Right(1) != BS) s+= BS
#define TRIM_BS(s) if (!s.IsEmpty() && s.Right(1) == BS) s.Delete(s.GetLength()-1)
#define CRLF _T("\r\n")

#define STR_LEN(s) (UINT)_tcslen(s)
#define STR_STR(s, p) _tcsstr(s, p)
#define STR_CPY(dst, src) _tcscpy(dst, src)
#define STR_CAT(dst, src) _tcscat(dst, src)
#define STR_SIZE(s) (sizeof(s)/sizeof(TCHAR))

#define round(expr) (LONG)((expr)+0.5)
#define IN_RANGE(val, _min, _max) val = min(max(val, _min), _max);

#define SHORTCUT_EXT _T(".lnk")
#define IS_SHORTCUT(fileName) (GetFileNameComp(fileName, eFcType).CompareNoCase(SHORTCUT_EXT) == 0)

#define GET_REG_INT(key, val) _stscanf_s(GetRegVal(key), _T("%d"), &val);
#define SET_REG_INT(key, val) { CString buf; buf.Format(_T("%d"), val); SetRegVal(key, buf); }

BOOL ExpEnvVars(CString& str);

BOOL SetAppRegRoot(LPCTSTR root);
CString GetRegVal(LPCTSTR valName, LPCTSTR defVal = EMPTY_STR, HKEY rootKey = NULL);
BOOL SetRegVal(LPCTSTR valName, LPCTSTR val, HKEY rootKey = NULL);
DWORD RecDelRegKey(LPCTSTR pKeyName = NULL, HKEY hStartKey = NULL);
BOOL DelRegVal(LPCTSTR valName, HKEY rootKey = NULL);

enum tFileComp {eFcDrive = 1, eFcDir = 2, eFcName = 4, eFcType = 8};
CString GetFileNameComp(LPCTSTR fileName, WORD type);
BOOL FindFiles(LPCTSTR path, CString& fileName, HANDLE *findInfo);
BOOL LocateFile(CString& path);
BOOL FileExists(LPCTSTR fileName);
BOOL IsDir(LPCTSTR name);
BOOL FileIsHidden(LPCTSTR name);
time_t FileModTime(LPCTSTR name);
BOOL CopyDir(LPCTSTR src, LPCTSTR dst);

HWND CreateTooltip(HWND hWnd, LPCTSTR text);

BOOL CreateShortcut(LPCTSTR linkPath, LPCTSTR targetPath, LPCTSTR comment = NULL, LPCTSTR arguments = NULL,
                    LPCTSTR workDir = NULL, LPCTSTR iconFile = NULL, int iconIndex = 0, int showCommand = SW_SHOWNORMAL);
BOOL GetShortcutInfo(LPCTSTR path, CString& targetPath, HICON *largeIcon = NULL, HICON *smallIcon = NULL, CString* comment = NULL);
BOOL GetFileIcons(LPCTSTR fileName, HICON *largeIcon = NULL, HICON *smallIcon = NULL);
CString GetTrueTarget(LPCTSTR path);

void InitPopupMenu(HMENU hMenu, BOOL useMenuCom = FALSE);
BOOL AddMenuItem(HMENU hMenu, LPCTSTR text, HMENU hSubMenu = NULL);
BOOL SetMenuItemData(HMENU hMenu, INT itemID, HICON smallIcon, HICON largeIcon = NULL, BOOL useLarge = FALSE, PVOID extra = NULL);
PVOID GetMenuItemData(HMENU hMenu, INT itemID);
BOOL SetLargeMenus(BOOL on);
BOOL HandleMenuItemIconMessage(UINT message, LPARAM lParam);
BOOL DestroyMenuData(HMENU hMenu, BOOL destroyMenu = FALSE);

BOOL DoRun(LPCTSTR com, LPCTSTR params, HWND hWnd, WORD showType = SW_SHOWNORMAL, LPCTSTR verb = NULL);

CString GetQuickLaunchDir();
CString GetAutoStartDir(BOOL allUsers = FALSE);
CString GetStartMenuDir(BOOL allUsers = FALSE);

BOOL StartWatchDir(LPCTSTR dir, HWND hWnd, DWORD messageID, LPVOID watchHand);
BOOL EndWatchDir(LPVOID watchHand);

CString LoadFormatResString(UINT ID, ...);
CString ErrorString(DWORD err = GetLastError());

enum tMessSeverity { eInfo, eWarning, eError, eFatal };
BOOL ShowMessage(LPCTSTR message, tMessSeverity sev = eInfo, HWND hWnd = NULL, LPCTSTR header = NULL);

void PositionDialog(HWND hDlg);
void CenterDlg(HWND hDlg);
CString PromptFileName(LPCTSTR startFileName, HWND fatherWind = NULL, BOOL save = FALSE, DWORD typeListID = IDS_ALL_FILES, BOOL center = FALSE);

BOOL IsWinPE();

BOOL RunDialog(HWND hWnd);
