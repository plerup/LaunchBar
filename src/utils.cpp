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

// utils.cpp
// Generic application utility functions
//

#include <windows.h>
#include <tchar.h>
#include <ShlObj.h>
#include  <io.h>

#include <commoncontrols.h>

#include "resource.h"

#include "utils.h"

#define BUFF_SIZE 1024

LPCTSTR AppRegRoot = NULL;
HKEY gRegRootKey = HKEY_CURRENT_USER;
#define GET_ROOT_KEY(hKey) (hKey ? hKey : gRegRootKey)

BOOL ExpEnvVars(CString& str)
{
   // Expand possible environment variables in the specified string
   TCHAR tmp[BUFF_SIZE];
   DWORD cnt = ExpandEnvironmentStrings(str, tmp, STR_SIZE(tmp));
   str = tmp;
   return (cnt > 0);
}

//--------------------------------------------------------------------------

BOOL SetAppRegRoot(LPCTSTR root)
{
   // Set default root for registry operations below
   AppRegRoot = root;
   return TRUE;
}

//--------------------------------------------------------------------------

void EncodeRegName(CString keyName, CString& path, CString& name)
{
   // Split name and path of a registry key name
   int pos = keyName.ReverseFind('\\');
   if (pos >= 0)
   {
      path = keyName.Left(pos+1);
      name = keyName.Right(keyName.GetLength()-pos-1);
   }
   else
   {
      path = AppRegRoot;
      name = keyName;
   }
}

//--------------------------------------------------------------------------

CString GetRegVal(LPCTSTR valName, LPCTSTR defVal, HKEY rootKey)
{
   // Get specified string registry value
   BOOL  ok = FALSE;
   HKEY  hKey;
   DWORD type, size;
   CString path, name, val = EMPTY_STR;

   EncodeRegName(valName, path, name);

   if (RegOpenKeyEx(GET_ROOT_KEY(rootKey), path, 0, KEY_QUERY_VALUE, &hKey) == ERROR_SUCCESS)
   {
      size = 0;
      RegQueryValueEx(hKey, name, 0, &type, NULL, &size);
      ok = RegQueryValueEx(hKey, name, 0, &type, (LPBYTE)val.GetBuffer(size/sizeof(TCHAR)), &size) == ERROR_SUCCESS;
      val.ReleaseBuffer();
      RegCloseKey(hKey);
   }
   if (!ok)
      val = defVal;
   return val;
}

//--------------------------------------------------------------------------

BOOL SetRegVal(LPCTSTR valName, LPCTSTR val, HKEY rootKey)
{
   // Set specified string registry value
   BOOL ok = FALSE;
   HKEY  hKey;
   DWORD disp, size;

   CString path, name;
   EncodeRegName(valName, path, name);

   if (RegCreateKeyEx(GET_ROOT_KEY(rootKey), path, 0, EMPTY_STR, REG_OPTION_NON_VOLATILE,
          KEY_ALL_ACCESS, NULL, &hKey, &disp) == ERROR_SUCCESS)
   {
      size = (DWORD)(STR_LEN(val)+1)*sizeof(TCHAR);
      ok = (RegSetValueEx(hKey, name, 0, REG_SZ, (BYTE*)val, size) == ERROR_SUCCESS);
      RegCloseKey(hKey);
   }
   return ok;
}

//--------------------------------------------------------------------------

DWORD RecDelRegKey(LPCTSTR pKeyName, HKEY hStartKey)
{
   // Delete all registry keys below the one specified
   DWORD   dwRtn, dwSubKeyLength;
   TCHAR   szSubKey[255];
   HKEY    hKey;

   if (!hStartKey) hStartKey = gRegRootKey;
   if (!pKeyName) pKeyName = AppRegRoot;

   if (pKeyName && STR_LEN(pKeyName))
   {
      if((dwRtn=RegOpenKeyEx(hStartKey,pKeyName, 0, KEY_ENUMERATE_SUB_KEYS | DELETE, &hKey )) == ERROR_SUCCESS)
      {
         while (dwRtn == ERROR_SUCCESS)
         {
            dwSubKeyLength = STR_SIZE(szSubKey);
            dwRtn=RegEnumKeyEx(
                           hKey,
                           0,       // always index zero
                           szSubKey,
                           &dwSubKeyLength,
                           NULL,
                           NULL,
                           NULL,
                           NULL
                         );
 
            if(dwRtn == ERROR_NO_MORE_ITEMS)
            {
               dwRtn = RegDeleteKey(hStartKey, pKeyName);
               break;
            }
            else if (dwRtn == ERROR_SUCCESS)
               dwRtn = RecDelRegKey(szSubKey, hKey);
         }
         RegCloseKey(hKey);
      }
   }
   else
      dwRtn = ERROR_BADKEY;
 
   return dwRtn;
}

//--------------------------------------------------------------------------

BOOL DelRegVal(LPCTSTR valName, HKEY rootKey)
{
   // Delete registry key value
   BOOL  ok = FALSE;
   HKEY  hKey;
   CString path, name;;

   EncodeRegName(valName, path, name);

   if (RegOpenKeyEx(GET_ROOT_KEY(rootKey), path, 0, KEY_ALL_ACCESS, &hKey) == ERROR_SUCCESS)
   {
      ok = RegDeleteValue(hKey, name) == ERROR_SUCCESS;
      RegCloseKey(hKey);
   }
   return ok;
}

//--------------------------------------------------------------------------

CString GetFileNameComp(LPCTSTR fileName, WORD type)
{
   // Get file name component
   CString compStr = EMPTY_STR;
   TCHAR drive[_MAX_DRIVE], dir[_MAX_DIR], name[_MAX_FNAME], ext[_MAX_EXT];
   _tsplitpath_s(fileName[0] == _T('"') ? &fileName[1] : fileName, drive, _MAX_DRIVE, dir, _MAX_DIR, name, _MAX_FNAME, ext, _MAX_EXT);
   if (type & eFcDrive)
      compStr += drive;
   if (type & eFcDir)
      compStr += dir;
   if (type & eFcName)
      compStr += name;
   if (type & eFcType)
     compStr += ext;
   compStr.Remove(_T('"'));
   return compStr;
}

//--------------------------------------------------------------------------

BOOL FindFiles(LPCTSTR path, CString& fileName, HANDLE *findInfo)
{
   // Expand possible wild card file specifications
   // File name must be initially empty
   BOOL ok = FALSE;
   WIN32_FIND_DATA findData;
   CString root = GetFileNameComp(path, eFcDrive | eFcDir);
   do
   {
      if (!ok && fileName.IsEmpty())
         ok = ((*findInfo = FindFirstFile(path, &findData)) != INVALID_HANDLE_VALUE);
      else
         ok = FindNextFile(*findInfo, &findData);
      if (ok)
      {
         fileName = root;
         fileName += findData.cFileName;
      }
      else if (*findInfo != INVALID_HANDLE_VALUE)
         FindClose(*findInfo);
   } while (ok && (!_tcscmp(findData.cFileName, _T(".")) || !_tcscmp(findData.cFileName, _T(".."))));

   return ok;
}

//--------------------------------------------------------------------------

BOOL LocateFile(CString& path)
{
   if (FileExists(path))
      return TRUE;
   BOOL found = PathFindOnPath(path.GetBuffer(MAX_PATH+1), NULL);
   path.ReleaseBuffer();
   return found;
}

//--------------------------------------------------------------------------

BOOL FileExists(LPCTSTR fileName)
{
   return PathFileExists(fileName);
}

//--------------------------------------------------------------------------

BOOL IsDir(LPCTSTR name)
{
   int attr = GetFileAttributes(name);
   return attr == -1 ? FALSE : (attr & FILE_ATTRIBUTE_DIRECTORY);
}

//--------------------------------------------------------------------------

BOOL FileIsHidden(LPCTSTR name)
{
   int attr = GetFileAttributes(name);
   return attr == -1 ? FALSE : (attr & FILE_ATTRIBUTE_HIDDEN);
}

//--------------------------------------------------------------------------

time_t FileModTime(LPCTSTR name)
{
	struct _stat buf;
	return _tstat(name, &buf) ? 0 : buf.st_mtime;

}

//--------------------------------------------------------------------------

BOOL CopyDir(LPCTSTR src, LPCTSTR dst)
{
   SHFILEOPSTRUCT sf;
   CString _src = CString(src) + BS + _T("*.*"),
           _dst = CString(dst);
   memset(&sf, 0, sizeof(sf));
   sf.hwnd = NULL;
   sf.wFunc = FO_COPY;
   DWORD len = _src.GetLength();
   PTCHAR p = _src.GetBufferSetLength(len+2);
   p[len+1] = 0;
   sf.pFrom = p;
   len = _dst.GetLength();
   p = _dst.GetBufferSetLength(len+2);
   p[len+1] = 0;
   sf.pTo = p;
   sf.fFlags = FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR | FOF_NOERRORUI | FOF_SILENT;
   DWORD res = SHFileOperation(&sf);
   return (res == 0);
}

//--------------------------------------------------------------------------

HWND CreateTooltip(HWND hWnd, LPCTSTR text)
{
   // Create a tooltip window
   HWND hTipWin;
   TOOLINFO toolInfo;

   hTipWin = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, NULL, WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
                            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, hWnd, NULL, NULL, NULL);
   ZeroMemory(&toolInfo, sizeof(toolInfo));
   toolInfo.cbSize = sizeof(toolInfo);
#ifdef UNICODE
   toolInfo.cbSize -= sizeof(void*); // Bug in Visual Studio declaration
#endif
   toolInfo.uFlags = TTF_SUBCLASS;
   toolInfo.hwnd = hWnd;
   toolInfo.hinst = CURR_INSTANCE;
   toolInfo.lpszText = (LPTSTR)text;
   SystemParametersInfo(SPI_GETWORKAREA, 0, &toolInfo.rect, 0);

   SendMessage(hTipWin, TTM_ADDTOOL, 0, (LPARAM)(LPTOOLINFO)&toolInfo);
   // Extend length
   SendMessage(hTipWin, TTM_SETMAXTIPWIDTH, 0, (LPARAM)(INT)2048);

   return hTipWin;
}

//--------------------------------------------------------------------------

BOOL CreateShortcut(LPCTSTR linkPath, LPCTSTR targetPath, LPCTSTR comment, LPCTSTR arguments,
                    LPCTSTR workDir, LPCTSTR iconFile, int iconIndex, int showCommand)
{
   // Create a shortcut file
   HRESULT hRes; 
   IShellLink* psl; 

   hRes = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID*)&psl); 
   if (SUCCEEDED(hRes))
   { 
      IPersistFile* ppf; 

      psl->SetPath(targetPath); 
      if (arguments)
         psl->SetArguments(arguments);
      if (workDir)
         psl->SetWorkingDirectory(workDir);
      if (iconFile || iconIndex)
         psl->SetIconLocation(iconFile ? iconFile : targetPath, iconIndex);
      if (comment)
         psl->SetDescription(comment);

      psl->SetShowCmd(showCommand);

      hRes = psl->QueryInterface(IID_IPersistFile, (LPVOID*)&ppf); 

      if (SUCCEEDED(hRes))
      {
#ifndef UNICODE
         WCHAR wsz[BUFF_SIZE]; 
         MultiByteToWideChar(CP_ACP, 0, linkPath, -1, wsz, BUFF_SIZE); 

         hRes = ppf->Save(wsz, TRUE);
#else
         hRes = ppf->Save(linkPath, TRUE);
#endif
         ppf->Release(); 
      } 
      psl->Release(); 
   } 
   return TRUE; 
} 

//--------------------------------------------------------------------------

BOOL GetShortcutInfo(LPCTSTR path, CString& targetPath, 
                     HICON *largeIcon, HICON *smallIcon, CString *comment)
{
   // Get the attributes of the specified shortcut
   HRESULT hRes; 
   IShellLink* psl;
   CString iconFileName;
   int iconInd;
   SHFILEINFO inf;

   hRes = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID*)&psl); 
   if (SUCCEEDED(hRes))
   { 
      IPersistFile* ppf; 

      hRes = psl->QueryInterface(IID_IPersistFile, (LPVOID*)&ppf); 

      if (SUCCEEDED(hRes))
      {
#ifndef UNICODE
         WCHAR wsz[BUFF_SIZE]; 
         MultiByteToWideChar(CP_ACP, 0, path, -1, wsz, BUFF_SIZE); 
         hRes = ppf->Load(wsz, STGM_READ);
#else
         hRes = ppf->Load(path, STGM_READ);
#endif
      }
      iconInd = 0;
      hRes = psl->GetPath(targetPath.GetBuffer(BUFF_SIZE), BUFF_SIZE, NULL, SLGP_RAWPATH);
      targetPath.ReleaseBuffer();
      if (!targetPath.IsEmpty())
         ExpEnvVars(targetPath);

      if (largeIcon && smallIcon)
      {
         hRes = psl->GetIconLocation(iconFileName.GetBuffer(BUFF_SIZE), BUFF_SIZE, &iconInd);
         iconFileName.ReleaseBuffer();
         if (!iconFileName.IsEmpty())
            ExpEnvVars(iconFileName);
         if (iconFileName.IsEmpty() && targetPath.IsEmpty())
         {
            SHGetFileInfo(path, 0, &inf, sizeof(inf), SHGFI_ICONLOCATION);
            iconFileName = inf.szDisplayName;
            iconInd = inf.iIcon;
         }
         if (iconFileName.IsEmpty())
         {
            GetFileIcons(targetPath, largeIcon, smallIcon);
         }
         else
         {
            if (iconInd == -1)
               iconInd = 0;
            hRes = ExtractIconEx(iconFileName, iconInd, largeIcon, smallIcon, 1);
            if (hRes == -1)
               GetFileIcons(path, largeIcon, smallIcon);
         }
      }

      if (comment)
      {
         *comment = GetFileNameComp(path, eFcName);
         CString descr;
         hRes = psl->GetDescription(descr.GetBuffer(BUFF_SIZE), BUFF_SIZE);
         descr.ReleaseBuffer();
         if (!descr.IsEmpty())
         {
            if (isdigit(comment->GetAt(0)) && isdigit(comment->GetAt(1)))
               *comment = EMPTY_CSTR;
            else
               *comment += CRLF;
            *comment += descr;
         }
      }

      ppf->Release(); 
      psl->Release();

   } 
   return TRUE; 
} 

//--------------------------------------------------------------------------

BOOL GetFileIcons(LPCTSTR fileName, HICON *largeIcon, HICON *smallIcon)
{
   //SHFILEINFO inf;
	SHFILEINFOW sfi = { 0 };
	SHGetFileInfo(fileName, -1, &sfi, sizeof(sfi), SHGFI_SYSICONINDEX);
	HIMAGELIST* imageList;
   if (largeIcon)
   {
      //SHGetFileInfo(fileName, 0, &inf, sizeof(inf), SHGFI_ICON | SHGFI_LARGEICON);
      //*largeIcon = inf.hIcon;
	   
	   HRESULT hResult = SHGetImageList(SHIL_EXTRALARGE, IID_IImageList, (void**)&imageList);
	   if (hResult == S_OK) {
		   HICON hIcon;
		   hResult = ((IImageList*)imageList)->GetIcon(sfi.iIcon, ILD_TRANSPARENT, &hIcon);
		   *largeIcon = hIcon;
	   }
   }
   if (smallIcon)
   {
      //SHGetFileInfo(fileName, 0, &inf, sizeof(inf), SHGFI_ICON | SHGFI_SMALLICON);
      //*smallIcon = inf.hIcon;

	   HRESULT hResult = SHGetImageList(SHIL_LARGE, IID_IImageList, (void**)&imageList);
	   if (hResult == S_OK) {
		   HICON hIcon;
		   hResult = ((IImageList*)imageList)->GetIcon(sfi.iIcon, ILD_TRANSPARENT, &hIcon);
		   *smallIcon = hIcon;
	   }
   }
   return TRUE;
}

//--------------------------------------------------------------------------

CString GetTrueTarget(LPCTSTR path)
{
   if (!IS_SHORTCUT(path))
      return path;
   CString target;
   GetShortcutInfo(path, target);
   return target;
}

//--------------------------------------------------------------------------

typedef struct {
   BOOL useLarge;             // Use the large icon in the menu
   HICON smallIcon;           // The small icon
   HICON largeIcon;           // The large icon
   PVOID extra;               // Possible extra data
} tItemData, *pItemData;

void InitPopupMenu(HMENU hMenu, BOOL useMenuCom)
{
   // Force update of a popup menu not connected to any window
   HMENU dummy = CreateMenu();
   AppendMenu(dummy, MF_POPUP, (UINT)hMenu, EMPTY_STR);
   RemoveMenu(dummy, 0, MF_BYPOSITION);
   DestroyMenu(dummy);

   if (useMenuCom)
   {
      // Setup to get WM_MENUCOMMAND instead of WM_COMMAND
      MENUINFO menuInf;
      ZeroMemory(&menuInf, sizeof(menuInf));
      menuInf.cbSize = sizeof(menuInf);
      menuInf.fMask = MIM_STYLE | MIM_APPLYTOSUBMENUS;
      menuInf.dwStyle = MNS_NOTIFYBYPOS;
      SetMenuInfo(hMenu, &menuInf);
   }
}

//--------------------------------------------------------------------------

BOOL AddMenuItem(HMENU hMenu, LPCTSTR text, HMENU hSubMenu)
{
   // Create a new menu item or sub menu
   MENUITEMINFO menuInf;
   ZeroMemory(&menuInf, sizeof(menuInf));
   menuInf.cbSize = sizeof(menuInf);
   menuInf.fMask = MIIM_FTYPE | MIIM_SUBMENU | MIIM_STRING;
   menuInf.fType = MFT_STRING;
   menuInf.dwTypeData = (LPTSTR)text;
   menuInf.cch = STR_LEN(text);
   menuInf.hSubMenu = hSubMenu;

   return InsertMenuItem(hMenu, -1, TRUE, &menuInf);
}

//--------------------------------------------------------------------------

BOOL SetMenuItemData(HMENU hMenu, INT itemID, HICON smallIcon, HICON largeIcon, BOOL useLarge, PVOID extra)
{
   // Set menu icons and possible extra data handle
   MENUITEMINFO menuInf;
   ZeroMemory(&menuInf, sizeof(menuInf));
   menuInf.cbSize = sizeof(menuInf);
   menuInf.fMask = MIIM_DATA | MIIM_BITMAP;
   menuInf.hbmpItem = HBMMENU_CALLBACK; // For icon draw
   
   pItemData pData = new tItemData;
   pData->useLarge = largeIcon != NULL && useLarge;
   pData->smallIcon = smallIcon;
   pData->largeIcon = largeIcon;
   pData->extra = extra;
   menuInf.dwItemData = (ULONG_PTR)pData;

   BOOL ok = SetMenuItemInfo(hMenu, abs(itemID), itemID < 1, &menuInf);
   return ok;
}

//--------------------------------------------------------------------------

PVOID GetMenuItemData(HMENU hMenu, INT itemID)
{
   // Get the extra data handle of a menu item
   MENUITEMINFO menuInf;
   ZeroMemory(&menuInf, sizeof(menuInf));
   menuInf.cbSize = sizeof(menuInf);
   menuInf.fMask = MIIM_DATA;
   BOOL ok = GetMenuItemInfo(hMenu, itemID, TRUE, &menuInf);
   return (ok ? ((pItemData)(menuInf.dwItemData))->extra : NULL);
}

//--------------------------------------------------------------------------

BOOL gForceLarge = FALSE;

BOOL SetLargeMenus(BOOL on)
{
   // Force all entries to large size
   gForceLarge = on;
   return TRUE;
}

//--------------------------------------------------------------------------

#define ICON_SIZE() (large ? 32 : 16)

BOOL HandleMenuItemIconMessage(UINT message, LPARAM lParam)
{
   pItemData pData;
   BOOL large;
   if (message == WM_MEASUREITEM)
   {
      // Measure the specified item
      MEASUREITEMSTRUCT* pMeas = (MEASUREITEMSTRUCT*)lParam;
      pData = (pItemData)pMeas->itemData;
      large = (gForceLarge || pData->useLarge) && (pData->largeIcon != NULL);
      DWORD size = ICON_SIZE();
      pMeas->itemWidth += large ? 24 : 5;
      if (pMeas->itemHeight < size)
          pMeas->itemHeight = size;
   }
   if (message == WM_DRAWITEM)
   {
      // Draw the item icon
      DRAWITEMSTRUCT* pDis = (DRAWITEMSTRUCT*)lParam;
      if (!pDis || pDis->CtlType != ODT_MENU || !pDis->itemData)
         return FALSE;

      pData = (pItemData)pDis->itemData;
      large = (gForceLarge || pData->useLarge) && (pData->largeIcon != NULL);
      DWORD size = ICON_SIZE();
      DrawIconEx(pDis->hDC,
         pDis->rcItem.left - (large ? size/2-2 : size-1),
         pDis->rcItem.top + (pDis->rcItem.bottom - pDis->rcItem.top - size) / 2,
         (HICON)(large ? pData->largeIcon : pData->smallIcon), 
         size, size, 0, NULL, DI_NORMAL);
   }
   else
      return FALSE;

   return TRUE;
}

//--------------------------------------------------------------------------

BOOL DestroyMenuData(HMENU hMenu, BOOL destroyMenu)
{
   INT i;
   for (i = 0; i < GetMenuItemCount(hMenu); i++)
   {
      MENUITEMINFO menuInf;
      ZeroMemory(&menuInf, sizeof(menuInf));
      menuInf.cbSize = sizeof(menuInf);
      menuInf.fMask = MIIM_DATA | MIIM_SUBMENU;
      if (GetMenuItemInfo(hMenu, i, TRUE, &menuInf))
      {
         pItemData pData = (pItemData)menuInf.dwItemData;
         DestroyIcon(pData->largeIcon);
         DestroyIcon(pData->smallIcon);
         if (pData->extra)
            delete pData->extra;
         delete pData;
      }
      if (menuInf.hSubMenu)
         DestroyMenuData(menuInf.hSubMenu);
   }
   if (destroyMenu)
	   DestroyMenu(hMenu);
   return TRUE;
}

//--------------------------------------------------------------------------

BOOL DoRun(LPCTSTR com, LPCTSTR params, HWND hWnd, WORD showType, LPCTSTR verb)
{
   // Start the specified command in a new process
   SHELLEXECUTEINFO shExecInfo;

   ZeroMemory(&shExecInfo, sizeof(shExecInfo));
   shExecInfo.cbSize = sizeof(shExecInfo);

   shExecInfo.hwnd = hWnd;
   shExecInfo.lpVerb = verb;
   shExecInfo.lpFile = com;
   shExecInfo.lpParameters = params;
   shExecInfo.nShow = showType;
   shExecInfo.fMask = SEE_MASK_INVOKEIDLIST;

   return ShellExecuteEx(&shExecInfo);
}

//--------------------------------------------------------------------------

CString GetExplorerDir(LPCTSTR keyName, BOOL allUsers = FALSE)
{
   CString dir = GetRegVal(CString(_T("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\")) + keyName, 
                           EMPTY_STR, allUsers ? HKEY_LOCAL_MACHINE : HKEY_CURRENT_USER);
   ExpEnvVars(dir);
   APPEND_BS(dir);
   return dir;
}

//--------------------------------------------------------------------------

CString GetQuickLaunchDir()
{
   // Get the path of the quick-launch directory
   CString dir = GetExplorerDir(_T("User Shell Folders\\AppData"));
   if (!dir.IsEmpty())
      dir += _T("Microsoft\\Internet Explorer\\Quick Launch\\");
   return dir;
}

//--------------------------------------------------------------------------

CString GetAutoStartDir(BOOL allUsers)
{
   return GetExplorerDir(allUsers ? _T("Shell Folders\\Common Startup") : _T("User Shell Folders\\Startup"), 
                         allUsers);
}

//--------------------------------------------------------------------------

CString GetStartMenuDir(BOOL allUsers)
{
   return GetExplorerDir(allUsers ? _T("Shell Folders\\Common Start Menu") : _T("User Shell Folders\\Start Menu"), 
                         allUsers);
}

//--------------------------------------------------------------------------

// Directory change data structure
typedef struct {
   HANDLE hChange;
   HANDLE hThread;
   HWND hWnd;
   DWORD messageID;
} tWatchInfo, *pWatchInfo;

DWORD WINAPI WatchThreadProc(LPVOID param)
{
   pWatchInfo pInf = (pWatchInfo)param;
   BOOL notified = FALSE;
   while (TRUE)
   {
      // Don't send notifications to the application until possible bursts has calmed down
      if (WaitForSingleObject(pInf->hChange, 1000) == WAIT_OBJECT_0)
      {
         notified = TRUE;
         // Wait for next
         FindNextChangeNotification(pInf->hChange);
      }
      else if (notified)
      {
         // Timer event, now report the pending change
         PostMessage(pInf->hWnd, pInf->messageID, NULL, NULL);
         notified = FALSE;
      }
   }
   return 0;
}

//--------------------------------------------------------------------------

BOOL StartWatchDir(LPCTSTR dir, HWND hWnd, DWORD messageID, LPVOID watchHand)
{
   // Watch for changes in the specified directory and report with the specified message id to the specified window
   pWatchInfo pInf = new tWatchInfo;
   BOOL ok =  pInf != NULL;
   if (ok)
   {
      pInf->hWnd = hWnd;
      pInf->messageID = messageID;
   }

   ok = ok && (pInf->hChange = FindFirstChangeNotification(dir, TRUE, FILE_NOTIFY_CHANGE_FILE_NAME | FILE_NOTIFY_CHANGE_DIR_NAME | FILE_NOTIFY_CHANGE_LAST_WRITE)) != INVALID_HANDLE_VALUE;
   ok = ok && (pInf->hThread = CreateThread(NULL, 0, WatchThreadProc, pInf, 0, NULL)) != NULL;

   if (ok && watchHand)
      watchHand = (PVOID)pInf;
   else if (!ok)
      delete pInf;

   return ok;
}

//--------------------------------------------------------------------------

BOOL EndWatchDir(LPVOID watchHand)
{
   // Stop directory watch
   BOOL ok = FALSE;
   pWatchInfo pInf = (pWatchInfo)watchHand;
   if (pInf)
   {
      ok = TerminateThread(pInf->hThread, 0);
      FindCloseChangeNotification(pInf->hChange);
   }
   return ok;
}

//--------------------------------------------------------------------------

CString LoadFormatResString(UINT ID, ...)
{
   // Load a resource string and format with the specified parameters
   CString resStr, formStr;
   resStr.LoadString(ID);
   va_list args;
   va_start(args, ID);

   formStr.FormatV(resStr, args);
   va_end(args);

   return formStr;
}

//--------------------------------------------------------------------------

CString ErrorString(DWORD err)
{
   // Return a system error message string
   CString errStr = EMPTY_STR;
   LPTSTR s;
   if (FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
                    NULL, err, 0, (LPTSTR)&s, 0, NULL))
   {
      errStr = s;
      LocalFree(s);
   }

   return errStr;
}

//--------------------------------------------------------------------------

BOOL ShowMessage(LPCTSTR message, tMessSeverity sev, HWND hWnd, LPCTSTR header)
{
   // Show a message with specified severity notification
   UINT type = MB_OK;
   if (sev == eInfo)
      type |= MB_ICONINFORMATION;
   else if (sev == eWarning)
      type |= MB_ICONWARNING;
   else
      type |= MB_ICONERROR;
   CString headStr = header ? header : LoadFormatResString(1000+sev);
   MessageBox(hWnd, message, headStr, type);
   if (sev == eFatal)
      exit(1);
   return FALSE;
}

//--------------------------------------------------------------------------

void PositionDialog(HWND hDlg)
{
   // Place a dialog box at an appropriate position next to the current cursor position
   POINT pos;
   RECT screenRect, windRect;
   SystemParametersInfo(SPI_GETWORKAREA, 0, &screenRect, 0);
   GetWindowRect(hDlg, &windRect);
   GetCursorPos(&pos);
   pos.x = min(pos.x, screenRect.right-(windRect.right-windRect.left));
   pos.y = min(pos.y, screenRect.bottom-(windRect.bottom-windRect.top));
   SetWindowPos(hDlg, NULL, pos.x, pos.y, 0, 0, SWP_NOSIZE | SWP_NOOWNERZORDER);
}

//--------------------------------------------------------------------------

UINT_PTR CALLBACK PositionDlgHook(HWND hDlg,
                      UINT message,
                      UINT wParam,
                      LONG lParam)
{
   if (message == WM_NOTIFY && ((LPOFNOTIFY)lParam)->hdr.code == CDN_INITDONE)
      PositionDialog(GetParent(hDlg));
   return 0;
}

//--------------------------------------------------------------------------

void CenterDlg(HWND hDlg)
{
   // Place a dialog at a suitable center position of the screen
   RECT rc;
   GetWindowRect(hDlg, &rc);
   SetWindowPos(hDlg, NULL, 
         ((GetSystemMetrics(SM_CXSCREEN)-(rc.right-rc.left))/2),
         ((GetSystemMetrics(SM_CYSCREEN)-(rc.bottom-rc.top))/3),
            0, 0, SWP_NOSIZE | SWP_NOACTIVATE);
}

//--------------------------------------------------------------------------

UINT_PTR CALLBACK CenterDlgHook(HWND hDlg,
                      UINT message,
                      UINT wParam,
                      LONG lParam)
{
   if (message == WM_NOTIFY && ((LPOFNOTIFY)lParam)->hdr.code == CDN_INITDONE)
      CenterDlg(GetParent(hDlg));
   return 0;
}

//--------------------------------------------------------------------------

CString PromptFileName(LPCTSTR startFileName, HWND fatherWind, BOOL save, DWORD typeListID, BOOL center)
{
   OPENFILENAME   ofn;
   BOOL           ok;
   CString        fileName(startFileName);
   CString        typeList = LoadFormatResString(typeListID);

   DWORD len = typeList.GetLength();
   if (!len)
      return EMPTY_CSTR;

   TCHAR rep = typeList[len-1];
   PTCHAR pTypeBuff = typeList.GetBuffer(len+1);
   for (DWORD i = 0; i < len; i++)
      if (*(pTypeBuff+i) == rep)
         *(pTypeBuff+i) = 0;

   // Set up OPENFILENAME structure 
   memset(&ofn, 0, sizeof(OPENFILENAME));
   ofn.lStructSize = sizeof(OPENFILENAME);
   ofn.hwndOwner = fatherWind;
   ofn.hInstance = CURR_INSTANCE;
   ofn.lpstrFilter = pTypeBuff;
   ofn.nFilterIndex = 0;
   ofn.lpstrFile = fileName.GetBuffer(MAX_PATH);
   ofn.nMaxFile = MAX_PATH;
   ofn.lpfnHook = (LPOFNHOOKPROC)(center ? CenterDlgHook : PositionDlgHook);
   ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_NOCHANGEDIR |
               OFN_ENABLEHOOK | OFN_HIDEREADONLY | OFN_EXPLORER;
   // Do the dialog
   ok = save ? GetSaveFileName(&ofn) : GetOpenFileName(&ofn);
   fileName.ReleaseBuffer();
   typeList.ReleaseBuffer();

   return (ok ? fileName : EMPTY_CSTR);
}

//--------------------------------------------------------------------------

BOOL IsWinPE()
{
   HKEY hKey;
   // Check for Windows PE registry key
   if (RegOpenKeyEx(HKEY_LOCAL_MACHINE,
               _T("system\\currentcontrolset\\control\\minint"),
               0, KEY_ALL_ACCESS, &hKey) != ERROR_SUCCESS)
      return FALSE;
   RegCloseKey(hKey);
   return TRUE;
}

//--------------------------------------------------------------------------

BOOL RunDialog(HWND hWnd)
{
	// Load and start undocumented function from shell32
	typedef VOID 
	( WINAPI* LPRUNFILEDLGPROC )(
		HWND hWndOwner, 
		HICON	hIcon, 
		LPCTSTR lpszDir,
		LPCTSTR lpszTitle,
		LPCTSTR lpszDesc,
		DWORD dwFlags
	);

	BOOL ok = FALSE;
	HMODULE hShell32 = LoadLibrary(_T("shell32.dll"));

	if (hShell32)
	{
		LPRUNFILEDLGPROC runDlg = (LPRUNFILEDLGPROC)GetProcAddress(hShell32, (LPCSTR)61);
		if (runDlg)
		{
			runDlg(hWnd, NULL, NULL, NULL, NULL, 0);
			ok = TRUE;
		}
		FreeLibrary(hShell32);
	}
	return ok;
}



