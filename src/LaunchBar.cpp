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

#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <windows.h>
#include <shlobj.h>
#include <commdlg.h>
#include <tchar.h>
#include <list>
#include <map>

#include <stdio.h>
#include <time.h>

#include "resource.h"
#include "utils.h"

#define PROG_NAME _T("LaunchBar")
#define VERSION_STR _T("3.0.0")
#define DATE_STR _T("2012-05-13")

// Window classes
#define BUTTON_CLASS_NAME _T("ICON_BUTTON")
#define WIND_CLASS_NAME PROG_NAME

// Global variables
HWND      gMainWindow;              // The main toolbar window
RECT      gWindowRect;              // Main window rectangle in screen coordinates
POINT     gWindowSize;              // Window size for convenience
CString   gMainDir(EMPTY_STR);      // Directory to watch and mirror
CString   gAppPath;                 // The path to the executable
BOOL      gConfigFileUsed = FALSE;  // Set when configuration file used instead of a directory
HMENU     gMainPopupMenu,           // Main window popup menu
          gFilePopupMenu,           // Popup menu for file buttons or menu entries
          gDirPopupMenu;            // Popup menu for directory buttons or menu entries
BOOL      gUseReg = TRUE;           // Use registry for saving preferences
CString   gStartMenuDir,            // Holds the path to the machine start menu
          gStartMenuLink;           // Possible start menu shortcut, handled specially

// Foward declarations
BOOL ReadPrefs();
BOOL SavePrefs();

// Button individual information
typedef struct {
   CString command;     // Command to execute
   CString params;      // Command parameters
   HMENU hMenu;         // Used for possible sub menu for folder buttons
   HICON hIcon;         // Large icon
   HICON hSmallIcon;    // Small icon
   HWND  hToolTip;      // Tooltip window
   WORD  showType;      // Window type for the application to launch
} tCommandInfo, *pCommandInfo;

// Button information is stored in the window long data field
#define GET_COM_INFO(hWnd) (pCommandInfo)GetWindowLongPtr(hWnd, GWLP_USERDATA)

// Message ID for directory change event
#define WM_USER_DIR_CHANGED WM_USER+1

// Maximum number of buttons allowed
#define MAX_BUTTONS 100

// Button list
struct {
   DWORD cnt;
   HWND list[MAX_BUTTONS];
} gButtons = {0};

// Layout data
BOOL gLargeIcons = FALSE; // Use large buttons
BOOL gLargeMenus = FALSE; // Use large folder menu entries
BOOL gOnTop = FALSE;      // Toolbar always on top
DWORD gAutoHide = 0;      // Hide toolbar when loosing focus
BOOL gHidden = FALSE;     // Visibility state
DWORD gLocation = 3;      // Corner position (1 = left, 2 = top, 3 = right, 4 = bottom)

DWORD xOffset, yOffset;  // Offset from window border to button
DWORD xInc, yInc;        // Increment between buttons

#define ICON_SIZE (gLargeIcons ? 32 : 16) // Standard icon sizes
#define ICON_OFF (gLargeIcons ? 5 : 3)    // Offset between the icon frame and the toolbar window frame
#define WIND_OFF 12                       // Offset between toolbar window edge and first/last button
#define MARK_RADIUS 3                     // Radius of a circle used to mark the toolbar window own space used for the popup menu

#define GET_ICON(id) LoadIcon(CURR_INSTANCE, MAKEINTRESOURCE(id))

BOOL PositionWindow()
{
   // Place the toolbar window at its normal position
   return SetWindowPos(gMainWindow, gOnTop ? HWND_TOPMOST : HWND_NOTOPMOST, gWindowRect.left, gWindowRect.top, 
                       gWindowSize.x, gWindowSize.y, SWP_SHOWWINDOW);
}

//--------------------------------------------------------------------------

INT_PTR CALLBACK PromptDirectory(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
   // Prompt for name and create a corresponding directory
   static CString root = EMPTY_CSTR;
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	   case WM_INITDIALOG:
         PositionDialog(hDlg);
         root = *((CString*)lParam);
         APPEND_BS(root);
		   return (INT_PTR)TRUE;
         break;


      case WM_COMMAND:
         switch (LOWORD(wParam))
         {
            case IDOK:
               {
                  TCHAR name[_MAX_PATH];
                  GetWindowText(GetDlgItem(hDlg, IDC_ANSWER), name, STR_SIZE(name));
                  CString dir = root + name;
                  if (FileExists(dir))
                  {
                     ShowMessage(LoadFormatResString(IDS_DIR_EXIST, dir), eWarning);
                     return FALSE;
                  }
                  else if (!CreateDirectory(dir, NULL))
                  {
                     ShowMessage(LoadFormatResString(IDS_DIR_FAIL, dir), eError);
                     return FALSE;
                  }
               }

               EndDialog(hDlg, TRUE);
               return (INT_PTR)TRUE;

            case IDCANCEL:
               EndDialog(hDlg, FALSE);
               return (INT_PTR)TRUE;
               break;

            case IDC_ANSWER:
               EnableWindow(GetDlgItem(hDlg, IDOK), GetWindowTextLength(GetDlgItem(hDlg, IDC_ANSWER)) > 0);
               break;

             default:
               break;
        }
	}
	return (INT_PTR)FALSE;
}

//--------------------------------------------------------------------------

BOOL File2Shortcut(CString& fileName, CString dir)
{
   // Create a shortcut to a file unless it already is a shortcut
   if (gConfigFileUsed)
      return FALSE;

   BOOL ok = FALSE;
   CString shortcut;
   APPEND_BS(dir);
   if (IS_SHORTCUT(fileName))
   {
      // Copy the shortcut to the specified directory
      shortcut = dir + GetFileNameComp(fileName, eFcName | eFcType);
      ok = CopyFile(fileName, shortcut, FALSE);
   }
   else
   {
      // Create a new shortcut to the file in the specified directory
      shortcut = dir + GetFileNameComp(fileName, eFcName) + SHORTCUT_EXT;
      ok = CreateShortcut(shortcut, fileName);
   }
   if (ok)
      fileName = shortcut;
   return ok;
}

//--------------------------------------------------------------------------

BOOL EnableAutoHide(BOOL on)
{
   const UINT id = 1;
   if (!gAutoHide)
      return FALSE;
   if (on)
      // Start timer for the hide monitoring
      SetTimer(gMainWindow, id, 1000, NULL);
   else
      // Stop hiding
      KillTimer(gMainWindow, id);
   return TRUE;
}

//--------------------------------------------------------------------------

BOOL AutoHide(BOOL hide)
{
   // Hide the main window by making it thin
   const int frameSize = 4;
   if (!gAutoHide)
      return FALSE;
   gHidden = hide;
   if (hide)
   {
      int x = gWindowRect.left, 
          y = gWindowRect.top,
          h = gWindowSize.y,
          w = gWindowSize.x;
      switch (gLocation)
      {
         case 1:
            x = gWindowRect.left-1;
            w = frameSize;
            break;
         case 2:
            y = gWindowRect.top-1;
            h = frameSize;
            break;
         case 3:
            x = gWindowRect.right-frameSize+1;
            w = frameSize;
            break;
         case 4:
            y = gWindowRect.bottom-frameSize;
            h = frameSize;
            break;
      }
      return SetWindowPos(gMainWindow, HWND_TOPMOST, x, y, w, h, SWP_SHOWWINDOW);
   }
   else
      // Restore
      return PositionWindow();
}

//--------------------------------------------------------------------------

BOOL SetupLayout()
{
   // Position the buttons according to their placement in the list
   DWORD xButtonPos, yButtonPos;
   DWORD xButtonSize, yButtonSize;
   DWORD inc;
   RECT screenRect;
   SystemParametersInfo(SPI_GETWORKAREA, 0, &screenRect, 0);

   xButtonSize = ICON_SIZE + ICON_OFF*2;
   yButtonSize = xButtonSize;
   xOffset = 2;
   yOffset = xOffset;
   inc = (gLargeIcons ? 2 : 0);

   IN_RANGE(gLocation, 1, 4);

   if (gLocation == 1)
   {
      // Left
      xInc = 0;
      yInc = yButtonSize + inc;
      xButtonPos = xOffset;
      yButtonPos = yOffset + WIND_OFF;
      gWindowRect.left = screenRect.left;
      gWindowRect.right = gWindowRect.left + xOffset*2 + xButtonSize + 2;
      gWindowRect.top = screenRect.top;
      gWindowRect.bottom = gWindowRect.top + yButtonPos*2 + gButtons.cnt*yInc;
   }
   else if (gLocation == 2)
   {
      // Top
      xInc = xButtonSize + inc;
      yInc = 0;
      xButtonPos = xOffset + WIND_OFF;
      yButtonPos = yOffset;
      gWindowRect.left = screenRect.left;
      gWindowRect.right = gWindowRect.left + xButtonPos*2 + gButtons.cnt*xInc;
      gWindowRect.top = screenRect.top;
      gWindowRect.bottom = gWindowRect.top + yOffset*2 + yButtonSize + 2;
   }
   else if (gLocation == 3)
   {
      // Right
      xInc = 0;
      yInc = yButtonSize + inc;
      xButtonPos = xOffset+1;
      yButtonPos = yOffset + WIND_OFF;
      gWindowRect.right = screenRect.right;
      gWindowRect.left = gWindowRect.right - xOffset*2 - xButtonSize - 2;
      gWindowRect.top = screenRect.top;
      gWindowRect.bottom = gWindowRect.top + yButtonPos*2 + gButtons.cnt*yInc;
   }
   else
   {
      // Bottom
      xInc = xButtonSize + inc;
      yInc = 0;
      xButtonPos = xOffset + WIND_OFF;
      yButtonPos = yOffset;
      gWindowRect.left = screenRect.left;
      gWindowRect.right = gWindowRect.left + xButtonPos*2 + gButtons.cnt*xInc;
      gWindowRect.bottom = screenRect.bottom;
      gWindowRect.top = gWindowRect.bottom - yOffset*2 - yButtonSize - 2;
   }

   DWORD i;
   for (i = 0; i < gButtons.cnt; i++)
   {
      // Set position
      SetWindowPos(gButtons.list[i], NULL, xButtonPos, yButtonPos, xButtonSize, yButtonSize, SWP_SHOWWINDOW | SWP_NOOWNERZORDER);
      xButtonPos += xInc;
      yButtonPos += yInc;
      InvalidateRect(gButtons.list[i], NULL, TRUE);
   }

   gWindowSize.x = gWindowRect.right-gWindowRect.left;
   gWindowSize.y = gWindowRect.bottom-gWindowRect.top;
   PositionWindow();
   // Make sure that the window contents is updated
   InvalidateRect(gMainWindow, NULL, TRUE);

   return TRUE;
}

//--------------------------------------------------------------------------

LONG Pos2Index(LONG x, LONG y)
{
   // Get the list index for the provided window position
   LONG ind = -1;
   switch (gLocation)
   {
      case 1:
      case 3:
         ind = round((y - (LONG)yOffset - WIND_OFF)/(double)yInc);
         break;

      case 2:
      case 4:
         ind = round((x - (LONG)xOffset - WIND_OFF)/(double)xInc);
   }

   return (DWORD)max(min(ind, (LONG)gButtons.cnt), 0);
}

//--------------------------------------------------------------------------

LONG Hwnd2Index(HWND hWnd)
{
   // Get the list index for the provided window handle if available, otherwise -1
   DWORD i;
   for (i = 0; i < gButtons.cnt; i++)
      if (gButtons.list[i] == hWnd)
         return i;
   return -1;
}

//--------------------------------------------------------------------------

LONG Com2Index(LPCTSTR com)
{
   // Get the list index for the provided command if available, otherwise -1
   DWORD i;
   for (i = 0; i < gButtons.cnt; i++)
   {
      pCommandInfo pCom = GET_COM_INFO(gButtons.list[i]);
      if (pCom && pCom->command.CompareNoCase(com) == 0)
         return i;
   }
   return -1;
}

//--------------------------------------------------------------------------

BOOL HandleDragAndDrop(CString& fileName)
{
   // Handle a file name dropped on the application window
   return File2Shortcut(fileName, gMainDir);
}

//--------------------------------------------------------------------------

BOOL DirSort(CString first, CString second)
{
   // Sort directories first
   BOOL d1 = IsDir(first),
        d2 = IsDir(second);
   CString n1 = GetFileNameComp(first, eFcName),
           n2 = GetFileNameComp(second, eFcName);
   if (d1 && !d2)
      return TRUE;
   else if (!d1 && d2)
      return FALSE;
   else
      // Case insensitive sorting when equal
      return n1.CompareNoCase(n2) < 0;
}

//--------------------------------------------------------------------------

HMENU AddDirMenu(CString dir, BOOL doInit, CString altDir)
{
   // Create popup menu with entries corresponding to all the shorcuts in a
   // specified directory and all its sub directories.
   HMENU hMenu = CreateMenu();

   // Find and sort the files and directories
   CString path = dir + _T("\\*"),
           fileName = EMPTY_CSTR;
   HANDLE ref;
   std::list<CString> fileList;        // List for sorting
   std::map<CString, CString> fileMap; // Hash used for possible multiple directories
   // Primary files
   while (FindFiles(path, fileName, &ref))
   {
      fileList.push_back(fileName);
      fileMap[GetFileNameComp(fileName, eFcName)] = EMPTY_CSTR;
   }
   // Special handling for the start menu which is present in two different locations (machine and user)
   if (altDir.IsEmpty() && dir == gStartMenuDir)
      // Merge the user start menu as well
      altDir = GetStartMenuDir(FALSE);

   if (!altDir.IsEmpty())
   {
      // Merge additional directory
      fileName = EMPTY_CSTR;
      APPEND_BS(altDir);
      path = altDir + _T("*");
      while (FindFiles(path, fileName, &ref))
      {
         CString name = GetFileNameComp(fileName, eFcName);
         if (fileMap.find(name) == fileMap.end())
            // New
            fileList.push_back(fileName);
         else if (IsDir(fileName))
            // Existing, merge if directory
            fileMap[name] = fileName;
      }
   }
   // Sort the entries found
   fileList.sort(DirSort);

   // Add the menu(s)
   std::list<CString>::iterator it;
   for (it=fileList.begin(); it!=fileList.end(); ++it)
   {
      CString fileName = *it;
      if (FileIsHidden(fileName))
         // Ignore hidden files
         continue;

      HICON smallIcon, largeIcon;
      CString target;
      CString itemName = GetFileNameComp(fileName, eFcName);
      HMENU subMenu = NULL;

      if (IS_SHORTCUT(fileName))
      {
         GetShortcutInfo(fileName, target, &largeIcon, &smallIcon);
         if (!FileExists(target))
            // Ignore shortcuts to missing targets
            continue;
         if (IsDir(target))
            // Use the target folder
            fileName = target;
      }
      else
         GetFileIcons(fileName, &largeIcon, &smallIcon);

      if (IsDir(fileName))
         // Recursively handle sub directories
         subMenu = AddDirMenu(fileName, FALSE, fileMap[itemName]);

      // Insert this item in the menu
      AddMenuItem(hMenu, itemName, subMenu);
      CString *link = new CString(fileName);
      SetMenuItemData(hMenu, -(GetMenuItemCount(hMenu)-1), smallIcon, largeIcon, FALSE, link);
      if (subMenu)
      {
         // Insert the name as menu info as well for sub menus.
         // Needed to handle context menu selection on cascading sub menus.
         MENUINFO menuInf;
         ZeroMemory(&menuInf, sizeof(menuInf));
         menuInf.cbSize = sizeof(menuInf);
         menuInf.fMask = MIM_MENUDATA;
         menuInf.dwMenuData = (ULONG_PTR)link;
         SetMenuInfo(subMenu, &menuInf);
      }
   }
   if (doInit)
      InitPopupMenu(hMenu, TRUE);

   return hMenu;

}

//--------------------------------------------------------------------------

BOOL AddNewButton(LPCTSTR command, 
                  DWORD pos = gButtons.cnt,
                  LPCTSTR params = EMPTY_CSTR, 
                  CString iconFile = EMPTY_CSTR, DWORD iconInd = 0, 
                  CString toolTip = EMPTY_CSTR, WORD showType = SW_SHOWNORMAL)
{
   if (!FileExists(command) || FileIsHidden(command) || gButtons.cnt >= MAX_BUTTONS)
      // Missing or too many already
      return FALSE;

   HWND hWndButton = CreateWindow(BUTTON_CLASS_NAME, EMPTY_STR,  WS_CHILD,
                                  0, 0, 100, 100, 
                                  gMainWindow, (HMENU)NULL, CURR_INSTANCE, NULL);
   if (!hWndButton)
      return FALSE;

   pCommandInfo pCom = new tCommandInfo;
   if (!pCom)
      return FALSE;

   // Save the information structure in window user data field
   SetWindowLongPtr(hWndButton, GWLP_USERDATA, (LONG)pCom);

   pCom->command = command;
   pCom->params = params;

   CString target = command;
   if (IS_SHORTCUT(command))
   {
      // Get shortcut icons and tooltip
      GetShortcutInfo(pCom->command, target, &pCom->hIcon, &pCom->hSmallIcon, &toolTip);
   }
   else
   {
      if (toolTip.IsEmpty())
         toolTip = GetFileNameComp(command, eFcName | eFcType);
      if (iconFile.IsEmpty())
         // Icon file not specified, use the command itself
         GetFileIcons(pCom->command, &pCom->hIcon, &pCom->hSmallIcon);
      else
         ExtractIconEx(iconFile, iconInd, &pCom->hIcon, &pCom->hSmallIcon, 1);
   }

   pCom->hToolTip = CreateTooltip(hWndButton, toolTip);
   pCom->showType = showType;
   pCom->hMenu = NULL;

   if (IsDir(target))
      // Create a popup menu for this button
      pCom->hMenu = AddDirMenu(target, TRUE, EMPTY_CSTR);

   // Want drag and drop of files for the buttons
   DragAcceptFiles(hWndButton, TRUE);

   // Insert in the list
   for (DWORD i = gButtons.cnt; i > pos; i--)
      gButtons.list[i] = gButtons.list[i-1];
   gButtons.list[pos] = hWndButton;

   gButtons.cnt++;

   return TRUE;
}

//--------------------------------------------------------------------------

// Brushes for highlighting of the buttons
HBRUSH gBgBrush = CreateSolidBrush(RGB(211, 218, 237));        // Normal background
HBRUSH gPushBrush = CreateSolidBrush(RGB(150, 150, 150));      // Button pushed
HPEN gLightPen = CreatePen(PS_SOLID, 0, RGB(255, 255, 255));   // Normal button light side
HPEN gShadowPen = CreatePen(PS_SOLID, 0, RGB(127, 127, 127));  // Normal button shadow side
HPEN gBgPen = CreatePen(PS_SOLID, 0, RGB(211, 218, 237));      // Normal background pen

enum eDrawType {eNormal, eHighlight, ePushed};
void DrawButton(HWND hWnd, HDC orgDC = NULL, eDrawType type = eNormal)
{
   // Draw a button in the specified way
   RECT r;
   GetClientRect(hWnd, &r);
   HDC hDC = orgDC ? orgDC : GetWindowDC(hWnd);
   SelectObject(hDC, GetStockObject(NULL_PEN));
   if (type == ePushed)
      SelectObject(hDC, gPushBrush);
   else
      SelectObject(hDC, gBgBrush);
   Rectangle(hDC, r.left, r.top, r.right, r.bottom);
   if (type != eNormal)
   {
      // Light side
      SelectObject(hDC, type == ePushed ? gShadowPen : gLightPen);
      MoveToEx(hDC, r.left, r.bottom, NULL);
      LineTo(hDC, r.left, r.top);
      LineTo(hDC, r.right-1, r.top);
      // Dark side
      SelectObject(hDC, type != ePushed ? gShadowPen : gLightPen);
      LineTo(hDC, r.right-1, r.bottom-1);
      LineTo(hDC, r.left, r.bottom-1);
   }

   pCommandInfo pCom = GET_COM_INFO(hWnd);
   if (pCom)
   {
      DrawIconEx(hDC, ICON_OFF, ICON_OFF, gLargeIcons ? pCom->hIcon : pCom->hSmallIcon, 
                 ICON_SIZE, ICON_SIZE, 0, NULL, DI_NORMAL);
      if (pCom->hMenu)
      {
         // Draw arrow in order to indicate the popup menu
         BeginPath(hDC);
         INT w = 4,
             h = 8,
             c = (r.bottom+r.top)/2;
         MoveToEx(hDC, r.right, c, NULL);
         LineTo(hDC, r.right-w, c-h/2);
         LineTo(hDC, r.right-w, c+h/2);
         LineTo(hDC, r.right, c);
         EndPath(hDC);
         SelectObject(hDC, GetStockObject(BLACK_BRUSH));
         FillPath(hDC);
         SelectObject(hDC, gBgPen);
         MoveToEx(hDC, r.right-w, c-h/2, NULL);
         LineTo(hDC, r.right-w, c+h/2);
      }
   }
   if (!orgDC)
      ReleaseDC(hWnd, hDC);
}

LRESULT APIENTRY ButtonWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) 
{
   // Window callback procedure for the button class
   pCommandInfo pCom;
   static HWND tracker = NULL;
   static BOOL dragging = FALSE;
   // The actual target for context menu operations is stored here
   static CString contMenuCom = EMPTY_CSTR;

   // Context menu setup and tracking
   static BOOL inContext = FALSE;
#define DO_MENU_CONTEXTMENU(com) \
   { \
      POINT pos; \
      GetCursorPos(&pos); \
      contMenuCom = com; \
      inContext = TRUE; \
      TrackPopupMenu(IsDir(GetTrueTarget(contMenuCom)) ? gDirPopupMenu : gFilePopupMenu, TPM_LEFTBUTTON | TPM_RECURSE, pos.x, pos.y, 0, hWnd, NULL); \
      inContext = FALSE; \
      PostMessage(hWnd, WM_CANCELMODE, 0, 0); \
   }

   // Need to keep track of opened submenus in order to handle right click on cascading submenus
   static HMENU openMenu = NULL;

	switch (message)
	{
      case WM_PAINT:
         {
            HDC         hDC;
            PAINTSTRUCT ps;
	         hDC = BeginPaint(hWnd, &ps);
            // Draw the button, higlighted if still tracking
            DrawButton(hWnd, hDC, tracker == hWnd ? eHighlight : eNormal);
	         EndPaint(hWnd, &ps);
         }
         break;

      case WM_LBUTTONDOWN:
         // Left button pressed, start visualization of a pressed button
         DrawButton(hWnd, NULL, ePushed);
         break;

      case WM_RBUTTONDOWN:
         // Right button pressed on button, show appropriate context menu
          DO_MENU_CONTEXTMENU((GET_COM_INFO(hWnd))->command);
         break;

      case WM_MENURBUTTONUP:
         // Right button pressed during TrackPopupMenu
         DO_MENU_CONTEXTMENU(*(CString*)GetMenuItemData((HMENU)lParam, (INT)wParam));
         break;

      case WM_MOUSEMOVE:
         // Cursor has entered this button
         if (!tracker)
         {
            // Start tracking in order to detect when leaving
            TRACKMOUSEEVENT tme;
            tme.cbSize = sizeof(TRACKMOUSEEVENT);
            tme.dwFlags = TME_LEAVE;
            tme.hwndTrack = hWnd;
            if (TrackMouseEvent(&tme))
            {
               // Highlight the button
               DrawButton(hWnd, NULL, eHighlight);
	            tracker = hWnd;
            }
         }
         break;

      case WM_MOUSELEAVE:
         // Cursor is leaving this button
         InvalidateRect(hWnd, NULL, TRUE); // Back to normal
         tracker = NULL;
         if (GetAsyncKeyState(VK_LBUTTON) & 0x8000)
         {
            // Button is down, start dragging
            dragging = TRUE;
            POINT pt = {0, 0};
            // Simulate window title drag event to initiate the drag operation
            PostMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, (LPARAM)&pt);
            SetCursor(LoadCursor(NULL, IDC_HAND));
         }
         break;

      case WM_LBUTTONUP:
         // Button released over this button, execute the corresponding command
         InvalidateRect(hWnd, NULL, TRUE); // Back to normal
         pCom = GET_COM_INFO(hWnd);
         if (pCom)
         {
            if (pCom->hMenu)
            {
               POINT pos;
               GetCursorPos(&pos);
               SetLargeMenus(gLargeMenus);
               openMenu = NULL;
               TrackPopupMenu(pCom->hMenu, 
                              TPM_LEFTBUTTON,
					               pos.x, pos.y, 0, hWnd, NULL);
               SetLargeMenus(FALSE);
            }
            else
               DoRun(pCom->command, pCom->params, gMainWindow, pCom->showType);
         }
         break;

      case WM_EXITSIZEMOVE:
         if (dragging)
         {
            // Dragging ended. Drop the button on the new location
            POINT pt = {0, 0};
            MapWindowPoints(hWnd, gMainWindow, &pt, 1);
            SetCursor(LoadCursor(NULL, IDC_ARROW));
            dragging = FALSE;
            LONG  oldInd = Hwnd2Index(hWnd),
                  newInd = Pos2Index(pt.x, pt.y),
                  i;
            if (oldInd < 0 || newInd < 0)
               break;
            if (newInd <= oldInd)
               // Move up
               for (i = oldInd; i > newInd; i--)
                  gButtons.list[i] = gButtons.list[i-1];
            else
               // Move down
               for (i = oldInd; i < newInd; i++)
                  gButtons.list[i] = gButtons.list[i+1];
            gButtons.list[newInd] = hWnd;
            // Update accordingly
            SetupLayout();
            SavePrefs();
         }
         break;

      case WM_DROPFILES:
         {
            // A file was dopped on this button
            CString fileName;
            DragQueryFile((HDROP)wParam, 0, fileName.GetBuffer(MAX_PATH), MAX_PATH);
            fileName.ReleaseBuffer();
            pCom = GET_COM_INFO(hWnd);
            if (pCom && pCom->hMenu)
            {
               // Dropped on a menu button,create shortcut in the corresponding directory
               File2Shortcut(fileName, pCom->command);
            }
            else if (IS_SHORTCUT(fileName) || SHGetFileInfo(fileName, 0, NULL, 0, SHGFI_EXETYPE))
            {
               // Shortcut or executable file, add it as a new button at this position
               if (HandleDragAndDrop(fileName))
                  AddNewButton(fileName, Hwnd2Index(hWnd));
            }
            else if (pCom)
            {
               // Non-executable file, use it as parameter for the corresponding command
               DoRun(pCom->command, fileName, gMainWindow, pCom->showType);
            }
            DragFinish((HDROP)wParam);
         }
         break;

      case WM_INITMENUPOPUP:
         // Submenu opened
         if (!inContext)
            openMenu = (HMENU)wParam;
         break;

      case WM_UNINITMENUPOPUP:
         // Submenu closed
         openMenu = NULL;
         break;

      case WM_CONTEXTMENU:
         if (openMenu)
         {
            // Right button pressed on top of cascading submenu
            // Get corresponding directory and do context menu popup
            MENUINFO menuInf;
            ZeroMemory(&menuInf, sizeof(menuInf));
            menuInf.cbSize = sizeof(menuInf);
            menuInf.fMask = MIM_MENUDATA;
            GetMenuInfo(openMenu, &menuInf);
            CString *link = (CString*)(menuInf.dwMenuData);
            if (link)
               DO_MENU_CONTEXTMENU(*link);
         }
         break;

	   case WM_COMMAND:
	      // Popup menu selection
         EnableAutoHide(FALSE);
	      switch (LOWORD(wParam))
	      {
	         case IDM_DELETE:
               {
                  // Show Explorer delete dialog
                  SHFILEOPSTRUCT fileOp;
                  CString fileName;
                  ZeroMemory(&fileOp, sizeof(fileOp));
                  fileName = contMenuCom;
                  fileOp.wFunc = FO_DELETE;
                  fileOp.fFlags = FOF_ALLOWUNDO | FOF_WANTNUKEWARNING;
                  PTCHAR  ptr = fileName.GetBuffer(fileName.GetLength()+2);
                  ptr[fileName.GetLength()+1] = 0; // Extra last zero for SHFILEOPSTRUCT
                  fileOp.pFrom = ptr;
                  SHFileOperation(&fileOp);
               }
		         break;

	         case IDM_PROPERTIES:
               DoRun(contMenuCom, EMPTY_STR, gMainWindow, SW_SHOWNORMAL, _T("properties"));
               break;

	         case IDM_EXPLORE:
               // Start Explorer with the target file selected
               DoRun(_T("explorer.exe"), _T("/select,\"") + GetTrueTarget(contMenuCom) + _T("\""), gMainWindow);
		         break;

	         case IDM_RUNASADMIN:
               // Elevate the corresponding command
               DoRun(contMenuCom, NULL, gMainWindow, SW_SHOWNORMAL, _T("runas"));
               break;

	         case IDM_NEW:
               {
                  CString fileName = PromptFileName(EMPTY_STR, gMainWindow);
                  if (fileName.GetLength())
                     File2Shortcut(fileName, contMenuCom);
               }
		         break;

	         case IDM_ADDMENU:
		         DialogBoxParam(CURR_INSTANCE, MAKEINTRESOURCE(IDD_ADDMENU), hWnd, PromptDirectory, (LPARAM)&contMenuCom);
		         break;

	         default:
		         return DefWindowProc(hWnd, message, wParam, lParam);
	      }
         EnableAutoHide(TRUE);
		   break;

	   case WM_MENUCOMMAND:
         {
            // Button popup menu selection, launch the appropriate command
            DWORD pos = (DWORD)wParam;
            HMENU hMenu = (HMENU)lParam;
            CString *path = (CString*)GetMenuItemData(hMenu, pos);
            if (path)
               DoRun(*path, EMPTY_STR, gMainWindow, SW_SHOWNORMAL);
         }
         break;

      case WM_MEASUREITEM:
      case WM_DRAWITEM:
         // Menu icon message
         HandleMenuItemIconMessage(message, lParam);
         break;

      default:
         return DefWindowProc(hWnd, message, wParam, lParam);
   }

   return 0;
}

//--------------------------------------------------------------------------

ATOM RegButtonWindClass()
{
   // Rgister a class for the button windows
	WNDCLASSEX wcex;

   ZeroMemory(&wcex, sizeof(WNDCLASSEX));
	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.lpfnWndProc	= ButtonWndProc;
	wcex.hInstance		= CURR_INSTANCE;
	wcex.lpszClassName = BUTTON_CLASS_NAME;

   wcex.hbrBackground = gBgBrush;

	return RegisterClassEx(&wcex);
}

//--------------------------------------------------------------------------

BOOL ParseSetting(LPCTSTR str)
{
   // Evaluate possible settings specification from config file or command line
   return (
       _stscanf_s(str, _T("POSITION=%d"), &gLocation) ||
       _stscanf_s(str, _T("LARGE=%d"), &gLargeIcons) ||
       _stscanf_s(str, _T("LARGEMENUS=%d"), &gLargeMenus) ||
       _stscanf_s(str, _T("ONTOP=%d"), &gOnTop) ||
       _stscanf_s(str, _T("AUTOHIDE=%d"), &gAutoHide)
       );
}

//--------------------------------------------------------------------------

#define SEP _T(";")

BOOL ParseConfigFile(LPCTSTR fileName)
{
   // Get buttons from a configuration file instead of a directory
   const DWORD buffSize = 1024;
   CString line;
	FILE *inFile;
   if (_tfopen_s(&inFile, fileName, _T("rt")))
      ShowMessage(LoadFormatResString(IDS_CONF_ERR, fileName), eFatal);
   while (_fgetts(line.GetBuffer(buffSize), buffSize, inFile))
	{
      line.ReleaseBuffer();
      line.Trim();
      if (line.IsEmpty() || line[0] == '#')
         // Comment or empty, ignore
         continue;

      // Settings specification
      if (ParseSetting(line))
          continue;

      ExpEnvVars(line);
      
      // Parse the information on this line
      int pos = 0;
      CString toolTip = line.Tokenize(SEP, pos);
      CString command = line.Tokenize(SEP, pos);
      if (command.IsEmpty() || !LocateFile(command))
         continue;
      CString params = EMPTY_CSTR, 
              iconFile = EMPTY_CSTR;
      // Get optional parameters
      int iconInd = 0;
      if (pos > 0)
         params = line.Tokenize(SEP, pos);
      if (pos > 0)
         iconFile = line.Tokenize(SEP, pos);
      if (pos > 0)
         iconInd = _tstoi(line.Tokenize(SEP, pos));

      AddNewButton(command, gButtons.cnt, params, iconFile, iconInd, toolTip);
	}
	fclose(inFile);
   gConfigFileUsed = TRUE;
   gUseReg = FALSE;

   return TRUE;
}

//--------------------------------------------------------------------------

BOOL UpdateButtons()
{
   // Update the buttons in accordance to the current state of the main directory
   static time_t lastUpdate = time(NULL);

   if (gConfigFileUsed)
      return TRUE;

   DWORD i;
   // Validate current buttons
   for (i = 0; i < gButtons.cnt; i++)
   {
      pCommandInfo pCom = GET_COM_INFO(gButtons.list[i]);
      if (!FileExists(pCom->command))
      {
         // The file or shortcut has been deleted
         if (pCom->hMenu)
            DestroyMenuData(pCom->hMenu, TRUE);
         delete pCom;
         pCom = NULL;
         DestroyWindow(gButtons.list[i]);
         // Move up
         DWORD k;
         for (k = i; k < gButtons.cnt; k++)
            gButtons.list[k] = gButtons.list[k+1];
         gButtons.cnt--;
      }
      else if ((IS_SHORTCUT(pCom->command) || IsDir(pCom->command)) && FileModTime(pCom->command) > lastUpdate)
      {
         // Update icons and tooltip
         CString dum, toolTip;
         DestroyIcon(pCom->hIcon); DestroyIcon(pCom->hSmallIcon); DestroyWindow(pCom->hToolTip);
         GetShortcutInfo(pCom->command, dum, &pCom->hIcon, &pCom->hSmallIcon, &toolTip);
         pCom->hToolTip = CreateTooltip(gButtons.list[i], toolTip);
      }
      if (pCom && pCom->hMenu)
      {
         // Always update folder menus (quick and dirty really)
         DestroyMenuData(pCom->hMenu, TRUE);
         pCom->hMenu = AddDirMenu(GetTrueTarget(pCom->command), TRUE, EMPTY_CSTR);
      }
   }

   // Scan for possible new entries in the directory and add buttons when applicable
   CString path = gMainDir + _T("*.*"), 
           fileName = EMPTY_CSTR;
   HANDLE ref;
   while (FindFiles(path, fileName, &ref))
      if (Com2Index(fileName) == -1)
         AddNewButton(fileName);

   SavePrefs();
   lastUpdate = time(NULL);
   return TRUE;

}

//--------------------------------------------------------------------------

BOOL Refresh()
{
	// Rescan and update all buttons and menus
   SetCursor(LoadCursor(NULL, IDC_WAIT));
	UpdateButtons();
	SetupLayout();
   SetCursor(LoadCursor(NULL, IDC_ARROW));
   return TRUE;
}

//--------------------------------------------------------------------------

INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
   // Show about dialog
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	   case WM_INITDIALOG:
         {
            TCHAR templ[100];
            GetDlgItemText(hDlg, IDC_VERSION, templ, STR_SIZE(templ));
            CString ver;
            ver.Format(templ, VERSION_STR, DATE_STR);
            SetDlgItemText(hDlg, IDC_VERSION, ver);
         }

         PositionDialog(hDlg);
		   return (INT_PTR)TRUE;

         case WM_NOTIFY:
            switch (((LPNMHDR)lParam)->code)
            {
               case NM_CLICK:
               case NM_RETURN:
               {
                  // Application home page link
                  PNMLINK pNMLink = (PNMLINK)lParam;
                  LITEM item = pNMLink->item;
                  DoRun(item.szUrl, EMPTY_STR, gMainWindow);
                  break;
               }
          }
          break;

	   case WM_COMMAND:
		   if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		   {
			   EndDialog(hDlg, LOWORD(wParam));
			   return (INT_PTR)TRUE;
		   }
		   break;
	}
	return (INT_PTR)FALSE;
}

//--------------------------------------------------------------------------

#define AUTOSTART_SHORTCUT GetAutoStartDir() + PROG_NAME + SHORTCUT_EXT
INT_PTR CALLBACK Settings(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
#define UPDATE_CLICK_OP \
   EnableWindow(GetDlgItem(hDlg, IDC_CHECK4), IsDlgButtonChecked(hDlg, IDC_CHECK3)); \
   if (!IsDlgButtonChecked(hDlg, IDC_CHECK3)) CheckDlgButton(hDlg, IDC_CHECK4, FALSE);
                        
   // Show settings dialog
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	   case WM_INITDIALOG:
         CheckRadioButton(hDlg, IDC_RADIO1, IDC_RADIO4, IDC_RADIO1 + gLocation-1);
         CheckDlgButton(hDlg, IDC_CHECK1, gLargeIcons);
         CheckDlgButton(hDlg, IDC_CHECK2, gOnTop);
         CheckDlgButton(hDlg, IDC_CHECK3, gAutoHide > 0);
         CheckDlgButton(hDlg, IDC_CHECK4, gAutoHide > 1);
         CheckDlgButton(hDlg, IDC_CHECK6, gLargeMenus);
         UPDATE_CLICK_OP
         CheckDlgButton(hDlg, IDC_CHECK5, FileExists(AUTOSTART_SHORTCUT));

         PositionDialog(hDlg);
		   return (INT_PTR)TRUE;

	   case WM_COMMAND:
		   if (LOWORD(wParam) == IDC_CHECK3)
         {
            UPDATE_CLICK_OP
         }
		   else if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		   {
            if (LOWORD(wParam) == IDOK)
            {
               if (IsDlgButtonChecked(hDlg, IDC_RADIO1))
                  gLocation = 1;
               else if (IsDlgButtonChecked(hDlg, IDC_RADIO2))
                  gLocation = 2;
               else if (IsDlgButtonChecked(hDlg, IDC_RADIO3))
                  gLocation = 3;
               else if (IsDlgButtonChecked(hDlg, IDC_RADIO4))
                  gLocation = 4;
               gLargeIcons = IsDlgButtonChecked(hDlg, IDC_CHECK1);
               gOnTop = IsDlgButtonChecked(hDlg, IDC_CHECK2);
               gAutoHide = 0;
               if (IsDlgButtonChecked(hDlg, IDC_CHECK3))
                  gAutoHide = 1;
               if (IsDlgButtonChecked(hDlg, IDC_CHECK4))
                  gAutoHide = 2;
               BOOL doRefresh = IsDlgButtonChecked(hDlg, IDC_CHECK6) != gLargeMenus;
               gLargeMenus = IsDlgButtonChecked(hDlg, IDC_CHECK6);
               if (gUseReg)
               {
                  // Create or delete autostart shortcut
                  if (IsDlgButtonChecked(hDlg, IDC_CHECK5))
                     CreateShortcut(AUTOSTART_SHORTCUT, gAppPath, LoadFormatResString(IDS_APPDESCR));
                  else
                     DeleteFile(AUTOSTART_SHORTCUT);
               }
               // Update current state accordingly
               if (doRefresh) Refresh(); else SetupLayout();
               SavePrefs();
               AutoHide(gAutoHide);
            }
			   EndDialog(hDlg, LOWORD(wParam));
			   return (INT_PTR)TRUE;
		   }
		   break;
	}
	return (INT_PTR)FALSE;
}

//--------------------------------------------------------------------------

LRESULT CALLBACK MainWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
   static BOOL selected = FALSE,
               tracking = FALSE;
   POINT pos;
	switch (message)
	{
      case WM_PAINT:
         {
            // Draw lines and simple markers at both ends of the main window
            HDC hDC;
            PAINTSTRUCT ps;
            POINT p1, p2;

	         hDC = BeginPaint(hWnd, &ps);
            SelectObject(hDC, gBgBrush);
            if (gLocation == 1 || gLocation == 3)
            {
               SelectObject(hDC, gLightPen);
               MoveToEx(hDC, 0, WIND_OFF-2, NULL);
               LineTo(hDC, gWindowSize.x, WIND_OFF-2);
               MoveToEx(hDC, 0, gWindowSize.y - WIND_OFF-1, NULL);
               LineTo(hDC, gWindowSize.x, gWindowSize.y - WIND_OFF-1);

               SelectObject(hDC, gShadowPen);
               MoveToEx(hDC, 0, WIND_OFF-1, NULL);
               LineTo(hDC, gWindowSize.x, WIND_OFF-1);
               MoveToEx(hDC, 0, gWindowSize.y - WIND_OFF, NULL);
               LineTo(hDC, gWindowSize.x, gWindowSize.y - WIND_OFF);

               p1.x = p2.x = (gWindowSize.x)/2 - 1;
               p1.y = MARK_RADIUS + yOffset;
               p2.y = gWindowSize.y - (MARK_RADIUS + yOffset) - 1;
            }
            else
            {
               SelectObject(hDC, gLightPen);
               MoveToEx(hDC, WIND_OFF-2, 0, NULL);
               LineTo(hDC, WIND_OFF-2, gWindowSize.y);
               MoveToEx(hDC, gWindowSize.x - WIND_OFF-1, 0, NULL);
               LineTo(hDC, gWindowSize.x - WIND_OFF-1, gWindowSize.y);

               SelectObject(hDC, gShadowPen);
               MoveToEx(hDC, WIND_OFF-1, 0, NULL);
               LineTo(hDC, WIND_OFF-1, gWindowSize.y);
               MoveToEx(hDC, gWindowSize.x - WIND_OFF, 0, NULL);
               LineTo(hDC, gWindowSize.x - WIND_OFF, gWindowSize.y);

               p1.x = MARK_RADIUS + xOffset;
               p2.x = gWindowSize.x - (MARK_RADIUS + xOffset) - 1;
               p1.y = p2.y = (gWindowSize.y)/2 - 1;
            }
            Ellipse(hDC, p1.x-MARK_RADIUS, p1.y-MARK_RADIUS, p1.x+MARK_RADIUS, p1.y+MARK_RADIUS);
            Ellipse(hDC, p2.x-MARK_RADIUS, p2.y-MARK_RADIUS, p2.x+MARK_RADIUS, p2.y+MARK_RADIUS);
	         EndPaint(hWnd, &ps);
         }
         break;

	   case WM_COMMAND:
         {
		      // Menu selection
            EnableAutoHide(FALSE);
		      switch (LOWORD(wParam))
		      {
		         case IDM_HELP:
                  DoRun(_T("http://www.lerup.com/LaunchBar"), EMPTY_STR, gMainWindow);
			         break;

		         case IDM_ABOUT:
			         DialogBox(CURR_INSTANCE, MAKEINTRESOURCE(IDD_ABOUT), hWnd, About);
			         break;

		         case IDM_SETTINGS:
			         DialogBox(CURR_INSTANCE, MAKEINTRESOURCE(IDD_SETTINGS), hWnd, Settings);
			         break;

		         case IDM_NEW:
                  {
                     CString fileName = PromptFileName(EMPTY_STR, gMainWindow);
                     if (fileName.GetLength())
                        // Simulate drag and drop
                        HandleDragAndDrop(fileName);
                  }
			         break;

		         case IDM_ADDMENU:
			         DialogBoxParam(CURR_INSTANCE, MAKEINTRESOURCE(IDD_ADDMENU), hWnd, PromptDirectory, (LPARAM)&gMainDir);
			         break;

		         case IDM_REFRESH:
                  Refresh();
			         break;

	            case IDM_EXPLORE:
                  // Start Explorer with the main directory selected
                  DoRun(_T("explorer.exe"), gMainDir, gMainWindow);
		            break;

	            case IDM_ADDSTARTMENU:
                  // Create start menu shortcut
                  CreateShortcut(gStartMenuLink, GetStartMenuDir(TRUE), NULL, NULL, NULL, gAppPath, -IDI_STARTMENU);
		            break;

		         case IDM_RUN:
                  // Show standard "run" dialog
			         RunDialog(hWnd);
			         break;

		         case IDM_EXIT:
			         DestroyWindow(hWnd);
			         PostQuitMessage(0);
			         break;

		         default:
			         return DefWindowProc(hWnd, message, wParam, lParam);
		      }
         }
         EnableAutoHide(TRUE);
		   break;

      case WM_DROPFILES:
         {
            // Drag and drop on main window
            CString fileName;
            DragQueryFile((HDROP)wParam, 0, fileName.GetBuffer(MAX_PATH), MAX_PATH);
            fileName.ReleaseBuffer();
            HandleDragAndDrop(fileName);
            DragFinish((HDROP)wParam);
         }
         break;

      case WM_LBUTTONDOWN:
      case WM_RBUTTONDOWN:
         if (gAutoHide > 1 && gHidden)
         {
            AutoHide(FALSE);
         }
         else
         {
            // Do menu selection
            GetCursorPos(&pos);
            EnableMenuItem(gMainPopupMenu, IDM_ADDSTARTMENU, MF_BYCOMMAND | (FileExists(gStartMenuLink) ? MF_GRAYED : MF_ENABLED));
            TrackPopupMenu(gMainPopupMenu, 
                           TPM_LEFTBUTTON,
					            pos.x, pos.y, 0, hWnd, NULL);
         }
         break;

      case WM_MOUSEMOVE:
         // Cursor has entered the window, show window if hidden
         if (gAutoHide == 1)
           AutoHide(FALSE);
         break;

      case WM_TIMER:
         if (gAutoHide && GetForegroundWindow() != gMainWindow)
         {
            // Not in foreground
            POINT p;
            GetCursorPos(&p);
            if (p.x < gWindowRect.left || p.x > gWindowRect.right ||
               p.y < gWindowRect.top || p.y > gWindowRect.bottom)
               // Cursor is outside of the toolbar, hide it
               AutoHide(TRUE);
         }
         break;

      case WM_USER_DIR_CHANGED:
         // Change occured in the directory, do a rescan
         Refresh();
         break;

      case WM_DISPLAYCHANGE:
         // Display has been resized, redo layout
         SetupLayout();
         break;

      case WM_MEASUREITEM:
      case WM_DRAWITEM:
         // Draw a menu icon
         HandleMenuItemIconMessage(message, lParam);
         break;

      case WM_CLOSE:
         // Avoid unintentional close
         break;

	   default:
		   return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

//--------------------------------------------------------------------------

ATOM RegMainWindClass()
{
   // Register the main window class
	WNDCLASSEX wcex;

   ZeroMemory(&wcex, sizeof(WNDCLASSEX));
	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.lpfnWndProc	= MainWndProc;
	wcex.hInstance		= CURR_INSTANCE;
	wcex.hIcon			= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_APP));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.lpszClassName = WIND_CLASS_NAME;
	wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_APP));

   wcex.hbrBackground = gBgBrush;

	return RegisterClassEx(&wcex);
}

//--------------------------------------------------------------------------

// Settings registry key names
#define LARGE_ICONS_KEY _T("UseLargeIcons")
#define LARGE_MENUS_KEY _T("UseLargeMenus")
#define ALWAYS_ON_TOP_KEY _T("AlwaysOnTop")
#define AUTOHIDE_KEY _T("AutoHide")
#define LOCATION_KEY _T("Location")
#define BUTTONS_KEY _T("Buttons")

BOOL ReadPrefs()
{
   if (!gUseReg)
      return FALSE;

   // Read preferences from the registry
   GET_REG_INT(LARGE_ICONS_KEY, gLargeIcons);
   GET_REG_INT(LARGE_MENUS_KEY, gLargeMenus);
   GET_REG_INT(ALWAYS_ON_TOP_KEY, gOnTop);
   GET_REG_INT(AUTOHIDE_KEY, gAutoHide);
   GET_REG_INT(LOCATION_KEY, gLocation);

   CString buf = GetRegVal(BUTTONS_KEY);
   int pos = 0;
   CString name = buf.Tokenize(SEP, pos);
   while (name != EMPTY_CSTR)
   {
      AddNewButton(gMainDir + name);
      name = buf.Tokenize(SEP, pos);
   }

   return TRUE;
}

//--------------------------------------------------------------------------

BOOL SavePrefs()
{
	if (!gUseReg)
	  return FALSE;

	// Save preferences to the registry
	SET_REG_INT(LARGE_ICONS_KEY, gLargeIcons);
	SET_REG_INT(LARGE_MENUS_KEY, gLargeMenus);
	SET_REG_INT(ALWAYS_ON_TOP_KEY, gOnTop);
	SET_REG_INT(AUTOHIDE_KEY, gAutoHide);
	SET_REG_INT(LOCATION_KEY, gLocation);

	// Save the current order of the buttons
   DWORD i;
   CString buf = EMPTY_STR;
   for (i = 0; i < gButtons.cnt; i++)
   {
      pCommandInfo pCom = GET_COM_INFO(gButtons.list[i]);
      buf += GetFileNameComp(pCom->command, eFcName | eFcType) + SEP;
   }
   SetRegVal(BUTTONS_KEY, buf);

	return TRUE;
}

//--------------------------------------------------------------------------

BOOL InitInstance(HINSTANCE hInstance, LPCTSTR lpCmdLine)
{
   // Register classes and create the main window
   if (!RegMainWindClass() || !RegButtonWindClass())
      return FALSE;

   gMainWindow = CreateWindow(WIND_CLASS_NAME, PROG_NAME, WS_POPUP | WS_BORDER,
                        0, 0, 100, 100,
                        NULL, NULL, hInstance, NULL);

   if (!gMainWindow)
      return FALSE;

   // Don't appear in the taskbar
   SetWindowLong(gMainWindow, GWL_EXSTYLE, GetWindowLong(gMainWindow, GWL_EXSTYLE) | WS_EX_TOOLWINDOW);

   GetModuleFileName(NULL, gAppPath.GetBuffer(MAX_PATH), MAX_PATH);
   gAppPath.ReleaseBuffer();

   // Check for possible command line parameters
   CString cmdLine = lpCmdLine;
   INT start = 0;
   if (!cmdLine.IsEmpty())
   {
      // Get possible settings specified on the command line
      CString par;
      INT pos = 0;
      while ((par = cmdLine.Tokenize(_T(" "), pos)) != EMPTY_STR)
      {
         if (ParseSetting(par))
         {
            // Setting found, avoid using the registry
            gUseReg = FALSE;
            start = pos;
         }
         else
			 // No more settings, break here
            break;

      }
   }
   // Get remaining part of the command line
   cmdLine.Delete(0, start);

   if (cmdLine.IsEmpty())
   {
      // Default, use the quick launch directory
      gMainDir = GetQuickLaunchDir();
   }
   else
   {
      if (IsDir(cmdLine))
         gMainDir = cmdLine;
      else
         ParseConfigFile(cmdLine);
   }

   // Define registry root and read possible saved preferences
   SetAppRegRoot(_T("Software\\Peter Lerup\\"PROG_NAME));

   if (!gConfigFileUsed)
   {
      // Validate
      if (gMainDir.IsEmpty())
         ShowMessage(LoadFormatResString(IDS_DIR_EMPTY), eFatal);
      else if (!FileExists(gMainDir))
         ShowMessage(LoadFormatResString(IDS_DIR_NOTFOUND, gMainDir), eFatal);

      // Directory must contain a trailing backspace
      APPEND_BS(gMainDir);

      // Read possible saved preferences
      ReadPrefs();
   }

   gStartMenuDir = GetStartMenuDir(TRUE);
   TRIM_BS(gStartMenuDir);
   gStartMenuLink = gMainDir + _T("Start Menu") + SHORTCUT_EXT;

   // Setup the layout
   Refresh();
   EnableAutoHide(TRUE);

   // Init application popup menus and set item icons
   gMainPopupMenu = LoadMenu(CURR_INSTANCE, MAKEINTRESOURCE(IDR_MENU1));
   InitPopupMenu(gMainPopupMenu);
   SetMenuItemData(gMainPopupMenu, IDM_HELP, GET_ICON(IDI_HELP));
   SetMenuItemData(gMainPopupMenu, IDM_ABOUT, GET_ICON(IDI_APP));
   SetMenuItemData(gMainPopupMenu, IDM_SETTINGS, GET_ICON(IDI_SETTINGS));
   SetMenuItemData(gMainPopupMenu, IDM_NEW, GET_ICON(IDI_NEW));
   SetMenuItemData(gMainPopupMenu, IDM_ADDMENU, GET_ICON(IDI_MENU));
   SetMenuItemData(gMainPopupMenu, IDM_EXPLORE, GET_ICON(IDI_EXPLORE));
   SetMenuItemData(gMainPopupMenu, IDM_REFRESH, GET_ICON(IDI_REFRESH));
   SetMenuItemData(gMainPopupMenu, IDM_ADDSTARTMENU, GET_ICON(IDI_STARTMENU));
   SetMenuItemData(gMainPopupMenu, IDM_RUN, GET_ICON(IDI_RUN));
   SetMenuItemData(gMainPopupMenu, IDM_EXIT, GET_ICON(IDI_EXIT));
   if (IsWinPE())
      DeleteMenu(gMainPopupMenu, IDM_EXIT, MF_BYCOMMAND);

   gFilePopupMenu = LoadMenu(CURR_INSTANCE, MAKEINTRESOURCE(IDR_MENU2));
   InitPopupMenu(gFilePopupMenu);
   SetMenuItemData(gFilePopupMenu, IDM_DELETE, GET_ICON(IDI_DELETE));
   SetMenuItemData(gFilePopupMenu, IDM_PROPERTIES, GET_ICON(IDI_PROPERTIES));
   SetMenuItemData(gFilePopupMenu, IDM_EXPLORE, GET_ICON(IDI_EXPLORE));
   SetMenuItemData(gFilePopupMenu, IDM_RUNASADMIN, GET_ICON(IDI_ADMIN));

   gDirPopupMenu = LoadMenu(CURR_INSTANCE, MAKEINTRESOURCE(IDR_MENU3));
   InitPopupMenu(gDirPopupMenu);
   SetMenuItemData(gDirPopupMenu, IDM_DELETE, GET_ICON(IDI_DELETE));
   SetMenuItemData(gDirPopupMenu, IDM_PROPERTIES, GET_ICON(IDI_PROPERTIES));
   SetMenuItemData(gDirPopupMenu, IDM_EXPLORE, GET_ICON(IDI_EXPLORE));
   SetMenuItemData(gDirPopupMenu, IDM_NEW, GET_ICON(IDI_NEW));
   SetMenuItemData(gDirPopupMenu, IDM_ADDMENU, GET_ICON(IDI_MENU));


   if (!gConfigFileUsed)
   {
      // Accept drag and drop in the main window spare areas
      DragAcceptFiles(gMainWindow, TRUE);
      StartWatchDir(gMainDir, gMainWindow, WM_USER_DIR_CHANGED, NULL);
   }
   else
   {
      // Remove some menu items 
      DeleteMenu(gMainPopupMenu, IDM_NEW, MF_BYCOMMAND);
      DeleteMenu(gMainPopupMenu, IDM_ADDMENU, MF_BYCOMMAND);
      DeleteMenu(gMainPopupMenu, IDM_REFRESH, MF_BYCOMMAND);
      DeleteMenu(gMainPopupMenu, IDM_EXPLORE, MF_BYCOMMAND);
   }

   // Tooltip
   CreateTooltip(gMainWindow, LoadFormatResString(IDS_MAIN_TIP));

   EnableAutoHide(TRUE);

   return TRUE;
}

//--------------------------------------------------------------------------

int APIENTRY WinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPSTR     lpCmdLine,
                     int       nCmdShow)
{
   CoInitialize(NULL);

	if (!InitInstance(hInstance, CString(lpCmdLine)))
		return 1;

	MSG msg;
	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0))
	{
	   TranslateMessage(&msg);
	   DispatchMessage(&msg);
	}

	return (int) msg.wParam;

}
