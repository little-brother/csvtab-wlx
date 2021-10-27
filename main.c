#define UNICODE
#define _UNICODE

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <tchar.h>
#include <stdlib.h>
#include <stdio.h>
#include <ctype.h>
#include <math.h>

#define LVS_EX_AUTOSIZECOLUMNS 0x10000000

#define WMU_UPDATE_CACHE       WM_USER + 1
#define WMU_UPDATE_GRID        WM_USER + 2
#define WMU_UPDATE_RESULTSET   WM_USER + 3
#define WMU_UPDATE_FILTER_SIZE WM_USER + 4
#define WMU_AUTO_COLUMN_SIZE   WM_USER + 5
#define WMU_RESET_CACHE        WM_USER + 6
#define WMU_SET_FONT           WM_USER + 7

#define IDC_MAIN               100
#define IDC_GRID               101
#define IDC_STATUSBAR          102
#define IDC_HEADER_EDIT        1000

#define IDM_COPY_CELL          5000
#define IDM_COPY_ROW           5001
#define IDM_HEADER_ROW         5002

#define SB_VERSION             0
#define SB_RESERVED            1
#define SB_CODEPAGE            2
#define SB_DELIMITER           3
#define SB_ROW_COUNT           4
#define SB_CURRENT_ROW         5
#define SB_ERROR               6

#define MAX_TEXT_LENGTH        32000
#define MAX_DATA_LENGTH        32000
#define MAX_COLUMN_LENGTH      2000
#define MAX_FILTER_LENGTH      2000
#define DELIMITERS             TEXT(",;|\t")
#define APP_NAME               TEXT("csvtab")
#define APP_VERSION            TEXT(" 0.9.0")

#define CP_UTF16LE             1200
#define CP_UTF16BE             1201

typedef struct {
	int size;
	DWORD PluginInterfaceVersionLow;
	DWORD PluginInterfaceVersionHi;
	char DefaultIniName[MAX_PATH];
} ListDefaultParamStruct;

static TCHAR iniPath[MAX_PATH] = {0};

LRESULT CALLBACK cbNewMain (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewFilterEdit (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
void setStoredValue(TCHAR* name, int value);
int getStoredValue(TCHAR* name, int defValue);
TCHAR* getStoredString(TCHAR* name, TCHAR* defValue);
int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam);
TCHAR* utf8to16(const char* in);
char* utf16to8(const TCHAR* in);
int detectCodePage(const unsigned char *data, int len);
TCHAR detectDelimiter(const TCHAR *data);
void setClipboardText(const TCHAR* text);
BOOL isNumber(TCHAR* val);
BOOL isUtf8(const char * string);
int ListView_AddColumn(HWND hListWnd, TCHAR* colName, int fmt);
int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax);

BOOL APIENTRY DllMain (HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
	return TRUE;
}

void __stdcall ListGetDetectString(char* DetectString, int maxlen) {
	snprintf(DetectString, maxlen, "MULTIMEDIA & ext=\"CSV\"");
}

void __stdcall ListSetDefaultParams(ListDefaultParamStruct* dps) {
	if (iniPath[0] == 0) {
		DWORD size = MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, NULL, 0);
		MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, iniPath, size);
	}
}

HWND APIENTRY ListLoad (HWND hListerWnd, char* fileToLoad, int showFlags) {
	DWORD size = MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, NULL, 0);
	TCHAR* filepath = (TCHAR*)calloc (size, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, filepath, size);

	struct _stat st = {0};
	if (_tstat(filepath, &st) != 0 || st.st_size > getStoredValue(TEXT("max-file-size"), 10000000))
		return 0;
	
	INITCOMMONCONTROLSEX icex;
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_LISTVIEW_CLASSES;
	InitCommonControlsEx(&icex);

	BOOL isStandalone = GetParent(hListerWnd) == HWND_DESKTOP;
	HWND hMainWnd = CreateWindow(WC_STATIC, APP_NAME, WS_CHILD | WS_VISIBLE | (isStandalone ? SS_SUNKEN : 0),
		0, 0, 100, 100, hListerWnd, (HMENU)IDC_MAIN, GetModuleHandle(0), NULL);

	SetProp(hMainWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hMainWnd, GWLP_WNDPROC, (LONG_PTR)&cbNewMain));
	SetProp(hMainWnd, TEXT("FILEPATH"), filepath);
	SetProp(hMainWnd, TEXT("FILESIZE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("DELIMITER"), calloc(1, sizeof(TCHAR)));
	SetProp(hMainWnd, TEXT("CODEPAGE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("HEADERROW"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("CACHE"), 0);
	SetProp(hMainWnd, TEXT("RESULTSET"), 0);
	SetProp(hMainWnd, TEXT("ORDERBY"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("COLCOUNT"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("ROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("TOTALROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("COLNO"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FONT"), 0);
	SetProp(hMainWnd, TEXT("FONTFAMILY"), getStoredString(TEXT("font"), TEXT("Arial")));
	SetProp(hMainWnd, TEXT("FONTSIZE"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("GRAYBRUSH"), CreateSolidBrush(GetSysColor(COLOR_BTNFACE)));

	*(int*)GetProp(hMainWnd, TEXT("FILESIZE")) = st.st_size;
	*(int*)GetProp(hMainWnd, TEXT("FONTSIZE")) = getStoredValue(TEXT("font-size"), 16);	
	*(int*)GetProp(hMainWnd, TEXT("HEADERROW")) = getStoredValue(TEXT("header-row"), 1);	
	
	HWND hStatusWnd = CreateStatusWindow(WS_CHILD | WS_VISIBLE |  (isStandalone ? SBARS_SIZEGRIP : 0), NULL, hMainWnd, IDC_STATUSBAR);
	int sizes[7] = {35, 140, 200, 230, 400, 500, -1};
	SendMessage(hStatusWnd, SB_SETPARTS, 7, (LPARAM)&sizes);
	SendMessage(hStatusWnd, SB_SETTEXT, SB_VERSION, (LPARAM)APP_VERSION);

	HWND hGridWnd = CreateWindow(WC_LISTVIEW, NULL, WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_SINGLESEL | LVS_OWNERDATA | WS_TABSTOP,
		205, 0, 100, 100, hMainWnd, (HMENU)IDC_GRID, GetModuleHandle(0), NULL);
	ListView_SetExtendedListViewStyle(hGridWnd, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
	SetProp(hGridWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hGridWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));	

	HWND hHeader = ListView_GetHeader(hGridWnd);
	LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
	SetWindowLongPtr(hHeader, GWL_STYLE, styles | HDS_FILTERBAR);

	HMENU hGridMenu = CreatePopupMenu();
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_CELL, TEXT("Copy cell"));
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_ROW, TEXT("Copy row"));
	AppendMenu(hGridMenu, MF_STRING, 0, NULL);	
	AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("HEADERROW")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_HEADER_ROW, TEXT("Header row"));	
	SetProp(hMainWnd, TEXT("GRIDMENU"), hGridMenu);
	SendMessage(hMainWnd, WMU_UPDATE_CACHE, 0, 0);
	SetFocus(hGridWnd);

	return hMainWnd;
}

void __stdcall ListCloseWindow(HWND hWnd) {
	setStoredValue(TEXT("font-size"), *(int*)GetProp(hWnd, TEXT("FONTSIZE")));
	setStoredValue(TEXT("header-row"), *(int*)GetProp(hWnd, TEXT("HEADERROW")));	
	
	SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
	free((TCHAR*)GetProp(hWnd, TEXT("FILEPATH")));
	free((int*)GetProp(hWnd, TEXT("FILESIZE")));	
	free((TCHAR*)GetProp(hWnd, TEXT("DELIMITER")));
	free((int*)GetProp(hWnd, TEXT("CODEPAGE")));
	free((int*)GetProp(hWnd, TEXT("HEADERROW")));	
	free((int*)GetProp(hWnd, TEXT("ORDERBY")));
	free((int*)GetProp(hWnd, TEXT("COLCOUNT")));	
	free((int*)GetProp(hWnd, TEXT("ROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("FONTSIZE")));	
	free((int*)GetProp(hWnd, TEXT("COLNO")));
	free((TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));

	DeleteFont(GetProp(hWnd, TEXT("FONT")));
	DeleteObject(GetProp(hWnd, TEXT("GRAYBRUSH")));
	DestroyMenu(GetProp(hWnd, TEXT("GRIDMENU")));

	RemoveProp(hWnd, TEXT("WNDPROC"));
	RemoveProp(hWnd, TEXT("CACHE"));
	RemoveProp(hWnd, TEXT("RESULTSET"));
	RemoveProp(hWnd, TEXT("FILEPATH"));
	RemoveProp(hWnd, TEXT("FILESIZE"));
	RemoveProp(hWnd, TEXT("DELIMITER"));
	RemoveProp(hWnd, TEXT("CODEPAGE"));
	RemoveProp(hWnd, TEXT("HEADERROW"));	
	RemoveProp(hWnd, TEXT("ORDERBY"));
	RemoveProp(hWnd, TEXT("COLCOUNT"));
	RemoveProp(hWnd, TEXT("ROWCOUNT"));
	RemoveProp(hWnd, TEXT("TOTALROWCOUNT"));
	RemoveProp(hWnd, TEXT("COLNO"));

	RemoveProp(hWnd, TEXT("FONT"));
	RemoveProp(hWnd, TEXT("FONTFAMILY"));
	RemoveProp(hWnd, TEXT("FONTSIZE"));
	RemoveProp(hWnd, TEXT("GRAYBRUSH"));
	RemoveProp(hWnd, TEXT("GRIDMENU"));

	DestroyWindow(hWnd);
	return;
}

LRESULT CALLBACK cbNewMain(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_SIZE: {
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, WM_SIZE, 0, 0);
			RECT rc;
			GetClientRect(hStatusWnd, &rc);
			int statusH = rc.bottom;

			GetClientRect(hWnd, &rc);
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			SetWindowPos(hGridWnd, 0, 0, 0, rc.right, rc.bottom - statusH, SWP_NOZORDER);
		}
		break;

		// https://groups.google.com/g/comp.os.ms-windows.programmer.win32/c/1XhCKATRXws
		case WM_NCHITTEST: {
			return 1;
		}
		break;
		
		case WM_MOUSEWHEEL: {
			if (LOWORD(wParam) == MK_CONTROL) {
				SendMessage(hWnd, WMU_SET_FONT, GET_WHEEL_DELTA_WPARAM(wParam) > 0 ? 1: -1, 0);
				return 1;
			}
		}
		break;
		
		case WM_KEYDOWN: {
			if (wParam == VK_ESCAPE)
				SendMessage(GetParent(hWnd), WM_CLOSE, 0, 0);

			if (wParam == VK_TAB) {
				HWND hFocus = GetFocus();
				HWND wnds[1000] = {0};
				EnumChildWindows(hWnd, (WNDENUMPROC)cbEnumTabStopChildren, (LPARAM)wnds);

				int no = 0;
				while(wnds[no] && wnds[no] != hFocus)
					no++;

				int cnt = no;
				while(wnds[cnt])
					cnt++;

				BOOL isBackward = HIWORD(GetKeyState(VK_CONTROL));
				no += isBackward ? -1 : 1;
				SetFocus(wnds[no] && no >= 0 ? wnds[no] : (isBackward ? wnds[cnt - 1] : wnds[0]));
			}
		}
		break;
				
		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);
			if (cmd == IDM_COPY_CELL || cmd == IDM_COPY_ROW) {
				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
				if (rowNo != -1) {
					TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
					int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
					int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
					if (!resultset || rowNo >= rowCount)
						return 0;

					int colCount = Header_GetItemCount(hHeader);

					int startNo = cmd == IDM_COPY_CELL ? *(int*)GetProp(hWnd, TEXT("COLNO")) : 0;
					int endNo = cmd == IDM_COPY_CELL ? startNo + 1 : colCount;
					if (startNo > colCount || endNo > colCount)
						return 0;

					int len = 0;
					for (int colNo = startNo; colNo < endNo; colNo++) {
						int _rowNo = resultset[rowNo];
						len += _tcslen(cache[_rowNo][colNo]) + 1 /* column delimiter: TAB */;
					}

					TCHAR* buf = calloc(len + 1, sizeof(TCHAR));
					for (int colNo = startNo; colNo < endNo; colNo++) {
						int _rowNo = resultset[rowNo];
						_tcscat(buf, cache[_rowNo][colNo]);
						if (colNo != endNo - 1)
							_tcscat(buf, TEXT("\t"));
					}

					setClipboardText(buf);
					free(buf);
				}
			}
			
			if (cmd == IDM_HEADER_ROW) {
				HMENU hMenu = (HMENU)GetProp(hWnd, TEXT("GRIDMENU"));
				int* pHeaderRow = (int*)GetProp(hWnd, TEXT("HEADERROW"));
				*pHeaderRow = (*pHeaderRow + 1) % 2;
				
				MENUITEMINFO mii = {0};
				mii.cbSize = sizeof(MENUITEMINFO);
				mii.fMask = MIIM_STATE;
				mii.fState = *pHeaderRow ? MFS_CHECKED : 0;
				SetMenuItemInfo(hMenu, IDM_HEADER_ROW, FALSE, &mii);				
				
				SendMessage(hWnd, WMU_UPDATE_GRID, 0, 0);				
			}
		}
		break;

		case WM_NOTIFY : {
			NMHDR* pHdr = (LPNMHDR)lParam;
			if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_GETDISPINFO) {
				LV_DISPINFO* pDispInfo = (LV_DISPINFO*)lParam;
				LV_ITEM* pItem = &(pDispInfo)->item;

				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
				if(resultset && pItem->mask & LVIF_TEXT) {
					int rowNo = resultset[pItem->iItem];
					pItem->pszText = cache[rowNo][pItem->iSubItem];
				}
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_COLUMNCLICK) {
				NMLISTVIEW* pLV = (NMLISTVIEW*)lParam;
				int colNo = pLV->iSubItem + 1;
				int* pOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));
				int orderBy = *pOrderBy;
				*pOrderBy = colNo == orderBy || colNo == -orderBy ? -orderBy : colNo;
				SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_RCLICK) {
				NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;

				POINT p;
				GetCursorPos(&p);
				*(int*)GetProp(hWnd, TEXT("COLNO")) = ia->iSubItem;
				TrackPopupMenu(GetProp(hWnd, TEXT("GRIDMENU")), TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_ITEMCHANGED) {
				HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);

				TCHAR buf[255] = {0};
				int pos = ListView_GetNextItem(pHdr->hwndFrom, -1, LVNI_SELECTED);
				if (pos != -1)
					_sntprintf(buf, 255, TEXT(" %i"), pos + 1);
				SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_ROW, (LPARAM)buf);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_KEYDOWN) {
				NMLVKEYDOWN* kd = (LPNMLVKEYDOWN) lParam;
				if (kd->wVKey == 0x43 && GetKeyState(VK_CONTROL)) // Ctrl + C
					SendMessage(hWnd, WM_COMMAND, IDM_COPY_ROW, 0);
			}

			if (pHdr->code == HDN_ITEMCHANGED && pHdr->hwndFrom == ListView_GetHeader(GetDlgItem(hWnd, IDC_GRID)))
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;
		
		case WMU_UPDATE_CACHE: {
			TCHAR* filepath = (TCHAR*)GetProp(hWnd, TEXT("FILEPATH"));
			int filesize = *(int*)GetProp(hWnd, TEXT("FILESIZE"));
			TCHAR delimiter = *(TCHAR*)GetProp(hWnd, TEXT("DELIMITER"));
			int codepage = *(int*)GetProp(hWnd, TEXT("CODEPAGE"));	
			
			SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);

			char* data = calloc(filesize + 2, sizeof(char)); // + 2!
			FILE *f = _tfopen(filepath, TEXT("rb"));
			fread(data, sizeof(char), filesize, f);
			fclose(f);
			data[filesize] = 0;	

			if (!codepage)
				codepage = detectCodePage(data, filesize);	
				
			TCHAR* data16 = 0;				
			if (codepage == CP_UTF16BE) {
				for (int i = 0; i < filesize/2; i++) {
					int c = data[2 * i];
					data[2 * i] = data[2 * i + 1];
					data[2 * i + 1] = c;
				}
			}
			
			if (codepage == CP_UTF16LE || codepage == CP_UTF16BE)
				data16 = (TCHAR*)data;
				
			if (codepage == CP_UTF8) {
				data16 = utf8to16(data);
				free(data);
			}
			
			if (codepage == CP_ACP) {
				DWORD len = MultiByteToWideChar(CP_ACP, 0, data, -1, NULL, 0);
				data16 = (TCHAR*)calloc (len, sizeof (TCHAR));
				MultiByteToWideChar(CP_ACP, 0, data, -1, data16, len);
				free(data);
			}
			
			if (!delimiter)
				delimiter = detectDelimiter(data16);

			int len = _tcslen(data16);	
			int colCount = 1;			
			BOOL inQuote = FALSE;
			for (int pos = 0; pos < len; pos++) {
				TCHAR c = data16[pos];
				if (!inQuote && c == TEXT('\n')) 
					break;
					
				colCount += c == delimiter;	
			}
				
			int cacheSize = len / 100 + 1;
			TCHAR*** cache = calloc(cacheSize, sizeof(TCHAR**));
								
			int rowNo = -1;
			inQuote = FALSE;
			int pos = 0;
			int start = 0;			
			while (pos < len) {
				rowNo++;
				if (rowNo >= cacheSize) {
					cacheSize += 100;
					cache = realloc(cache, cacheSize * sizeof(TCHAR**));
				}
				cache[rowNo] = (TCHAR**)calloc (colCount, sizeof (TCHAR*));
				
				int colNo = 0;
				for (; pos < len && colNo < colCount; pos++) {
					TCHAR c = data16[pos];
					
					if (!inQuote && (c == delimiter || c == TEXT('\n') || c == TEXT('\r')) || pos == len - 1) {
						TCHAR* value = calloc(pos - start + 1, sizeof(TCHAR));
						BOOL isQuoted = data16[start] == TEXT('"');
						int vLen = pos - start - 2 * isQuoted;
						_tcsncpy(value, data16 + start + isQuoted, vLen);
						if (_tcschr(value, TEXT('"'))) {
							int j = 0;
							for (int i = 0; i < vLen; i++) {
								if (value[i] == TEXT('"') && value[i + 1] == TEXT('"'))
									continue;
									
								value[j] = value[i];
								j++;
							}
							value[j + 1] = 0;
						}
						
						cache[rowNo][colNo] = value;
						start = pos + 1;
						colNo++;
					}
					
					if (!inQuote && c == TEXT('\n')) 
						break;
		
					if (c == TEXT('"')) { 
						inQuote = !inQuote;
						continue;
					}
				}
				pos++;
			}
								
			cache = realloc(cache, (rowNo + 1) * sizeof(TCHAR*));					
			free(data16);
						
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_CODEPAGE, (LPARAM)(
				codepage == CP_UTF8 ? TEXT("    UTF-8") : 
				codepage == CP_UTF16LE ? TEXT(" UTF-16LE") : 
				codepage == CP_UTF16BE ? TEXT(" UTF-16BE") : TEXT("     ANSI")
			));	
			
			TCHAR buf[32] = {0};
			if (delimiter != TEXT('\t'))
				_sntprintf(buf, 32, TEXT(" %c"), delimiter);
			else
				_sntprintf(buf, 32, TEXT(" TAB"));
			SendMessage(hStatusWnd, SB_SETTEXT, SB_DELIMITER, (LPARAM)buf);
									
			SetProp(hWnd, TEXT("CACHE"), cache);
			*(int*)GetProp(hWnd, TEXT("COLCOUNT")) = colCount;			
			*(int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")) = rowNo + 1;
			*(int*)GetProp(hWnd, TEXT("CODEPAGE")) = codepage;			
			*(TCHAR*)GetProp(hWnd, TEXT("DELIMITER")) = delimiter;
			*(int*)GetProp(hWnd, TEXT("ORDERBY")) = 0;
					
			SendMessage(hWnd, WMU_UPDATE_GRID, 0, 0);
		}
		break;

		case WMU_UPDATE_GRID: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int rowCount = *(int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
			int isHeaderRow = *(int*)GetProp(hWnd, TEXT("HEADERROW"));
			
			int colCount = Header_GetItemCount(hHeader);
			for (int colNo = 0; colNo < colCount; colNo++)
				DestroyWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo));

			for (int colNo = 0; colNo < colCount; colNo++)
				ListView_DeleteColumn(hGridWnd, colCount - colNo - 1);
			
			colCount = *(int*)GetProp(hWnd, TEXT("COLCOUNT"));
			for (int i = 0; i < colCount; i++) {
				int fmt = rowCount > 1 && isNumber(cache[1][i]) ? LVCFMT_RIGHT : LVCFMT_LEFT;				
				TCHAR colName[64];
				_sntprintf(colName, 64, TEXT("Column #%i"), i);
				ListView_AddColumn(hGridWnd, isHeaderRow ? cache[0][i] : colName, fmt);
			}
			
			for (int colNo = 0; colNo < colCount; colNo++) {
				// Use WS_BORDER to vertical text aligment
				HWND hEdit = CreateWindowEx(WS_EX_TOPMOST, WC_EDIT, NULL, ES_CENTER | ES_AUTOHSCROLL | WS_VISIBLE | WS_CHILD | WS_TABSTOP | WS_BORDER,
					0, 0, 0, 0, hHeader, (HMENU)(INT_PTR)(IDC_HEADER_EDIT + colNo), GetModuleHandle(0), NULL);
				SendMessage(hEdit, WM_SETFONT, (LPARAM)GetProp(hWnd, TEXT("FONT")), TRUE);
				SetProp(hEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)cbNewFilterEdit));
			}			
			
			SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);
			SendMessage(hWnd, WMU_SET_FONT, 0, 0);
			PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
		}
		break;

		case WMU_UPDATE_RESULTSET: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			HWND hHeader = ListView_GetHeader(hGridWnd);

			ListView_SetItemCount(hGridWnd, 0);
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int* pTotalRowCount = (int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
			int* pRowCount = (int*)GetProp(hWnd, TEXT("ROWCOUNT"));
			int* pOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));
			int isHeaderRow = *(int*)GetProp(hWnd, TEXT("HEADERROW"));
			int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
			if (resultset)
				free(resultset);

			if (!cache)
				return 1;
				
			if (*pTotalRowCount == 0)	
				return 1;
				
			int colCount = Header_GetItemCount(hHeader);
			if (colCount == 0)
				return 1;
				
			BOOL* bResultset = (BOOL*)calloc(*pTotalRowCount, sizeof(BOOL));
			bResultset[0] = !isHeaderRow;
			for (int rowNo = 1; rowNo < *pTotalRowCount; rowNo++)
				bResultset[rowNo] = TRUE;

			for (int colNo = 0; colNo < colCount; colNo++) {
				HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
				TCHAR filter[MAX_FILTER_LENGTH];
				GetWindowText(hEdit, filter, MAX_FILTER_LENGTH);
				int len = _tcslen(filter);
				if (len == 0)
					continue;

				for (int rowNo = 0; rowNo < *pTotalRowCount; rowNo++) {
					if (!bResultset[rowNo])
						continue;

					TCHAR* value = cache[rowNo][colNo];
					if (len > 1 && (filter[0] == TEXT('<') || filter[0] == TEXT('>')) && isNumber(filter + 1)) {
						TCHAR* end = 0;
						double df = _tcstod(filter + 1, &end);
						double dv = _tcstod(value, &end);
						bResultset[rowNo] = (filter[0] == TEXT('<') && dv < df) || (filter[0] == TEXT('>') && dv > df);
					} else {
						bResultset[rowNo] = len == 1 ? _tcsstr(value, filter) != 0 :
							filter[0] == TEXT('=') ? _tcscmp(value, filter + 1) == 0 :
							filter[0] == TEXT('!') ? _tcsstr(value, filter + 1) == 0 :
							filter[0] == TEXT('<') ? _tcscmp(value, filter + 1) < 0 :
							filter[0] == TEXT('>') ? _tcscmp(value, filter + 1) > 0 :
							_tcsstr(value, filter) != 0;
					}
				}
			}

			int rowCount = 0;
			resultset = (int*)calloc(*pTotalRowCount, sizeof(int));
			for (int rowNo = 0; rowNo < *pTotalRowCount; rowNo++) {
				if (!bResultset[rowNo])
					continue;

				resultset[rowCount] = rowNo;
				rowCount++;
			}
			free(bResultset);

			if (rowCount > 0) {
				if (rowCount > *pTotalRowCount)
					MessageBeep(0);
				resultset = realloc(resultset, rowCount * sizeof(int));
				SetProp(hWnd, TEXT("RESULTSET"), (HANDLE)resultset);
				int orderBy = *pOrderBy;

				if (orderBy) {
					int colNo = orderBy > 0 ? orderBy - 1 : - orderBy - 1;
					BOOL isNum = TRUE;
					double* nums = 0;
					
					for (int i = isHeaderRow; i < *pTotalRowCount && i <= 5; i++) 
						isNum = isNum && isNumber(cache[i][colNo]);
						
					if (isNum) {
						nums = calloc(rowCount, sizeof(double));
						for (int i = 0; i < rowCount; i++)
							nums[i] = _tcstod(cache[resultset[i]][colNo], NULL);
					}	
				
					// Bubble-sort
					for (int i = 0; i < rowCount; i++) {
						for (int j = i + 1; j < rowCount; j++) {
							int a = resultset[i];
							int b = resultset[j];
							
							int cmp = isNum ? nums[i] - nums[j] : _tcscmp(cache[a][colNo], cache[b][colNo]);	
							if(orderBy > 0 && cmp > 0 || orderBy < 0 && cmp < 0) {
								resultset[i] = b;
								resultset[j] = a;
								
								if (isNum) { 
									double tmp = nums[i];
									nums[i] = nums[j];
									nums[j] = tmp;
								}
							}
						}
					}
					
					if (nums)
						free(nums);
				}
			} else {
				SetProp(hWnd, TEXT("RESULTSET"), (HANDLE)0);			
				free(resultset);
			}

			*pRowCount = rowCount;
			ListView_SetItemCount(hGridWnd, rowCount);
			InvalidateRect(hGridWnd, NULL, TRUE);

			TCHAR buf[255];
			_sntprintf(buf, 255, TEXT(" Rows: %i/%i"), rowCount, *pTotalRowCount - isHeaderRow);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_ROW_COUNT, (LPARAM)buf);

			PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;

		case WMU_UPDATE_FILTER_SIZE: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);
			SendMessage(hHeader, WM_SIZE, 0, 0);
			for (int colNo = 0; colNo < colCount; colNo++) {
				RECT rc;
				Header_GetItemRect(hHeader, colNo, &rc);
				int h2 = round((rc.bottom - rc.top) / 2);
				SetWindowPos(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), 0, rc.left - (colNo > 0), h2, rc.right - rc.left + 1, h2 + 1, SWP_NOZORDER);							
			}
		}
		break;

		case WMU_AUTO_COLUMN_SIZE: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);

			for (int colNo = 0; colNo < colCount - 1; colNo++)
				ListView_SetColumnWidth(hGridWnd, colNo, colNo < colCount - 1 ? LVSCW_AUTOSIZE_USEHEADER : LVSCW_AUTOSIZE);

			if (colCount == 1 && ListView_GetColumnWidth(hGridWnd, 0) < 100)
				ListView_SetColumnWidth(hGridWnd, 0, 100);
				
			int maxWidth = getStoredValue(TEXT("max-column-width"), 300);
			if (colCount > 1) {
				for (int colNo = 0; colNo < colCount; colNo++) {
					if (ListView_GetColumnWidth(hGridWnd, colNo) > maxWidth)
						ListView_SetColumnWidth(hGridWnd, colNo, maxWidth);
				}
			}

			// Fix last column				
			if (colCount > 1) {
				int colNo = colCount - 1;
				ListView_SetColumnWidth(hGridWnd, colNo, LVSCW_AUTOSIZE);
				TCHAR name16[MAX_COLUMN_LENGTH + 1];
				Header_GetItemText(hHeader, colNo, name16, MAX_COLUMN_LENGTH);
				
				SIZE s = {0};
				HDC hDC = GetDC(hHeader);
				HFONT hOldFont = (HFONT)SelectObject(hDC, (HFONT)GetProp(hWnd, TEXT("FONT")));
				GetTextExtentPoint32(hDC, name16, _tcslen(name16), &s);
				SelectObject(hDC, hOldFont);
				ReleaseDC(hHeader, hDC);

				int w = s.cx + 12;
				if (ListView_GetColumnWidth(hGridWnd, colNo) < w)
					ListView_SetColumnWidth(hGridWnd, colNo, w);
					
				if (ListView_GetColumnWidth(hGridWnd, colNo) > maxWidth)
					ListView_SetColumnWidth(hGridWnd, colNo, maxWidth);	
			}

			SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
			InvalidateRect(hGridWnd, NULL, TRUE);

			PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;

		case WMU_RESET_CACHE: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int* pRowCount = (int*)GetProp(hWnd, TEXT("ROWCOUNT"));

			int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
			if (colCount > 0 && cache != 0) {
				for (int rowNo = 0; rowNo < *pRowCount; rowNo++) {
					if (cache[rowNo]) {
						for (int colNo = 0; colNo < colCount; colNo++)
							if (cache[rowNo][colNo])
								free(cache[rowNo][colNo]);

						free(cache[rowNo]);
					}
					cache[rowNo] = 0;
				}
				free(cache);
			}

			SetProp(hWnd, TEXT("CACHE"), 0);
			*pRowCount = 0;
		}
		break;
		
		// wParam - size delta
		case WMU_SET_FONT: {
			int* pFontSize = (int*)GetProp(hWnd, TEXT("FONTSIZE"));
			if (*pFontSize + wParam < 10 || *pFontSize + wParam > 48)
				return 0;
			*pFontSize += wParam;
			DeleteFont(GetProp(hWnd, TEXT("FONT")));

			HFONT hFont = CreateFont (*pFontSize, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, (TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			SendMessage(hGridWnd, WM_SETFONT, (LPARAM)hFont, TRUE);

			HWND hHeader = ListView_GetHeader(hGridWnd);
			for (int colNo = 0; colNo < Header_GetItemCount(hHeader); colNo++)
				SendMessage(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), WM_SETFONT, (LPARAM)hFont, TRUE);

			SetProp(hWnd, TEXT("FONT"), hFont);
			PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
		}
		break;		
	}
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	WNDPROC cbDefault = (WNDPROC)GetProp(hWnd, TEXT("WNDPROC"));

	switch(msg){
		// Win10+ fix: draw an upper border
		case WM_PAINT: {
			cbDefault(hWnd, msg, wParam, lParam);

			RECT rc;
			GetWindowRect(hWnd, &rc);
			HDC hDC = GetWindowDC(hWnd);
			HPEN hPen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
			HPEN oldPen = SelectObject(hDC, hPen);
			MoveToEx(hDC, 1, 0, 0);
			LineTo(hDC, rc.right - 1, 0);
			SelectObject(hDC, oldPen);
			DeleteObject(hPen);
			ReleaseDC(hWnd, hDC);

			return 0;
		}
		break;

		case WM_CHAR: {
			return CallWindowProc(cbHotKey, hWnd, msg, wParam, lParam);
		}
		break;

		case WM_KEYDOWN: {
			if (wParam == VK_RETURN) {			
				HWND hHeader = GetParent(hWnd);
				HWND hGridWnd = GetParent(hHeader);
				HWND hMainWnd = GetParent(hGridWnd);
				SendMessage(hMainWnd, WMU_UPDATE_RESULTSET, 0, 0);
				
				return 0;
			}
			
			if (wParam == VK_TAB || wParam == VK_ESCAPE)
				return CallWindowProc(cbHotKey, hWnd, msg, wParam, lParam);
		}
		break;

		case WM_DESTROY: {
			RemoveProp(hWnd, TEXT("WNDPROC"));
		}
		break;
	}

	return CallWindowProc(cbDefault, hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_KEYDOWN && (wParam == VK_TAB || wParam == VK_ESCAPE)) {
		HWND hMainWnd = hWnd;
		while (hMainWnd && GetDlgCtrlID(hMainWnd) != IDC_MAIN)
			hMainWnd = GetParent(hMainWnd);
		SendMessage(hMainWnd, WM_KEYDOWN, wParam, lParam);
		return 0;
	}
	
	// Prevent beep
	if (msg == WM_CHAR && (wParam == VK_RETURN || wParam == VK_ESCAPE || wParam == VK_TAB))
		return 0;
	
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

void setStoredValue(TCHAR* name, int value) {
	TCHAR buf[128];
	_sntprintf(buf, 128, TEXT("%i"), value);
	WritePrivateProfileString(APP_NAME, name, buf, iniPath);
}

int getStoredValue(TCHAR* name, int defValue) {
	TCHAR buf[128];
	return GetPrivateProfileString(APP_NAME, name, NULL, buf, 128, iniPath) ? _ttoi(buf) : defValue;
}

TCHAR* getStoredString(TCHAR* name, TCHAR* defValue) { 
	TCHAR* buf = calloc(256, sizeof(TCHAR));
	if (0 == GetPrivateProfileString(APP_NAME, name, NULL, buf, 128, iniPath) && defValue)
		_tcsncpy(buf, defValue, 255);
	return buf;	
}

int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam) {
	if (GetWindowLong(hWnd, GWL_STYLE) & WS_TABSTOP && IsWindowVisible(hWnd)) {
		int no = 0;
		HWND* wnds = (HWND*)lParam;
		while (wnds[no])
			no++;
		wnds[no] = hWnd;
	}

	return TRUE;
}

TCHAR* utf8to16(const char* in) {
	TCHAR *out;
	if (!in || strlen(in) == 0) {
		out = (TCHAR*)calloc (1, sizeof (TCHAR));
	} else  {
		DWORD size = MultiByteToWideChar(CP_UTF8, 0, in, -1, NULL, 0);
		out = (TCHAR*)calloc (size, sizeof (TCHAR));
		MultiByteToWideChar(CP_UTF8, 0, in, -1, out, size);
	}
	return out;
}

char* utf16to8(const TCHAR* in) {
	char* out;
	if (!in || _tcslen(in) == 0) {
		out = (char*)calloc (1, sizeof(char));
	} else  {
		int len = WideCharToMultiByte(CP_UTF8, 0, in, -1, NULL, 0, 0, 0);
		out = (char*)calloc (len, sizeof(char));
		WideCharToMultiByte(CP_UTF8, 0, in, -1, out, len, 0, 0);
	}
	return out;
}

// https://stackoverflow.com/a/25023604/6121703
int detectCodePage(const unsigned char *data, int len) {
	int cp = 0;
	
	// BOM
	cp = strncmp(data, "\xEF\xBB\xBF", 3) == 0 ? CP_UTF8 :
		strncmp(data, "\xFE\xFF", 2) == 0 ? CP_UTF16BE :
		strncmp(data, "\xFF\xFE", 2) == 0 ? CP_UTF16LE :
		0;
	
	// By \n: UTF16LE - 0x0A 0x00, UTF16BE - 0x00 0x0A
	if (!cp) {
		int nCount = 0;
		int leadZeros = 0;
		int tailZeros = 0;
		for (int i = 0; i < len; i++) {
			if (data[i] != '\n')
				continue;
			
			nCount++;
			leadZeros += i > 0 && data[i - 1] == 0;	
			tailZeros += (i < len - 1) && data[i + 1] == 0;				
		}
		
		cp = nCount == 0 ? 0 :
			nCount == tailZeros ? CP_UTF16LE :
			nCount == leadZeros ? CP_UTF16BE :
			0;	
	}
	
	if (!cp && isUtf8(data)) 
		cp = CP_UTF8;	
	
	if (!cp)	
		cp = CP_ACP;
		
	return cp;			
}

TCHAR detectDelimiter(const TCHAR *data) {
	int rowCount = 2;
	int delimCount = _tcslen(DELIMITERS);
	int colCount[rowCount][delimCount];
	memset(colCount, 0, sizeof(int) * (size_t)(rowCount * delimCount));
	
	int pos = 0;
	int len = _tcslen(data);
	BOOL inQuote = FALSE;
	int rowNo = 0;
	for (; rowNo < rowCount && pos < len; rowNo++) {
		for (; pos < len; pos++) {
			TCHAR c = data[pos];
			if (!inQuote && c == TEXT('\n')) 
				break;

			if (c == TEXT('"'))
				inQuote = !inQuote;

			for (int delimNo = 0; delimNo < delimCount && !inQuote; delimNo++)
				colCount[rowNo][delimNo] += data[pos] == DELIMITERS[delimNo];
		}
		pos++;
	}

	TCHAR maxNo = 0;
	for (int delimNo = 1; delimNo < delimCount; delimNo++) {
		if (maxNo < colCount[0][delimNo] && (!rowNo || rowNo && colCount[0][delimNo] == colCount[1][delimNo]))
			maxNo = delimNo;
	}
	
	return DELIMITERS[maxNo];
}

void setClipboardText(const TCHAR* text) {
	int len = (_tcslen(text) + 1) * sizeof(TCHAR);
	HGLOBAL hMem =  GlobalAlloc(GMEM_MOVEABLE, len);
	memcpy(GlobalLock(hMem), text, len);
	GlobalUnlock(hMem);
	OpenClipboard(0);
	EmptyClipboard();
	SetClipboardData(CF_UNICODETEXT, hMem);
	CloseClipboard();
}

BOOL isNumber(TCHAR* val) {
	int len = _tcslen(val);
	BOOL res = TRUE;
	int pCount = 0;
	for (int i = 0; res && i < len; i++) {
		pCount += val[i] == TEXT('.');
		res = _istdigit(val[i]) || val[i] == TEXT('.');
	}
	return res && pCount < 2;
}

// https://stackoverflow.com/a/1031773/6121703
BOOL isUtf8(const char * string) {
	if (!string)
		return FALSE;

	const unsigned char * bytes = (const unsigned char *)string;
	while (*bytes) {
		if((bytes[0] == 0x09 || bytes[0] == 0x0A || bytes[0] == 0x0D || (0x20 <= bytes[0] && bytes[0] <= 0x7E))) {
			bytes += 1;
			continue;
		}

		if (((0xC2 <= bytes[0] && bytes[0] <= 0xDF) && (0x80 <= bytes[1] && bytes[1] <= 0xBF))) {
			bytes += 2;
			continue;
		}

		if ((bytes[0] == 0xE0 && (0xA0 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF)) ||
			(((0xE1 <= bytes[0] && bytes[0] <= 0xEC) || bytes[0] == 0xEE || bytes[0] == 0xEF) && (0x80 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF)) ||
			(bytes[0] == 0xED && (0x80 <= bytes[1] && bytes[1] <= 0x9F) && (0x80 <= bytes[2] && bytes[2] <= 0xBF))
		) {
			bytes += 3;
			continue;
		}

		if ((bytes[0] == 0xF0 && (0x90 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF) && (0x80 <= bytes[3] && bytes[3] <= 0xBF)) ||
			((0xF1 <= bytes[0] && bytes[0] <= 0xF3) && (0x80 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF) && (0x80 <= bytes[3] && bytes[3] <= 0xBF)) ||
			(bytes[0] == 0xF4 && (0x80 <= bytes[1] && bytes[1] <= 0x8F) && (0x80 <= bytes[2] && bytes[2] <= 0xBF) && (0x80 <= bytes[3] && bytes[3] <= 0xBF))
		) {
			bytes += 4;
			continue;
		}

		return FALSE;
	}

	return TRUE;
}

int ListView_AddColumn(HWND hListWnd, TCHAR* colName, int fmt) {
	int colNo = Header_GetItemCount(ListView_GetHeader(hListWnd));
	LVCOLUMN lvc = {0};
	lvc.mask = LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM | LVCF_FMT;
	lvc.iSubItem = colNo;
	lvc.pszText = colName;
	lvc.cchTextMax = _tcslen(colName) + 1;
	lvc.cx = 100;
	lvc.fmt = fmt;
	return ListView_InsertColumn(hListWnd, colNo, &lvc);
}

int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax) {
	if (i < 0)
		return FALSE;

	TCHAR buf[cchTextMax];

	HDITEM hdi = {0};
	hdi.mask = HDI_TEXT;
	hdi.pszText = buf;
	hdi.cchTextMax = cchTextMax;
	int rc = Header_GetItem(hWnd, i, &hdi);

	_tcsncpy(pszText, buf, cchTextMax);
	return rc;
}