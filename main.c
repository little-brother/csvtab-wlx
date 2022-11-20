#define UNICODE
#define _UNICODE

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <uxtheme.h>
#include <locale.h>
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
#define WMU_SET_HEADER_FILTERS WM_USER + 5
#define WMU_AUTO_COLUMN_SIZE   WM_USER + 6
#define WMU_SET_CURRENT_CELL   WM_USER + 7
#define WMU_RESET_CACHE        WM_USER + 8
#define WMU_SET_FONT           WM_USER + 9
#define WMU_SET_THEME          WM_USER + 10
#define WMU_HIDE_COLUMN        WM_USER + 11
#define WMU_SHOW_COLUMNS       WM_USER + 12
#define WMU_HOT_KEYS           WM_USER + 13  
#define WMU_HOT_CHARS          WM_USER + 14

#define IDC_MAIN               100
#define IDC_GRID               101
#define IDC_STATUSBAR          102
#define IDC_HEADER_EDIT        1000

#define IDM_COPY_CELL          5000
#define IDM_COPY_ROWS          5001
#define IDM_COPY_COLUMN        5002
#define IDM_FILTER_ROW         5003
#define IDM_HEADER_ROW         5004
#define IDM_DARK_THEME         5005
#define IDM_HIDE_COLUMN        5008

#define IDM_ANSI               5010
#define IDM_UTF8               5011
#define IDM_UTF16LE            5012
#define IDM_UTF16BE            5013

#define IDM_COMMA              5020
#define IDM_SEMICOLON          5021
#define IDM_VBAR               5022
#define IDM_TAB                5023
#define IDM_COLON              5024

#define IDM_DEFAULT            5030
#define IDM_NO_PARSE           5031
#define IDM_NO_SHOW            5032

#define SB_VERSION             0
#define SB_CODEPAGE            1
#define SB_DELIMITER           2
#define SB_COMMENTS            3
#define SB_ROW_COUNT           4
#define SB_CURRENT_CELL        5
#define SB_AUXILIARY           6

#define MAX_COLUMN_COUNT       128
#define MAX_COLUMN_LENGTH      2000
#define MAX_FILTER_LENGTH      2000
#define DELIMITERS             TEXT(",;|\t:")
#define APP_NAME               TEXT("csvtab")
#define APP_VERSION            TEXT("0.9.9")

#define CP_UTF16LE             1200
#define CP_UTF16BE             1201

#define LCS_FINDFIRST          1
#define LCS_MATCHCASE          2
#define LCS_WHOLEWORDS         4
#define LCS_BACKWARDS          8

typedef struct {
	int size;
	DWORD PluginInterfaceVersionLow;
	DWORD PluginInterfaceVersionHi;
	char DefaultIniName[MAX_PATH];
} ListDefaultParamStruct;

static TCHAR iniPath[MAX_PATH] = {0};

void __stdcall ListCloseWindow(HWND hWnd);
LRESULT CALLBACK cbNewMain (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewFilterEdit (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);

HWND getMainWindow(HWND hWnd);
void setStoredValue(TCHAR* name, int value);
int getStoredValue(TCHAR* name, int defValue);
TCHAR* getStoredString(TCHAR* name, TCHAR* defValue);
int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam);
TCHAR* utf8to16(const char* in);
char* utf16to8(const TCHAR* in);
int detectCodePage(const unsigned char *data, int len);
TCHAR detectDelimiter(const TCHAR *data, BOOL skipComments);
void setClipboardText(const TCHAR* text);
BOOL isEOL(TCHAR c);
BOOL isNumber(TCHAR* val);
BOOL isUtf8(const char * string);
int findString(TCHAR* text, TCHAR* word, BOOL isMatchCase, BOOL isWholeWords);
BOOL hasString (const TCHAR* str, const TCHAR* sub, BOOL isCaseSensitive);
TCHAR* extractUrl(TCHAR* data);
void mergeSort(int indexes[], void* data, int l, int r, BOOL isBackward, BOOL isNums);
int ListView_AddColumn(HWND hListWnd, TCHAR* colName, int fmt);
int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax);
void Menu_SetItemState(HMENU hMenu, UINT wID, UINT fState);

BOOL APIENTRY DllMain (HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
	if (ul_reason_for_call == DLL_PROCESS_ATTACH && iniPath[0] == 0) {
		TCHAR path[MAX_PATH + 1] = {0};
		GetModuleFileName(hModule, path, MAX_PATH);
		TCHAR* dot = _tcsrchr(path, TEXT('.'));
		_tcsncpy(dot, TEXT(".ini"), 5);
		if (_taccess(path, 0) == 0)
			_tcscpy(iniPath, path);	
	}
	return TRUE;
}

void __stdcall ListGetDetectString(char* DetectString, int maxlen) {
	TCHAR* detectString16 = getStoredString(TEXT("detect-string"), TEXT("MULTIMEDIA & (ext=\"CSV\" | ext=\"TAB\" | ext=\"TSV\")"));
	char* detectString8 = utf16to8(detectString16);
	snprintf(DetectString, maxlen, detectString8);
	free(detectString16);
	free(detectString8);
}

void __stdcall ListSetDefaultParams(ListDefaultParamStruct* dps) {
	if (iniPath[0] == 0) {
		DWORD size = MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, NULL, 0);
		MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, iniPath, size);
	}
}

int __stdcall ListSearchTextW(HWND hWnd, TCHAR* searchString, int searchParameter) {
	HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
	HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);	
	
	TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
	int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
	int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
	int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
	if (!resultset || rowCount == 0)
		return 0;

	BOOL isFindFirst = searchParameter & LCS_FINDFIRST;		
	BOOL isBackward = searchParameter & LCS_BACKWARDS;
	BOOL isMatchCase = searchParameter & LCS_MATCHCASE;
	BOOL isWholeWords = searchParameter & LCS_WHOLEWORDS;	

	if (isFindFirst) {
		*(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")) = 0;
		*(int*)GetProp(hWnd, TEXT("SEARCHCELLPOS")) = 0;	
		*(int*)GetProp(hWnd, TEXT("CURRENTROWNO")) = isBackward ? rowCount - 1 : 0;
	}	
		
	int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
	int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
	int *pStartPos = (int*)GetProp(hWnd, TEXT("SEARCHCELLPOS"));	
	rowNo = rowNo == -1 || rowNo >= rowCount ? 0 : rowNo;
	colNo = colNo == -1 || colNo >= colCount ? 0 : colNo;	
			
	int pos = -1;
	do {
		for (; (pos == -1) && colNo < colCount; colNo++) {
			pos = findString(cache[resultset[rowNo]][colNo] + *pStartPos, searchString, isMatchCase, isWholeWords);
			if (pos != -1) 
				pos += *pStartPos;
			*pStartPos = pos == -1 ? 0 : pos + _tcslen(searchString);
		}
		colNo = pos != -1 ? colNo - 1 : 0;
		rowNo += pos != -1 ? 0 : isBackward ? -1 : 1; 	
	} while ((pos == -1) && (isBackward ? rowNo > 0 : rowNo < rowCount));
	ListView_SetItemState(hGridWnd, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);

	TCHAR buf[256] = {0};
	if (pos != -1) {
		ListView_EnsureVisible(hGridWnd, rowNo, FALSE);
		ListView_SetItemState(hGridWnd, rowNo, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
		
		TCHAR* val = cache[resultset[rowNo]][colNo];
		int len = _tcslen(searchString);
		_sntprintf(buf, 255, TEXT("%ls%.*ls%ls"),
			pos > 0 ? TEXT("...") : TEXT(""), 
			len + pos + 10, val + pos,
			_tcslen(val + pos + len) > 10 ? TEXT("...") : TEXT(""));
		SendMessage(hWnd, WMU_SET_CURRENT_CELL, rowNo, colNo);
	} else { 
		MessageBox(hWnd, searchString, TEXT("Not found:"), MB_OK);
	}
	SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)buf);	
	SetFocus(hGridWnd);

	return 0;
}

int __stdcall ListSearchText(HWND hWnd, char* searchString, int searchParameter) {
	DWORD len = MultiByteToWideChar(CP_ACP, 0, searchString, -1, NULL, 0);
	TCHAR* searchString16 = (TCHAR*)calloc (len, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, searchString, -1, searchString16, len);
	int rc = ListSearchTextW(hWnd, searchString16, searchParameter);
	free(searchString16);
	return rc;
}	

HWND APIENTRY ListLoadW (HWND hListerWnd, TCHAR* fileToLoad, int showFlags) {		
	int size = _tcslen(fileToLoad);
	TCHAR* filepath = calloc(size + 1, sizeof(TCHAR));
	_tcsncpy(filepath, fileToLoad, size);
		
	int maxFileSize = getStoredValue(TEXT("max-file-size"), 10000000);	
	struct _stat st = {0};
	if (_tstat(filepath, &st) != 0 || st.st_size == 0 || maxFileSize > 0 && st.st_size > maxFileSize)
		return 0;

	INITCOMMONCONTROLSEX icex;
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_LISTVIEW_CLASSES;
	InitCommonControlsEx(&icex);
	
	setlocale(LC_CTYPE, "");

	BOOL isStandalone = GetParent(hListerWnd) == HWND_DESKTOP;
	HWND hMainWnd = CreateWindow(WC_STATIC, APP_NAME, WS_CHILD | (isStandalone ? SS_SUNKEN : 0),
		0, 0, 100, 100, hListerWnd, (HMENU)IDC_MAIN, GetModuleHandle(0), NULL);

	SetProp(hMainWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hMainWnd, GWLP_WNDPROC, (LONG_PTR)&cbNewMain));
	SetProp(hMainWnd, TEXT("FILEPATH"), filepath);
	SetProp(hMainWnd, TEXT("FILESIZE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("DELIMITER"), calloc(1, sizeof(TCHAR)));
	SetProp(hMainWnd, TEXT("CODEPAGE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FILTERROW"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("HEADERROW"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("SKIPCOMMENTS"), calloc(1, sizeof(int)));		
	SetProp(hMainWnd, TEXT("CACHE"), 0);
	SetProp(hMainWnd, TEXT("RESULTSET"), 0);
	SetProp(hMainWnd, TEXT("ORDERBY"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("COLCOUNT"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("ROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("TOTALROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("CURRENTROWNO"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("CURRENTCOLNO"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SEARCHCELLPOS"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FONT"), 0);
	SetProp(hMainWnd, TEXT("FONTFAMILY"), getStoredString(TEXT("font"), TEXT("Arial")));
	SetProp(hMainWnd, TEXT("FONTSIZE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FILTERALIGN"), calloc(1, sizeof(int)));	
	
	SetProp(hMainWnd, TEXT("DARKTHEME"), calloc(1, sizeof(int)));			
	SetProp(hMainWnd, TEXT("TEXTCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("BACKCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("BACKCOLOR2"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("FILTERTEXTCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FILTERBACKCOLOR"), calloc(1, sizeof(int)));		
	SetProp(hMainWnd, TEXT("CURRENTCELLCOLOR"), calloc(1, sizeof(int)));	
	
	*(int*)GetProp(hMainWnd, TEXT("FILESIZE")) = st.st_size;
	*(int*)GetProp(hMainWnd, TEXT("CODEPAGE")) = -1;
	*(int*)GetProp(hMainWnd, TEXT("FONTSIZE")) = getStoredValue(TEXT("font-size"), 16);	
	*(int*)GetProp(hMainWnd, TEXT("FILTERROW")) = getStoredValue(TEXT("filter-row"), 1);
	*(int*)GetProp(hMainWnd, TEXT("HEADERROW")) = getStoredValue(TEXT("header-row"), 1);	
	*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) = getStoredValue(TEXT("dark-theme"), 0);	
	*(int*)GetProp(hMainWnd, TEXT("SKIPCOMMENTS")) = getStoredValue(TEXT("skip-comments"), 0);
	*(int*)GetProp(hMainWnd, TEXT("FILTERALIGN")) = getStoredValue(TEXT("filter-align"), 0);
		
	HWND hStatusWnd = CreateStatusWindow(WS_CHILD | WS_VISIBLE | SBT_TOOLTIPS | (isStandalone ? SBARS_SIZEGRIP : 0), NULL, hMainWnd, IDC_STATUSBAR);
	HDC hDC = GetDC(hMainWnd);
	float z = GetDeviceCaps(hDC, LOGPIXELSX) / 96.0; // 96 = 100%, 120 = 125%, 144 = 150%
	ReleaseDC(hMainWnd, hDC);	
	int sizes[7] = {35 * z, 95 * z, 125 * z, 155 * z, 275 * z, 355 * z, -1};
	SendMessage(hStatusWnd, SB_SETPARTS, 7, (LPARAM)&sizes);
	TCHAR buf[32];
	_sntprintf(buf, 32, TEXT(" %ls"), APP_VERSION);
	SendMessage(hStatusWnd, SB_SETTEXT, SB_VERSION, (LPARAM)buf);
	SendMessage(hStatusWnd, SB_SETTIPTEXT, SB_COMMENTS, (LPARAM)TEXT("How to process lines starting with #. Check Wiki to get details."));

	HWND hGridWnd = CreateWindow(WC_LISTVIEW, NULL, WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA | WS_TABSTOP,
		205, 0, 100, 100, hMainWnd, (HMENU)IDC_GRID, GetModuleHandle(0), NULL);
		
	int noLines = getStoredValue(TEXT("disable-grid-lines"), 0);	
	ListView_SetExtendedListViewStyle(hGridWnd, LVS_EX_FULLROWSELECT | (noLines ? 0 : LVS_EX_GRIDLINES) | LVS_EX_LABELTIP);
	SetProp(hGridWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hGridWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));

	HWND hHeader = ListView_GetHeader(hGridWnd);
	SetWindowTheme(hHeader, TEXT(" "), TEXT(" "));
	SetProp(hHeader, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hHeader, GWLP_WNDPROC, (LONG_PTR)cbNewHeader));	 

	HMENU hGridMenu = CreatePopupMenu();
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_CELL, TEXT("Copy cell"));
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_ROWS, TEXT("Copy row(s)"));
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_COLUMN, TEXT("Copy column"));	
	AppendMenu(hGridMenu, MF_STRING, 0, NULL);
	AppendMenu(hGridMenu, MF_STRING, IDM_HIDE_COLUMN, TEXT("Hide column"));	
	AppendMenu(hGridMenu, MF_STRING, 0, NULL);	
	AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("FILTERROW")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_FILTER_ROW, TEXT("Filters"));		
	AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("HEADERROW")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_HEADER_ROW, TEXT("Header row"));	
	AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_DARK_THEME, TEXT("Dark theme"));		
	SetProp(hMainWnd, TEXT("GRIDMENU"), hGridMenu);

	HMENU hCodepageMenu = CreatePopupMenu();
	AppendMenu(hCodepageMenu, MF_STRING, IDM_ANSI, TEXT("ANSI"));
	AppendMenu(hCodepageMenu, MF_STRING, IDM_UTF8, TEXT("UTF-8"));
	AppendMenu(hCodepageMenu, MF_STRING, IDM_UTF16LE, TEXT("UTF-16LE"));
	AppendMenu(hCodepageMenu, MF_STRING, IDM_UTF16BE, TEXT("UTF-16BE"));	
	SetProp(hMainWnd, TEXT("CODEPAGEMENU"), hCodepageMenu);

	HMENU hDelimiterMenu = CreatePopupMenu();
	AppendMenu(hDelimiterMenu, MF_STRING, IDM_COMMA, TEXT(","));
	AppendMenu(hDelimiterMenu, MF_STRING, IDM_SEMICOLON, TEXT(";"));
	AppendMenu(hDelimiterMenu, MF_STRING, IDM_VBAR, TEXT("|"));
	AppendMenu(hDelimiterMenu, MF_STRING, IDM_TAB, TEXT("TAB"));	
	AppendMenu(hDelimiterMenu, MF_STRING, IDM_COLON, TEXT(":"));	
	SetProp(hMainWnd, TEXT("DELIMITERMENU"), hDelimiterMenu);
	
	HMENU hCommentMenu = CreatePopupMenu();
	AppendMenu(hCommentMenu, MF_STRING, IDM_DEFAULT, TEXT("#0 - No special"));
	AppendMenu(hCommentMenu, MF_STRING, IDM_NO_PARSE, TEXT("#1 - Don't parse"));
	AppendMenu(hCommentMenu, MF_STRING, IDM_NO_SHOW, TEXT("#2 - Don't show"));
	SetProp(hMainWnd, TEXT("COMMENTMENU"), hCommentMenu);	

	SendMessage(hMainWnd, WMU_SET_FONT, 0, 0);
	SendMessage(hMainWnd, WMU_SET_THEME, 0, 0);	
	SendMessage(hMainWnd, WMU_UPDATE_CACHE, 0, 0);
	ShowWindow(hMainWnd, SW_SHOW);
	SetFocus(hGridWnd);
	
	return hMainWnd;
}

HWND APIENTRY ListLoad (HWND hListerWnd, char* fileToLoad, int showFlags) {
	DWORD size = MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, NULL, 0);
	TCHAR* fileToLoadW = (TCHAR*)calloc (size, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, fileToLoadW, size);
	HWND hWnd = ListLoadW(hListerWnd, fileToLoadW, showFlags);
	free(fileToLoadW);
	return hWnd;
}

void __stdcall ListCloseWindow(HWND hWnd) {
	setStoredValue(TEXT("font-size"), *(int*)GetProp(hWnd, TEXT("FONTSIZE")));
	setStoredValue(TEXT("filter-row"), *(int*)GetProp(hWnd, TEXT("FILTERROW")));		
	setStoredValue(TEXT("header-row"), *(int*)GetProp(hWnd, TEXT("HEADERROW")));	
	setStoredValue(TEXT("dark-theme"), *(int*)GetProp(hWnd, TEXT("DARKTHEME")));
	setStoredValue(TEXT("skip-comments"), *(int*)GetProp(hWnd, TEXT("SKIPCOMMENTS")));
	
	SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
	free((TCHAR*)GetProp(hWnd, TEXT("FILEPATH")));
	free((int*)GetProp(hWnd, TEXT("FILESIZE")));	
	free((TCHAR*)GetProp(hWnd, TEXT("DELIMITER")));
	free((int*)GetProp(hWnd, TEXT("CODEPAGE")));
	free((int*)GetProp(hWnd, TEXT("FILTERROW")));
	free((int*)GetProp(hWnd, TEXT("HEADERROW")));	
	free((int*)GetProp(hWnd, TEXT("DARKTHEME")));		
	free((int*)GetProp(hWnd, TEXT("ORDERBY")));
	free((int*)GetProp(hWnd, TEXT("COLCOUNT")));	
	free((int*)GetProp(hWnd, TEXT("ROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("FONTSIZE")));
	free((int*)GetProp(hWnd, TEXT("CURRENTROWNO")));				
	free((int*)GetProp(hWnd, TEXT("CURRENTCOLNO")));
	free((int*)GetProp(hWnd, TEXT("SEARCHCELLPOS")));	
	free((TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
	free((int*)GetProp(hWnd, TEXT("FILTERALIGN")));	
		
	free((int*)GetProp(hWnd, TEXT("TEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("BACKCOLOR")));
	free((int*)GetProp(hWnd, TEXT("BACKCOLOR2")));	
	free((int*)GetProp(hWnd, TEXT("FILTERTEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("FILTERBACKCOLOR")));
	free((int*)GetProp(hWnd, TEXT("CURRENTCELLCOLOR")));	
		
	DeleteFont(GetProp(hWnd, TEXT("FONT")));
	DeleteObject(GetProp(hWnd, TEXT("BACKBRUSH")));	
	DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));
	DestroyMenu(GetProp(hWnd, TEXT("GRIDMENU")));
	DestroyMenu(GetProp(hWnd, TEXT("CODEPAGEMENU")));	
	DestroyMenu(GetProp(hWnd, TEXT("DELIMITERMENU")));	
	DestroyMenu(GetProp(hWnd, TEXT("COMMENTMENU")));	

	RemoveProp(hWnd, TEXT("WNDPROC"));
	RemoveProp(hWnd, TEXT("CACHE"));
	RemoveProp(hWnd, TEXT("RESULTSET"));
	RemoveProp(hWnd, TEXT("FILEPATH"));
	RemoveProp(hWnd, TEXT("FILESIZE"));
	RemoveProp(hWnd, TEXT("DELIMITER"));
	RemoveProp(hWnd, TEXT("CODEPAGE"));
	RemoveProp(hWnd, TEXT("FILTERROW"));	
	RemoveProp(hWnd, TEXT("HEADERROW"));	
	RemoveProp(hWnd, TEXT("DARKTHEME"));	
	RemoveProp(hWnd, TEXT("ORDERBY"));
	RemoveProp(hWnd, TEXT("COLCOUNT"));
	RemoveProp(hWnd, TEXT("ROWCOUNT"));
	RemoveProp(hWnd, TEXT("TOTALROWCOUNT"));
	RemoveProp(hWnd, TEXT("CURRENTROWNO"));	
	RemoveProp(hWnd, TEXT("CURRENTCOLNO"));
	RemoveProp(hWnd, TEXT("SEARCHCELLPOS"));	
	RemoveProp(hWnd, TEXT("FILTERALIGN"));	
	RemoveProp(hWnd, TEXT("LASTFOCUS"));	
	
	RemoveProp(hWnd, TEXT("TEXTCOLOR"));
	RemoveProp(hWnd, TEXT("BACKCOLOR"));
	RemoveProp(hWnd, TEXT("BACKCOLOR2"));	
	RemoveProp(hWnd, TEXT("FILTERTEXTCOLOR"));
	RemoveProp(hWnd, TEXT("FILTERBACKCOLOR"));	
	RemoveProp(hWnd, TEXT("CURRENTCELLCOLOR"));	
	RemoveProp(hWnd, TEXT("BACKBRUSH"));
	RemoveProp(hWnd, TEXT("FILTERBACKBRUSH"));	
	RemoveProp(hWnd, TEXT("FONT"));
	RemoveProp(hWnd, TEXT("FONTFAMILY"));
	RemoveProp(hWnd, TEXT("FONTSIZE"));	
	RemoveProp(hWnd, TEXT("GRIDMENU"));
	RemoveProp(hWnd, TEXT("CODEPAGEMENU"));	
	RemoveProp(hWnd, TEXT("DELIMITERMENU"));	
	RemoveProp(hWnd, TEXT("COMMENTMENU"));	

	DestroyWindow(hWnd);
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
		
		case WM_SETCURSOR: {
			SetCursor(LoadCursor(0, IDC_ARROW));
			return TRUE;
		}
		break;
				
		case WM_SETFOCUS: {
			SetFocus(GetProp(hWnd, TEXT("LASTFOCUS")));
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
			if (SendMessage(hWnd, WMU_HOT_KEYS, wParam, lParam))
				return 0;			
		}
		break;
				
		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);
			if (cmd == IDM_COPY_CELL || cmd == IDM_COPY_ROWS || cmd == IDM_COPY_COLUMN) {
				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
				int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
				
				int colCount = Header_GetItemCount(hHeader);
				int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
				int selCount = ListView_GetSelectedCount(hGridWnd);

				if (rowNo == -1 ||
					rowNo >= rowCount ||
					colCount == 0 ||
					cmd == IDM_COPY_CELL && colNo == -1 || 
					cmd == IDM_COPY_CELL && colNo >= colCount || 
					cmd == IDM_COPY_COLUMN && colNo == -1 || 
					cmd == IDM_COPY_COLUMN && colNo >= colCount || 					
					cmd == IDM_COPY_ROWS && selCount == 0) {
					setClipboardText(TEXT(""));
					return 0;
				}
						
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));				
				if (!resultset)
					return 0;
					
				TCHAR* delimiter = getStoredString(TEXT("column-delimiter"), TEXT("\t"));
				
				int len = 0;
				if (cmd == IDM_COPY_CELL) 
					len = _tcslen(cache[resultset[rowNo]][colNo]);
				
				if (cmd == IDM_COPY_ROWS) {
					int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1) {
						for (int colNo = 0; colNo < colCount; colNo++) {
							if (ListView_GetColumnWidth(hGridWnd, colNo)) 
								len += _tcslen(cache[resultset[rowNo]][colNo]) + 1; /* column delimiter */
						}
													
						len++; /* \n */		
						rowNo = ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);	
					}
				}

				if (cmd == IDM_COPY_COLUMN) {
					int rowNo = selCount < 2 ? 0 : ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1 && rowNo < rowCount) {
						len += _tcslen(cache[resultset[rowNo]][colNo]) + 1 /* \n */;
						rowNo = selCount < 2 ? rowNo + 1 : ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);
					} 
				}	

				TCHAR* buf = calloc(len + 1, sizeof(TCHAR));
				if (cmd == IDM_COPY_CELL)
					_tcscat(buf, cache[resultset[rowNo]][colNo]);
				
				if (cmd == IDM_COPY_ROWS) {
					int pos = 0;
					int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1) {
						for (int colNo = 0; colNo < colCount; colNo++) {
							if (ListView_GetColumnWidth(hGridWnd, colNo)) {
								int len = _tcslen(cache[resultset[rowNo]][colNo]);
								_tcsncpy(buf + pos, cache[resultset[rowNo]][colNo], len);
								buf[pos + len] = delimiter[0];
								pos += len + 1;
							}
						}

						buf[pos - (pos > 0)] = TEXT('\n');
						rowNo = ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);	
					}
					buf[pos - 1] = 0; // remove last \n
				}

				if (cmd == IDM_COPY_COLUMN) {
					int pos = 0;
					int rowNo = selCount < 2 ? 0 : ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1 && rowNo < rowCount) {
						int len = _tcslen(cache[resultset[rowNo]][colNo]);
						_tcsncpy(buf + pos, cache[resultset[rowNo]][colNo], len);
						rowNo = selCount < 2 ? rowNo + 1 : ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);
						if (rowNo != -1 && rowNo < rowCount)
							buf[pos + len] = TEXT('\n');
						pos += len + 1;								
					} 
				}
									
				setClipboardText(buf);
				free(buf);
				free(delimiter);
			}
			
			if (cmd == IDM_HIDE_COLUMN) {
				int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
				SendMessage(hWnd, WMU_HIDE_COLUMN, colNo, 0);
			}					
			
			if (cmd == IDM_FILTER_ROW || cmd == IDM_HEADER_ROW || cmd == IDM_DARK_THEME) {
				HMENU hMenu = (HMENU)GetProp(hWnd, TEXT("GRIDMENU"));
				int* pOpt = (int*)GetProp(hWnd, cmd == IDM_FILTER_ROW ? TEXT("FILTERROW") : cmd == IDM_HEADER_ROW ? TEXT("HEADERROW" : TEXT("DARKTHEME")));
				*pOpt = (*pOpt + 1) % 2;
				Menu_SetItemState(hMenu, cmd, *pOpt ? MFS_CHECKED : 0);
				
				UINT msg = cmd == IDM_FILTER_ROW ? WMU_SET_HEADER_FILTERS : cmd == IDM_HEADER_ROW ? WMU_UPDATE_GRID : WMU_SET_THEME;
				SendMessage(hWnd, msg, 0, 0);				
			}

			if (cmd == IDM_ANSI || cmd == IDM_UTF8 || cmd == IDM_UTF16LE || cmd == IDM_UTF16BE) {
				int* pCodepage = (int*)GetProp(hWnd, TEXT("CODEPAGE"));
				*pCodepage = cmd == IDM_ANSI ? CP_ACP : cmd == IDM_UTF8 ? CP_UTF8 : cmd == IDM_UTF16LE ? CP_UTF16LE : CP_UTF16BE;
				SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
			}
			
			if (cmd == IDM_COMMA || cmd == IDM_SEMICOLON || cmd == IDM_VBAR || cmd == IDM_TAB || cmd == IDM_COLON) {
				TCHAR* pDelimiter = (TCHAR*)GetProp(hWnd, TEXT("DELIMITER"));
				*pDelimiter = cmd == IDM_COMMA ? TEXT(',') : 
					cmd == IDM_SEMICOLON ? TEXT(';') : 
					cmd == IDM_TAB ? TEXT('\t') : 
					cmd == IDM_VBAR ? TEXT('|') : 
					cmd == IDM_COLON ? TEXT(':') : 
					0;
				SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
			}
			
			if (cmd == IDM_DEFAULT || cmd == IDM_NO_PARSE || cmd == IDM_NO_SHOW) {
				int* pSkipComments = (int*)GetProp(hWnd, TEXT("SKIPCOMMENTS"));
				*pSkipComments = cmd == IDM_DEFAULT ? 0 : cmd == IDM_NO_PARSE ? 1 : 2;
				SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
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
				NMLISTVIEW* lv = (NMLISTVIEW*)lParam;
				if (HIWORD(GetKeyState(VK_CONTROL))) 
					return SendMessage(hWnd, WMU_HIDE_COLUMN, lv->iSubItem, 0);

				int colNo = lv->iSubItem + 1;
				int* pOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));
				int orderBy = *pOrderBy;
				*pOrderBy = colNo == orderBy || colNo == -orderBy ? -orderBy : colNo;
				SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);				
			}

			if (pHdr->idFrom == IDC_GRID && (pHdr->code == (DWORD)NM_CLICK || pHdr->code == (DWORD)NM_RCLICK)) {
				NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;
				SendMessage(hWnd, WMU_SET_CURRENT_CELL, ia->iItem, ia->iSubItem);
			}
			
			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_CLICK && HIWORD(GetKeyState(VK_MENU))) {	
				NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
				
				TCHAR* url = extractUrl(cache[resultset[ia->iItem]][ia->iSubItem]);
				ShellExecute(0, TEXT("open"), url, 0, 0 , SW_SHOW);
				free(url);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_RCLICK) {
				POINT p;
				GetCursorPos(&p);
				TrackPopupMenu(GetProp(hWnd, TEXT("GRIDMENU")), TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
			}
			
			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_ITEMCHANGED) {
				NMLISTVIEW* lv = (NMLISTVIEW*)lParam;
				if (lv->uOldState != lv->uNewState && (lv->uNewState & LVIS_SELECTED))				
					SendMessage(hWnd, WMU_SET_CURRENT_CELL, lv->iItem, *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")));	
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_KEYDOWN) {
				NMLVKEYDOWN* kd = (LPNMLVKEYDOWN) lParam;
				if (kd->wVKey == 0x43) { // C
					BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
					BOOL isShift = HIWORD(GetKeyState(VK_SHIFT));
					BOOL isCopyColumn = getStoredValue(TEXT("copy-column"), 0) && ListView_GetSelectedCount(pHdr->hwndFrom) > 1;
					if (!isCtrl && !isShift)
						return FALSE;
						
					int action = !isShift && !isCopyColumn ? IDM_COPY_CELL : isCtrl || isCopyColumn ? IDM_COPY_COLUMN : IDM_COPY_ROWS;
					SendMessage(hWnd, WM_COMMAND, action, 0);
					return TRUE;
				}

				if (kd->wVKey == 0x41 && HIWORD(GetKeyState(VK_CONTROL))) { // Ctrl + A
					HWND hGridWnd = pHdr->hwndFrom;
					SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
					int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
					int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));					
					ListView_SetItemState(hGridWnd, -1, LVIS_SELECTED, LVIS_SELECTED | LVIS_FOCUSED);
					SendMessage(hWnd, WMU_SET_CURRENT_CELL, rowNo, colNo);
					SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
					InvalidateRect(hGridWnd, NULL, TRUE);
				}

				if (kd->wVKey == 0x20 && HIWORD(GetKeyState(VK_CONTROL))) { // Ctrl + Space				
					SendMessage(hWnd, WMU_SHOW_COLUMNS, 0, 0);										
					return TRUE;
				}
				
				if (kd->wVKey == VK_LEFT || kd->wVKey == VK_RIGHT) {
					int colCount = Header_GetItemCount(ListView_GetHeader(pHdr->hwndFrom));
					int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")) + (kd->wVKey == VK_RIGHT ? 1 : -1);
					colNo = colNo < 0 ? colCount - 1 : colNo > colCount - 1 ? 0 : colNo;
					SendMessage(hWnd, WMU_SET_CURRENT_CELL, *(int*)GetProp(hWnd, TEXT("CURRENTROWNO")), colNo);
					return TRUE;
				}
			}

			if (pHdr->code == HDN_ITEMCHANGED && pHdr->hwndFrom == ListView_GetHeader(GetDlgItem(hWnd, IDC_GRID)))
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
				
			if (pHdr->idFrom == IDC_STATUSBAR && (pHdr->code == NM_CLICK || pHdr->code == NM_RCLICK)) {
				NMMOUSE* pm = (NMMOUSE*)lParam;
				int id = pm->dwItemSpec;
				if (id != SB_CODEPAGE && id != SB_DELIMITER && id != SB_COMMENTS)
					return 0;
					
				RECT rc, rc2;
				GetWindowRect(pHdr->hwndFrom, &rc);
				SendMessage(pHdr->hwndFrom, SB_GETRECT, id, (LPARAM)&rc2);
				HMENU hMenu = GetProp(hWnd, id == SB_CODEPAGE ? TEXT("CODEPAGEMENU") : id == SB_DELIMITER ? TEXT("DELIMITERMENU") : TEXT("COMMENTMENU"));
				
				if (id == SB_CODEPAGE) {
					int codepage = *(int*)GetProp(hWnd, TEXT("CODEPAGE"));
					Menu_SetItemState(hMenu, IDM_ANSI, codepage == CP_ACP ? MFS_CHECKED : 0);
					Menu_SetItemState(hMenu, IDM_UTF8, codepage == CP_UTF8 ? MFS_CHECKED : 0);					
					Menu_SetItemState(hMenu, IDM_UTF16LE, codepage == CP_UTF16LE ? MFS_CHECKED : 0);
					Menu_SetItemState(hMenu, IDM_UTF16BE, codepage == CP_UTF16BE ? MFS_CHECKED : 0);					
				} 
				
				if (id == SB_DELIMITER) {
					TCHAR delimiter = *(TCHAR*)GetProp(hWnd, TEXT("DELIMITER"));
					Menu_SetItemState(hMenu, IDM_COMMA, delimiter == TEXT(',') ? MFS_CHECKED : 0);
					Menu_SetItemState(hMenu, IDM_SEMICOLON, delimiter == TEXT(';') ? MFS_CHECKED : 0);
					Menu_SetItemState(hMenu, IDM_VBAR, delimiter == TEXT('|') ? MFS_CHECKED : 0);
					Menu_SetItemState(hMenu, IDM_TAB, delimiter == TEXT('\t') ? MFS_CHECKED : 0);
					Menu_SetItemState(hMenu, IDM_COLON, delimiter == TEXT(':') ? MFS_CHECKED : 0);					
				}

				if (id == SB_COMMENTS) {
					int skipComments = *(int*)GetProp(hWnd, TEXT("SKIPCOMMENTS"));
					Menu_SetItemState(hMenu, IDM_DEFAULT, skipComments == 0 ? MFS_CHECKED : 0);
					Menu_SetItemState(hMenu, IDM_NO_PARSE, skipComments == 1 ? MFS_CHECKED : 0);
					Menu_SetItemState(hMenu, IDM_NO_SHOW, skipComments == 2 ? MFS_CHECKED : 0);
				}
				
				POINT p = {rc.left + rc2.left, rc.top};				
				TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_BOTTOMALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
			}
			
			if (pHdr->code == (UINT)NM_SETFOCUS) 
				SetProp(hWnd, TEXT("LASTFOCUS"), pHdr->hwndFrom);
			
			if (pHdr->idFrom == IDC_GRID && pHdr->code == (UINT)NM_CUSTOMDRAW) {
				int result = CDRF_DODEFAULT;
	
				NMLVCUSTOMDRAW* pCustomDraw = (LPNMLVCUSTOMDRAW)lParam;
				if (pCustomDraw->nmcd.dwDrawStage == CDDS_PREPAINT)
					result = CDRF_NOTIFYITEMDRAW;
	
				if (pCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)
					result = CDRF_NOTIFYSUBITEMDRAW | CDRF_NEWFONT;
	
				if (pCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM)) {
					pCustomDraw->clrTextBk = pCustomDraw->nmcd.dwItemSpec % 2 == 0 ? *(int*)GetProp(hWnd, TEXT("BACKCOLOR")) : *(int*)GetProp(hWnd, TEXT("BACKCOLOR2"));
					result = CDRF_NOTIFYPOSTPAINT;
				}
	
				if ((pCustomDraw->nmcd.dwDrawStage == CDDS_POSTPAINT) | CDDS_SUBITEM) {
					int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
					int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
					if ((pCustomDraw->nmcd.dwItemSpec == (DWORD)rowNo) && (pCustomDraw->iSubItem == colNo)) {
						HPEN hPen = CreatePen(PS_DOT, 1, *(int*)GetProp(hWnd, TEXT("CURRENTCELLCOLOR")));
						HDC hDC = pCustomDraw->nmcd.hdc;
						SelectObject(hDC, hPen);
	
						RECT rc = {0};
						ListView_GetSubItemRect(pHdr->hwndFrom, rowNo, colNo, LVIR_BOUNDS, &rc);
						if (colNo == 0) 
							rc.right = ListView_GetColumnWidth(pHdr->hwndFrom, 0);

						MoveToEx(hDC, rc.left - 2, rc.top + 1, 0);
						LineTo(hDC, rc.right - 1, rc.top + 1);
						LineTo(hDC, rc.right - 1, rc.bottom - 2);
						LineTo(hDC, rc.left + 1, rc.bottom - 2);
						LineTo(hDC, rc.left + 1, rc.top);
						DeleteObject(hPen);
					}
				}
	
				return result;
			}
		}
		break;
		
		// wParam = colNo
		case WMU_HIDE_COLUMN: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);		
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colNo = (int)wParam;

			HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
			SetWindowLongPtr(hEdit, GWLP_USERDATA, (LONG_PTR)ListView_GetColumnWidth(hGridWnd, colNo));				
			ListView_SetColumnWidth(hGridWnd, colNo, 0); 
			InvalidateRect(hHeader, NULL, TRUE);
		}
		break;
		
		case WMU_SHOW_COLUMNS: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
			for (int colNo = 0; colNo < colCount; colNo++) {
				if (ListView_GetColumnWidth(hGridWnd, colNo) == 0) {
					HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
					ListView_SetColumnWidth(hGridWnd, colNo, (int)GetWindowLongPtr(hEdit, GWLP_USERDATA));
				}
			}

			InvalidateRect(hGridWnd, NULL, TRUE);		
		}
		break;		
		
		case WMU_UPDATE_CACHE: {
			TCHAR* filepath = (TCHAR*)GetProp(hWnd, TEXT("FILEPATH"));
			int filesize = *(int*)GetProp(hWnd, TEXT("FILESIZE"));
			TCHAR delimiter = *(TCHAR*)GetProp(hWnd, TEXT("DELIMITER"));
			int codepage = *(int*)GetProp(hWnd, TEXT("CODEPAGE"));
			int skipComments = *(int*)GetProp(hWnd, TEXT("SKIPCOMMENTS"));					
			
			SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);

			char* rawdata = calloc(filesize + 2, sizeof(char)); // + 2!
			FILE *f = _tfopen(filepath, TEXT("rb"));
			fread(rawdata, sizeof(char), filesize, f);
			fclose(f);
			rawdata[filesize] = 0;
			rawdata[filesize + 1] = 0;
			
			int leadZeros = 0;
			for (int i = 0; leadZeros == i && i < filesize; i++)
				leadZeros += rawdata[i] == 0;
			
			if (leadZeros == filesize) {
				PostMessage(GetParent(hWnd), WM_KEYDOWN, 0x33, 0x20001); // Switch to Hex-mode
				
				// A special TC-command doesn't work under TC x32. 
				// https://flint-inc.ru/tcinfo/all_cmd.ru.htm#Misc
				// PostMessage(GetAncestor(hWnd, GA_ROOT), WM_USER + 51, 4005, 0);
				keybd_event(VK_TAB, VK_TAB, KEYEVENTF_EXTENDEDKEY, 0);

				return FALSE;
			}
			
			if (codepage == -1)
				codepage = detectCodePage(rawdata, filesize);
				
			// Fix unexpected zeros
			if (codepage == CP_UTF16BE || codepage == CP_UTF16LE) {
				for (int i = 0; i < filesize / 2; i++) {
					if (rawdata[2 * i] == 0 && rawdata[2 * i + 1] == 0)
						rawdata[2 * i + codepage == CP_UTF16LE] = ' ';
				}
			}
			
			if (codepage == CP_UTF8 || codepage == CP_ACP) {
				for (int i = 0; i < filesize; i++) {
					if (rawdata[i] == 0)
						rawdata[i] = ' ';
				}
			}
			// end fix
			
			TCHAR* data = 0;				
			if (codepage == CP_UTF16BE) {
				for (int i = 0; i < filesize/2; i++) {
					int c = rawdata[2 * i];
					rawdata[2 * i] = rawdata[2 * i + 1];
					rawdata[2 * i + 1] = c;
				}
			}
			
			if (codepage == CP_UTF16LE || codepage == CP_UTF16BE)
				data = (TCHAR*)(rawdata + leadZeros);
				
			if (codepage == CP_UTF8) {
				data = utf8to16(rawdata + leadZeros);
				free(rawdata);
			}
			
			if (codepage == CP_ACP) {
				DWORD len = MultiByteToWideChar(CP_ACP, 0, rawdata, -1, NULL, 0);
				data = (TCHAR*)calloc (len, sizeof (TCHAR));
				if (!MultiByteToWideChar(CP_ACP, 0, rawdata + leadZeros, -1, data, len))
					codepage = -1;
				free(rawdata);
			}
			
			if (codepage == -1) {
				MessageBox(hWnd, TEXT("Can't detect codepage"), NULL, MB_OK);
				return 0;
			}
			
			if (!delimiter)
				delimiter = detectDelimiter(data, skipComments);
										
			int colCount = 1;				
			int rowNo = -1;
			int len = _tcslen(data);
			int cacheSize = len / 100 + 1;
			TCHAR*** cache = calloc(cacheSize, sizeof(TCHAR**));
			BOOL isTrim = getStoredValue(TEXT("trim-values"), 1);

			// Two step parsing: 0 - count columns, 1 - fill cache
			for (int stepNo = 0; stepNo < 2; stepNo++) {					
				rowNo = -1;
				int start = 0;			
				for(int pos = 0; pos < len; pos++) {
					rowNo++;
					if (stepNo == 1) {
						if (rowNo >= cacheSize) {
							cacheSize += 100;
							cache = realloc(cache, cacheSize * sizeof(TCHAR**));
						}

						cache[rowNo] = (TCHAR**)calloc (colCount, sizeof (TCHAR*));
					}
					
					int colNo = 0;
					start = pos;
					for (; pos < len && colNo < MAX_COLUMN_COUNT; pos++) {
						TCHAR c = data[pos];
						
						while (start == pos && skipComments == 2 && c == TEXT('#')) {
							while (data[pos] && !isEOL(data[pos]))
								pos++;
							while (pos < len && isEOL(data[pos]))
								pos++;		
							c = data[pos];
							start = pos;	
						}
						
						if (c == delimiter || isEOL(c) || pos >= len - 1) {
							int vLen = pos - start + (pos >= len - 1);		
							int qPos = -1;
							for (int i = 0; i < vLen; i++) {
								TCHAR c = data[start + i];
								if (c != TEXT(' ') && c != TEXT('\t')) {
									qPos = c == TEXT('"') ? i : -1;
									break;
								}
							}
							
							if (qPos != -1) {
								int qCount = 0;								
								for (int i = qPos; i <= vLen; i++)
									qCount += data[start + i] == TEXT('"');
									
								if (pos < len && qCount % 2)	
									continue;
									
								while(vLen > 0 && data[start + vLen] != TEXT('"'))
									vLen--;
									
								vLen -= qPos + 1;
							} 
							
							if (stepNo == 1) {
								if (isTrim) {
									int l = 0, r = 0;
									if (qPos == -1) {
										for (int i = 0; i < vLen && (data[start + i] == TEXT(' ') || data[start + i] == TEXT('\t')); i++) 
											l++;
									}
									for (int i = vLen - 1; i > 0 && (data[start + i] == TEXT(' ') || data[start + i] == TEXT('\t')); i--) 
										r++;
										
									start += l; 
									vLen -= l + r;
								}
								
								TCHAR* value = calloc(vLen + 1, sizeof(TCHAR));
								if (vLen > 0) {
									_tcsncpy(value, data + start + qPos + 1, vLen);
									
									// replace "" by " in quoted value
									if (qPos != -1) {
										int qCount = 0, j = 0;
										for (int i = 0; i < vLen; i++) {
											BOOL isQ = value[i] == TEXT('"'); 
											qCount += isQ;
											value[j] = value[i];
											j += isQ && qCount % 2 || !isQ; 
										}
										value[j] = 0;
									}
								}
								
								cache[rowNo][colNo] = value;
							}
			 		
							start = pos + 1;
							colNo++;
						}
						
						if (isEOL(data[pos])) { 
							while (isEOL(data[pos + 1]))
								pos++;
							break;
						}
					}
					
					while (pos < len && !isEOL(data[pos - 1]))
						pos++;
					
					while (stepNo == 1 && colNo < colCount) {
						cache[rowNo][colNo] = calloc(1, sizeof(TCHAR));
						colNo++;
					}
					
					if (stepNo == 0 && colCount < colNo)
						colCount = colNo;
				}
				
				if (stepNo == 1) {
					cache = realloc(cache, (rowNo + 1) * sizeof(TCHAR**));
					if (codepage == CP_UTF16LE || codepage == CP_UTF16BE)
						free(rawdata);
					else						
						free(data);
				}
			}	
						
			if (colCount > MAX_COLUMN_COUNT) {
				HWND hListerWnd = GetParent(hWnd);
				TCHAR msg[255];
				_sntprintf(msg, 255, TEXT("Column count is overflow.\nFound: %i, max: %i"), colCount, MAX_COLUMN_COUNT);
				MessageBox(hWnd, msg, NULL, 0);
				SendMessage(hListerWnd, WM_CLOSE, 0, 0);
				return 0;
			}
						
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
			
			_sntprintf(buf, 32, TEXT(" %ls"), skipComments == 0 ? TEXT("#0   ") : skipComments == 1 ? TEXT("#1   ") : TEXT("#2   "));
			SendMessage(hStatusWnd, SB_SETTEXT, SB_COMMENTS, (LPARAM)buf);
									
			SetProp(hWnd, TEXT("CACHE"), cache);
			*(int*)GetProp(hWnd, TEXT("COLCOUNT")) = colCount;			
			*(int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")) = rowNo + 1;
			*(int*)GetProp(hWnd, TEXT("CODEPAGE")) = codepage;			
			*(TCHAR*)GetProp(hWnd, TEXT("DELIMITER")) = delimiter;
			*(int*)GetProp(hWnd, TEXT("ORDERBY")) = 0;
									
			SendMessage(hWnd, WMU_UPDATE_GRID, 0, 0);
			
			return TRUE;
		}
		break;

		case WMU_UPDATE_GRID: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int rowCount = *(int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
			int isHeaderRow = *(int*)GetProp(hWnd, TEXT("HEADERROW"));
			int filterAlign = *(int*)GetProp(hWnd, TEXT("FILTERALIGN"));			

			int colCount = Header_GetItemCount(hHeader);
			for (int colNo = 0; colNo < colCount; colNo++)
				DestroyWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo));

			for (int colNo = 0; colNo < colCount; colNo++)
				ListView_DeleteColumn(hGridWnd, colCount - colNo - 1);
				
			colCount = *(int*)GetProp(hWnd, TEXT("COLCOUNT"));
			for (int colNo = 0; colNo < colCount; colNo++) {
				int cNum = 0;
				int cCount = rowCount < 6 ? rowCount : 6;
				for (int rowNo = 1; rowNo < cCount; rowNo++)
					cNum += cache[rowNo][colNo] && _tcslen(cache[rowNo][colNo]) && isNumber(cache[rowNo][colNo]);
				
				int fmt = cNum > cCount / 2 ? LVCFMT_RIGHT : LVCFMT_LEFT;				
				TCHAR colName[64];
				_sntprintf(colName, 64, TEXT("Column #%i"), colNo + 1);
				ListView_AddColumn(hGridWnd, isHeaderRow && cache[0][colNo] && _tcslen(cache[0][colNo]) > 0 ? cache[0][colNo] : colName, fmt);
			}
	
			int align = filterAlign == -1 ? ES_LEFT : filterAlign == 1 ? ES_RIGHT : ES_CENTER;
			for (int colNo = 0; colNo < colCount; colNo++) {
				// Use WS_BORDER to vertical text aligment
				HWND hEdit = CreateWindowEx(WS_EX_TOPMOST, WC_EDIT, NULL, align | ES_AUTOHSCROLL | WS_CHILD | WS_TABSTOP | WS_BORDER,
					0, 0, 0, 0, hHeader, (HMENU)(INT_PTR)(IDC_HEADER_EDIT + colNo), GetModuleHandle(0), NULL);
				SendMessage(hEdit, WM_SETFONT, (LPARAM)GetProp(hWnd, TEXT("FONT")), TRUE);
				SetProp(hEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)cbNewFilterEdit));
			}			

			SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);
			SendMessage(hWnd, WMU_SET_HEADER_FILTERS, 0, 0);			
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
			BOOL isCaseSensitive = getStoredValue(TEXT("filter-case-sensitive"), 0);
			
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
						bResultset[rowNo] = len == 1 ? hasString(value, filter, isCaseSensitive) :
							filter[0] == TEXT('=') && isCaseSensitive ? _tcscmp(value, filter + 1) == 0 :
							filter[0] == TEXT('=') && !isCaseSensitive ? _tcsicmp(value, filter + 1) == 0 :							
							filter[0] == TEXT('!') ? !hasString(value, filter + 1, isCaseSensitive) :
							filter[0] == TEXT('<') ? _tcscmp(value, filter + 1) < 0 :
							filter[0] == TEXT('>') ? _tcscmp(value, filter + 1) > 0 :
							hasString(value, filter, isCaseSensitive);
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
					BOOL isBackward = orderBy < 0;
					
					BOOL isNum = TRUE;
					for (int i = isHeaderRow; i < *pTotalRowCount && i <= 5; i++) 
						isNum = isNum && isNumber(cache[i][colNo]);
												
					if (isNum) {
						double* nums = calloc(*pTotalRowCount, sizeof(double));
						for (int i = 0; i < rowCount; i++)
							nums[resultset[i]] = _tcstod(cache[resultset[i]][colNo], NULL);

						mergeSort(resultset, (void*)nums, 0, rowCount - 1, isBackward, isNum);
						free(nums);
					} else {
						TCHAR** strings = calloc(*pTotalRowCount, sizeof(TCHAR*));
						for (int i = 0; i < rowCount; i++)
							strings[resultset[i]] = cache[resultset[i]][colNo];
						mergeSort(resultset, (void*)strings, 0, rowCount - 1, isBackward, isNum);
						free(strings);
					}
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
				SetWindowPos(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), 0, rc.left, h2, rc.right - rc.left, h2 + 1, SWP_NOZORDER);
			}
		}
		break;
		
		case WMU_SET_HEADER_FILTERS: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int isFilterRow = *(int*)GetProp(hWnd, TEXT("FILTERROW"));
			int colCount = Header_GetItemCount(hHeader);
			
			SendMessage(hWnd, WM_SETREDRAW, FALSE, 0);
			LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
			styles = isFilterRow ? styles | HDS_FILTERBAR : styles & (~HDS_FILTERBAR);
			SetWindowLongPtr(hHeader, GWL_STYLE, styles);

			for (int colNo = 0; colNo < colCount; colNo++) 		
				ShowWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), isFilterRow ? SW_SHOW : SW_HIDE);

			if (isFilterRow)				
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);											

			// Bug fix: force Windows to redraw header
			if (IsWindowVisible(hGridWnd)) { // Win10x64, TCx32 
				int w = ListView_GetColumnWidth(hGridWnd, 0);
				ListView_SetColumnWidth(hGridWnd, 0, w + 1);
				ListView_SetColumnWidth(hGridWnd, 0, w);			
			}
						
			SendMessage(hWnd, WM_SETREDRAW, TRUE, 0);
			InvalidateRect(hWnd, NULL, TRUE);
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
		
		// wParam = rowNo, lParam = colNo
		case WMU_SET_CURRENT_CELL: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)0);
						
			int *pRowNo = (int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
			int *pColNo = (int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
			if (*pRowNo == wParam && *pColNo == lParam)
				return 0;
			
			RECT rc, rc2;
			ListView_GetSubItemRect(hGridWnd, *pRowNo, *pColNo, LVIR_BOUNDS, &rc);
			if (*pColNo == 0)
				rc.right = ListView_GetColumnWidth(hGridWnd, *pColNo);			
			InvalidateRect(hGridWnd, &rc, TRUE);
			
			*pRowNo = wParam;
			*pColNo = lParam;
			ListView_GetSubItemRect(hGridWnd, *pRowNo, *pColNo, LVIR_BOUNDS, &rc);
			if (*pColNo == 0)
				rc.right = ListView_GetColumnWidth(hGridWnd, *pColNo);
			InvalidateRect(hGridWnd, &rc, FALSE);
			
			GetClientRect(hGridWnd, &rc2);
			int w = rc.right - rc.left;
			int dx = rc2.right < rc.right ? rc.left - rc2.right + w : rc.left < 0 ? rc.left : 0;

			ListView_Scroll(hGridWnd, dx, 0);
			
			TCHAR buf[32] = {0};
			if (*pColNo != - 1 && *pRowNo != -1)
				_sntprintf(buf, 32, TEXT(" %i:%i"), *pRowNo + 1, *pColNo + 1);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_CELL, (LPARAM)buf);
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
		
		case WMU_SET_THEME: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			BOOL isDark = *(int*)GetProp(hWnd, TEXT("DARKTHEME"));
			
			int textColor = !isDark ? getStoredValue(TEXT("text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("text-color-dark"), RGB(220, 220, 220));
			int backColor = !isDark ? getStoredValue(TEXT("back-color"), RGB(255, 255, 255)) : getStoredValue(TEXT("back-color-dark"), RGB(32, 32, 32));
			int backColor2 = !isDark ? getStoredValue(TEXT("back-color2"), RGB(240, 240, 240)) : getStoredValue(TEXT("back-color2-dark"), RGB(52, 52, 52));
			int filterTextColor = !isDark ? getStoredValue(TEXT("filter-text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("filter-text-color-dark"), RGB(255, 255, 255));
			int filterBackColor = !isDark ? getStoredValue(TEXT("filter-back-color"), RGB(240, 240, 240)) : getStoredValue(TEXT("filter-back-color-dark"), RGB(60, 60, 60));
			int currCellColor = !isDark ? getStoredValue(TEXT("current-cell-color"), RGB(20, 250, 250)) : getStoredValue(TEXT("current-cell-color-dark"), RGB(20, 250, 250));
			
			*(int*)GetProp(hWnd, TEXT("TEXTCOLOR")) = textColor;
			*(int*)GetProp(hWnd, TEXT("BACKCOLOR")) = backColor;
			*(int*)GetProp(hWnd, TEXT("BACKCOLOR2")) = backColor2;			
			*(int*)GetProp(hWnd, TEXT("FILTERTEXTCOLOR")) = filterTextColor;
			*(int*)GetProp(hWnd, TEXT("FILTERBACKCOLOR")) = filterBackColor;
			*(int*)GetProp(hWnd, TEXT("CURRENTCELLCOLOR")) = currCellColor;			
			
			DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));			
			SetProp(hWnd, TEXT("FILTERBACKBRUSH"), CreateSolidBrush(filterBackColor));

			ListView_SetTextColor(hGridWnd, textColor);
			ListView_SetBkColor(hGridWnd, backColor);
			ListView_SetTextBkColor(hGridWnd, backColor);			
			InvalidateRect(hWnd, NULL, TRUE);	
		}
		break;
		
		case WMU_HOT_KEYS: {
			BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
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

				no += isCtrl ? -1 : 1;
				SetFocus(wnds[no] && no >= 0 ? wnds[no] : (isCtrl ? wnds[cnt - 1] : wnds[0]));
			}
			
			if (wParam == VK_F1) {
				ShellExecute(0, 0, TEXT("https://github.com/little-brother/csvtab-wlx/wiki"), 0, 0 , SW_SHOW);
				return TRUE;
			}
			
			if (wParam == 0x20 && isCtrl) { // Ctrl + Space
				SendMessage(hWnd, WMU_SHOW_COLUMNS, 0, 0);
				return TRUE;
			}
			
			if (wParam == VK_ESCAPE || wParam == VK_F11 ||
				wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7 || (isCtrl && wParam == 0x46) || // Ctrl + F
				((wParam >= 0x31 && wParam <= 0x38) && !getStoredValue(TEXT("disable-num-keys"), 0) || // 1 - 8
				(wParam == 0x4E || wParam == 0x50) && !getStoredValue(TEXT("disable-np-keys"), 0)) && // N, P
				GetDlgCtrlID(GetFocus()) / 100 * 100 != IDC_HEADER_EDIT) { 
				SetFocus(GetParent(hWnd));		
				keybd_event(wParam, wParam, KEYEVENTF_EXTENDEDKEY, 0);

				return TRUE;
			}			
			
			return FALSE;					
		}
		break;
		
		case WMU_HOT_CHARS: {
			BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
			return !_istprint(wParam) && (
				wParam == VK_ESCAPE || wParam == VK_F11 || wParam == VK_F1 ||
				wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7) ||
				wParam == VK_TAB || wParam == VK_RETURN ||
				isCtrl && (wParam == 0x46 || wParam == 0x20);
		}
		break;			
	}
	
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_KEYDOWN && SendMessage(getMainWindow(hWnd), WMU_HOT_KEYS, wParam, lParam))
		return 0;

	// Prevent beep
	if (msg == WM_CHAR && SendMessage(getMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
		return 0;	

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_CTLCOLOREDIT) {
		HWND hMainWnd = getMainWindow(hWnd);
		SetBkColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERBACKCOLOR")));
		SetTextColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERTEXTCOLOR")));
		return (INT_PTR)(HBRUSH)GetProp(hMainWnd, TEXT("FILTERBACKBRUSH"));	
	}
	
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	WNDPROC cbDefault = (WNDPROC)GetProp(hWnd, TEXT("WNDPROC"));

	switch(msg){
		case WM_PAINT: {
			cbDefault(hWnd, msg, wParam, lParam);

			RECT rc;
			GetClientRect(hWnd, &rc);
			HWND hMainWnd = getMainWindow(hWnd);
			BOOL isDark = *(int*)GetProp(hMainWnd, TEXT("DARKTHEME")); 

			HDC hDC = GetWindowDC(hWnd);
			HPEN hPen = CreatePen(PS_SOLID, 1, *(int*)GetProp(hMainWnd, TEXT("FILTERBACKCOLOR")));
			HPEN oldPen = SelectObject(hDC, hPen);
			MoveToEx(hDC, 1, 0, 0);
			LineTo(hDC, rc.right - 1, 0);
			LineTo(hDC, rc.right - 1, rc.bottom - 1);
			
			if (isDark) {
				DeleteObject(hPen);
				hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNFACE));
				SelectObject(hDC, hPen);
				
				MoveToEx(hDC, 0, 0, 0);
				LineTo(hDC, 0, rc.bottom);
				MoveToEx(hDC, 0, rc.bottom - 1, 0);
				LineTo(hDC, rc.right, rc.bottom - 1);
				MoveToEx(hDC, 0, rc.bottom - 2, 0);
				LineTo(hDC, rc.right, rc.bottom - 2);
			}
			
			SelectObject(hDC, oldPen);
			DeleteObject(hPen);
			ReleaseDC(hWnd, hDC);

			return 0;
		}
		break;
		
		case WM_SETFOCUS: {
			SetProp(getMainWindow(hWnd), TEXT("LASTFOCUS"), hWnd);
		}
		break;

		case WM_KEYDOWN: {
			HWND hMainWnd = getMainWindow(hWnd);
			if (wParam == VK_RETURN) {
				SendMessage(hMainWnd, WMU_UPDATE_RESULTSET, 0, 0);
				return 0;			
			}
			
			if (SendMessage(hMainWnd, WMU_HOT_KEYS, wParam, lParam))
				return 0;
		}
		break;
	
		// Prevent beep
		case WM_CHAR: {
			if (SendMessage(getMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
				return 0;	
		}
		break;
		
		case WM_DESTROY: {
			RemoveProp(hWnd, TEXT("WNDPROC"));
		}
		break;
	}

	return CallWindowProc(cbDefault, hWnd, msg, wParam, lParam);
}

HWND getMainWindow(HWND hWnd) {
	HWND hMainWnd = hWnd;
	while (hMainWnd && GetDlgCtrlID(hMainWnd) != IDC_MAIN)
		hMainWnd = GetParent(hMainWnd);
	return hMainWnd;	
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

TCHAR detectDelimiter(const TCHAR *data, BOOL skipComments) {
	int rowCount = 5;
	int delimCount = _tcslen(DELIMITERS);
	int colCount[rowCount][delimCount];
	memset(colCount, 0, sizeof(int) * (size_t)(rowCount * delimCount));
	
	int pos = 0;
	int len = _tcslen(data);
	BOOL inQuote = FALSE;		
	int rowNo = 0;	
	for (; rowNo < rowCount && pos < len;) {
		int total = 0;
		for (; pos < len; pos++) {
			TCHAR c = data[pos];
			
			while (!inQuote && skipComments == 2 && c == TEXT('#')) {
				while (data[pos] && !isEOL(data[pos]))
					pos++;
				while (pos < len && isEOL(data[pos]))
					pos++;		
				c = data[pos];
			}		
			
			if (!inQuote && c == TEXT('\n')) {
				break;
			}

			if (c == TEXT('"'))
				inQuote = !inQuote;
				
			for (int delimNo = 0; delimNo < delimCount && !inQuote; delimNo++) {
				colCount[rowNo][delimNo] += data[pos] == DELIMITERS[delimNo];
				total += colCount[rowNo][delimNo];
			}				
		}
		
		rowNo += total > 0;
		pos++;
	}

	rowCount = rowNo;
	int maxNo = 0;
	int total[delimCount];
	memset(total, 0, sizeof(int) * (size_t)(delimCount));
	for (int delimNo = 0; delimNo < delimCount; delimNo++) {
		for (int i = 0; i < rowCount; i++) {
			for (int j = 0; j < rowCount; j++) {
				total[delimNo] += colCount[i][delimNo] == colCount[j][delimNo] && colCount[i][delimNo] > 0 ? 10 + colCount[i][delimNo] :
					colCount[i][delimNo] != 0 ? 5 :
					0;	 
			}				
		}
	}

	int maxCount = 0;
	for (int delimNo = 0; delimNo < delimCount; delimNo++) {
		if (maxCount < total[delimNo]) {
			maxNo = delimNo;
			maxCount = total[delimNo];
		}
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

BOOL isEOL(TCHAR c) {
	return c == TEXT('\r') || c == TEXT('\n');
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

int findString(TCHAR* text, TCHAR* word, BOOL isMatchCase, BOOL isWholeWords) {
	if (!text || !word)
		return -1;
		
	int res = -1;
	int tlen = _tcslen(text);
	int wlen = _tcslen(word);	
	if (!tlen || !wlen)
		return res;
	
	if (!isMatchCase) {
		TCHAR* ltext = calloc(tlen + 1, sizeof(TCHAR));
		_tcsncpy(ltext, text, tlen);
		text = _tcslwr(ltext);

		TCHAR* lword = calloc(wlen + 1, sizeof(TCHAR));
		_tcsncpy(lword, word, wlen);
		word = _tcslwr(lword);
	}

	if (isWholeWords) {
		for (int pos = 0; (res  == -1) && (pos <= tlen - wlen); pos++) 
			res = (pos == 0 || pos > 0 && !_istalnum(text[pos - 1])) && 
				!_istalnum(text[pos + wlen]) && 
				_tcsncmp(text + pos, word, wlen) == 0 ? pos : -1;
	} else {
		TCHAR* s = _tcsstr(text, word);
		res = s != NULL ? s - text : -1;
	}
	
	if (!isMatchCase) {
		free(text);
		free(word);
	}

	return res; 
}	

BOOL hasString (const TCHAR* str, const TCHAR* sub, BOOL isCaseSensitive) {
	BOOL res = FALSE;

	TCHAR* lstr = _tcsdup(str);
	_tcslwr(lstr);
	TCHAR* lsub = _tcsdup(sub);
	_tcslwr(lsub);
	res = isCaseSensitive ? _tcsstr(str, sub) != 0 : _tcsstr(lstr, lsub) != 0;
	free(lstr);
	free(lsub);

	return res;
};

TCHAR* extractUrl(TCHAR* data) {
	int len = data ? _tcslen(data) : 0;
	int start = len;
	int end = len;
	
	TCHAR* url = calloc(len + 10, sizeof(TCHAR));
	
	TCHAR* slashes = _tcsstr(data, TEXT("://"));
	if (slashes) {
		start = len - _tcslen(slashes);
		end = start + 3;
		for (; start > 0 && _istalpha(data[start - 1]); start--);
		for (; end < len && data[end] != TEXT(' ') && data[end] != TEXT('"') && data[end] != TEXT('\''); end++);
		_tcsncpy(url, data + start, end - start);
		
	} else if (_tcschr(data, TEXT('.'))) {
		_sntprintf(url, len + 10, TEXT("https://%ls"), data);
	}
	
	return url;
}	

void mergeSortJoiner(int indexes[], void* data, int l, int m, int r, BOOL isBackward, BOOL isNums) {
	int n1 = m - l + 1;
	int n2 = r - m;
	
	int* L = calloc(n1, sizeof(int));
	int* R = calloc(n2, sizeof(int)); 
	
	for (int i = 0; i < n1; i++)
		L[i] = indexes[l + i];
	for (int j = 0; j < n2; j++)
		R[j] = indexes[m + 1 + j];
	
	int i = 0, j = 0, k = l;
	while (i < n1 && j < n2) {
		int cmp = isNums ? ((double*)data)[L[i]] <= ((double*)data)[R[j]] : _tcscmp(((TCHAR**)data)[L[i]], ((TCHAR**)data)[R[j]]) <= 0;
		if (isBackward)
			cmp = !cmp;
		
		if (cmp) {
			indexes[k] = L[i];
			i++;
		} else {
			indexes[k] = R[j];
			j++;
		}
		k++;
	}
	
	while (i < n1) {
		indexes[k] = L[i];
		i++;
		k++;
	}
	
	while (j < n2) {
		indexes[k] = R[j];
		j++;
		k++;
	}
	
	free(L);
	free(R);
}

void mergeSort(int indexes[], void* data, int l, int r, BOOL isBackward, BOOL isNums) {
	if (l < r) {
		int m = l + (r - l) / 2;
		mergeSort(indexes, data, l, m, isBackward, isNums);
		mergeSort(indexes, data, m + 1, r, isBackward, isNums);
		mergeSortJoiner(indexes, data, l, m, r, isBackward, isNums);
	}
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

void Menu_SetItemState(HMENU hMenu, UINT wID, UINT fState) {
	MENUITEMINFO mii = {0};
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask = MIIM_STATE;
	mii.fState = fState;
	SetMenuItemInfo(hMenu, wID, FALSE, &mii);
}