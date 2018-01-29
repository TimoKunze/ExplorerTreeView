// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"
#include ".\maindlg.h"


BOOL CMainDlg::IsComctl32Version600OrNewer(void)
{
	BOOL ret = FALSE;

	HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
	if(hMod) {
		DLLGETVERSIONPROC pDllGetVersion = reinterpret_cast<DLLGETVERSIONPROC>(GetProcAddress(hMod, "DllGetVersion"));
		if(pDllGetVersion) {
			DLLVERSIONINFO versionInfo = {0};
			versionInfo.cbSize = sizeof(versionInfo);
			HRESULT hr = (*pDllGetVersion)(&versionInfo);
			if(SUCCEEDED(hr)) {
				ret = (versionInfo.dwMajorVersion >= 6);
			}
		}
		FreeLibrary(hMod);
	}

	return ret;
}


BOOL CMainDlg::PreTranslateMessage(MSG* pMsg)
{
	if((pMsg->message == WM_KEYDOWN) && ((pMsg->wParam == VK_ESCAPE) || (pMsg->wParam == VK_RETURN))) {
		// otherwise label-edit wouldn't work
		return FALSE;
	} else {
		return CWindow::IsDialogMessage(pMsg);
	}
}

LRESULT CMainDlg::OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	CloseDialog(0);
	return 0;
}

LRESULT CMainDlg::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	CComPtr<ITreeViewItems> pItems = controls.extvwU->GetTreeItems();
	CComQIPtr<IEnumVARIANT> pEnum = pItems;
	pEnum->Reset();     // NOTE: It's important to call Reset()!
	ULONG ul = 0;
	_variant_t v;
	v.Clear();
	while(pEnum->Next(1, &v, &ul) == S_OK) {
		CComQIPtr<ITreeViewItem> pItem = v.pdispVal;
		HGDIOBJ h = static_cast<HGDIOBJ>(LongToHandle(pItem->GetItemData()));
		if(GetObjectType(h) == OBJ_FONT) {
			DeleteObject(h);
		}
	}

	if(controls.extvwU) {
		DispEventUnadvise(controls.extvwU);
		controls.extvwU.Release();
	}

	if(hFontBold) {
		DeleteObject(hFontBold);
	}

	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);

	bHandled = FALSE;
	return S_OK;
}

LRESULT CMainDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	// Init resizing
	DlgResize_Init(false, false);

	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	extvwUWnd.SubclassWindow(GetDlgItem(IDC_EXTVWU));
	extvwUWnd.QueryControl(__uuidof(IExplorerTreeView), reinterpret_cast<LPVOID*>(&controls.extvwU));
	if(controls.extvwU) {
		DispEventAdvise(controls.extvwU);

		// create a bold font
		hFontBold = reinterpret_cast<HFONT>(SendMessage(static_cast<HWND>(LongToHandle(controls.extvwU->GethWnd())), WM_GETFONT, 0, 0));
		LOGFONT lf = {0};
		GetObject(hFontBold, sizeof(LOGFONT), &lf);
		lf.lfWeight |= FW_BOLD;
		hFontBold = CreateFontIndirect(&lf);

		InsertItems();
	}

	if(CTheme::IsThemingSupported() && IsComctl32Version600OrNewer()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed && pfnIsAppThemed()) {
				SetWindowTheme(static_cast<HWND>(LongToHandle(controls.extvwU->GethWnd())), L"explorer", NULL);
			}
			FreeLibrary(hThemeDLL);
		}
	}

	return TRUE;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.extvwU) {
		controls.extvwU->About();
	}
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

int CALLBACK CMainDlg::EnumFontFamExProc(const LPENUMLOGFONTEX lpElfe, const NEWTEXTMETRICEX* /*lpntme*/, DWORD /*FontType*/, LPARAM lParam)
{
	static CAtlString lastFontName = TEXT("");

	if(lastFontName != CAtlString(lpElfe->elfFullName)) {
		LPENUMFONTPARAM pParam = reinterpret_cast<LPENUMFONTPARAM>(lParam);
		if(pParam->firstCall) {
			//pParam->pParentItem->PutText(_variant_t(lpElfe->elfScript));
			pParam->firstCall = FALSE;
		}

		lpElfe->elfLogFont.lfWidth = pParam->lfTreeView.lfWidth;
		lpElfe->elfLogFont.lfHeight = pParam->lfTreeView.lfHeight;

		HFONT hFont = CreateFontIndirect(&lpElfe->elfLogFont);
		_variant_t emptyVariant;
		pParam->pSubItemsToAddTo->Add(W2OLE(lpElfe->elfFullName), NULL, emptyVariant, heNo, 0, 0, 0, HandleToLong(hFont), VARIANT_FALSE, 1);
	}
	lastFontName = lpElfe->elfFullName;

	return 1;
}

void CMainDlg::InsertItems(void)
{
	// setup tooltip
	CWindow wnd = static_cast<HWND>(LongToHandle(controls.extvwU->GethWndToolTip()));
	wnd.ModifyStyle(WS_BORDER, TTS_BALLOON);
	CToolTipCtrl tooltip = wnd;
	tooltip.SetTitle(1, TEXT("Font name:"));

	// insert items
	CComPtr<ITreeViewItems> pItems = controls.extvwU->GetTreeItems();
	if(pItems) {
		_variant_t emptyVariant;
		emptyVariant.Clear();

		CComPtr<ITreeViewItem> pColorsRootItem = pItems->Add(OLESTR("Some Colors and their (HTML) names"), NULL, emptyVariant, heYes, 0, 0, 0, 3, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0xFFFFFF   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -1, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x0000FF   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -2, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x00FF00   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -3, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0xFF0000   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -4, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0xFF00FF   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -5, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0xFFFF00   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -6, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x00FFFF   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -7, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x000000   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -8, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x93DB70   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -9, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x17335C   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -10, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x9F5F9F   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -11, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x42A6B5   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -12, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x19D9D9   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -13, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x2A2AA6   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -14, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x53788C   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -15, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x3D7DA6   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -16, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x9F9F5F   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -17, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x1987D9   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -18, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x3373B8   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -19, VARIANT_FALSE, 1);
		pItems->Add(OLESTR("0x007FFF   "), pColorsRootItem, emptyVariant, heNo, 0, 0, 0, -20, VARIANT_FALSE, 1);

		HFONT hFont = reinterpret_cast<HFONT>(SendMessage(static_cast<HWND>(LongToHandle(controls.extvwU->GethWnd())), WM_GETFONT, 0, 0));
		ENUMFONTPARAM param = {0};
		ZeroMemory(&param.lfTreeView, sizeof(LOGFONT));
		GetObject(hFont, sizeof(LOGFONT), &param.lfTreeView);
		param.lfTreeView.lfFaceName[0] = '\0';
		HDC hDC = ::GetDC(NULL);

		CComPtr<ITreeViewItem> pFontsRootItem = pItems->Add(OLESTR("Fonts installed in the system"), NULL, emptyVariant, heYes, 0, 0, 0, 3, VARIANT_FALSE, 1);
		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("ANSI_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -100, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = ANSI_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("BALTIC_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -101, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = BALTIC_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("CHINESEBIG5_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -102, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = CHINESEBIG5_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("DEFAULT_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -103, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = DEFAULT_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("EASTEUROPE_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -104, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = EASTEUROPE_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("GB2312_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -105, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = GB2312_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("GREEK_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -106, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = GREEK_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("HANGUL_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -107, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = HANGUL_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("MAC_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -108, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = MAC_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("OEM_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -109, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = OEM_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("RUSSIAN_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -110, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = RUSSIAN_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("SHIFTJIS_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -111, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = SHIFTJIS_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("SYMBOL_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -112, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = SYMBOL_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.firstCall = TRUE;
		param.pParentItem = pItems->Add(OLESTR("TURKISH_CHARSET"), pFontsRootItem, emptyVariant, heYes, 0, 0, 0, -113, VARIANT_FALSE, 1);
		param.pSubItemsToAddTo = param.pParentItem->GetSubItems();
		param.lfTreeView.lfCharSet = TURKISH_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfTreeView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		pFontsRootItem->SortSubItems(sobText, sobNone, sobNone, sobNone, sobNone, VARIANT_TRUE, VARIANT_FALSE);

		::ReleaseDC(NULL, hDC);
	}
}

void __stdcall CMainDlg::CustomDrawExtvwu(LPDISPATCH treeItem, long* /*itemLevel*/, OLE_COLOR* textColor, OLE_COLOR* /*textBackColor*/, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants itemState, long hDC, RECTANGLE* /*drawingRectangle*/, CustomDrawReturnValuesConstants* furtherProcessing)
{
	switch(drawStage) {
		case cdsPrePaint:
			*furtherProcessing = static_cast<CustomDrawReturnValuesConstants>(cdrvNotifyPostPaint | cdrvNotifyItemDraw);
			break;

		case cdsItemPrePaint: {
			CComQIPtr<ITreeViewItem> pItem = treeItem;
			if(pItem) {
				if((pItem->GetItemData() >= -20) && (pItem->GetItemData() <= -1)) {
					if(itemState & cdisFocus) {
					// use the following line instead of the previous line, if you've set AllowDragDrop to True
					//if(*textBackColor = GetSysColor(COLOR_HIGHLIGHT)) {
						SelectObject(static_cast<HDC>(LongToHandle(hDC)), hFontBold);
					} else {
						*textColor = 0;
						LPWSTR pText = OLE2W(pItem->GetText());
						if(lstrlenW(pText)) {
							LPWSTR pDummy;
							*textColor = wcstol(pText, &pDummy, 16);
						}
					}
				} else {
					HGDIOBJ h = static_cast<HGDIOBJ>(LongToHandle(pItem->GetItemData()));
					if(GetObjectType(h) == OBJ_FONT) {
						SelectObject(static_cast<HDC>(LongToHandle(hDC)), h);
					}
				}
				*furtherProcessing = static_cast<CustomDrawReturnValuesConstants>(cdrvNewFont | cdrvNotifyPostPaint);
			}
			break;
		}

		case cdsItemPostPaint: {
			CComQIPtr<ITreeViewItem> pItem = treeItem;
			if(pItem) {
				if((pItem->GetItemData() >= -20) && (pItem->GetItemData() <= -1)) {
					RECT rc = {0};
					try {
						pItem->GetRectangle(irtItem, &rc.left, &rc.top, &rc.right, &rc.bottom);
					} catch(_com_error) {
						// GetRectangle will fail if the item is not yet visible to the user
						*furtherProcessing = cdrvDoDefault;
						return;
					}
					rc.left = rc.right + 5;
					CRect clientRectangle;
					extvwUWnd.GetClientRect(&clientRectangle);
					rc.right = clientRectangle.Width();

					if(itemState & cdisFocus) {
						COLORREF nOldClr = SetTextColor(static_cast<HDC>(LongToHandle(hDC)), RGB(0, 0, 255));
						switch(pItem->GetItemData()) {
							case -1:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("White"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -2:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Red"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -3:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Green"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -4:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Blue"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -5:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Magenta"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -6:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Cyan"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -7:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Yellow"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -8:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Black"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -9:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Aquamarine"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -10:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Baker's Chocolate"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -11:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Blue Violet"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -12:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Brass"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -13:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Bright Gold"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -14:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Brown"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -15:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Bronze"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -16:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Bronze II"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -17:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Cadet Blue"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -18:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Cool Copper"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -19:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Copper"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
							case -20:
								DrawText(static_cast<HDC>(LongToHandle(hDC)), TEXT("Coral"), -1, &rc, DT_LEFT | DT_VCENTER);
								break;
						}
						SetTextColor(static_cast<HDC>(LongToHandle(hDC)), nOldClr);
					} else {
						FillRect(static_cast<HDC>(LongToHandle(hDC)), &rc, GetSysColorBrush(COLOR_WINDOW));
					}
				}
			}
			break;
		}
	}
}

void __stdcall CMainDlg::FreeItemDataExtvwu(LPDISPATCH treeItem)
{
	CComQIPtr<ITreeViewItem> pItem = treeItem;
	if(pItem) {
		HGDIOBJ h = static_cast<HGDIOBJ>(LongToHandle(pItem->GetItemData()));
		if(GetObjectType(h) == OBJ_FONT) {
			DeleteObject(h);
		}
	}
}

void __stdcall CMainDlg::ItemGetInfoTipTextExtvwu(LPDISPATCH treeItem, long /*maxInfoTipLength*/, BSTR* infoTipText, VARIANT_BOOL* /*abortToolTip*/)
{
	CComQIPtr<ITreeViewItem> pItem = treeItem;
	if(pItem) {
		HGDIOBJ h = static_cast<HGDIOBJ>(LongToHandle(pItem->GetItemData()));
		if(GetObjectType(h) == OBJ_FONT) {
			*infoTipText = _bstr_t(pItem->GetText()).Detach();
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowExtvwu(long /*hWnd*/)
{
	InsertItems();
}
