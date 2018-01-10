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
	if(controls.extvwU) {
		DispEventUnadvise(controls.extvwU);
		controls.extvwU.Release();
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

	controls.imageList.CreateFromImage(IDB_ICONS, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	extvwUWnd.SubclassWindow(GetDlgItem(IDC_EXTVWU));
	extvwUWnd.QueryControl(__uuidof(IExplorerTreeView), reinterpret_cast<LPVOID*>(&controls.extvwU));
	if(controls.extvwU) {
		DispEventAdvise(controls.extvwU);
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

void CMainDlg::InsertItems(void)
{
	controls.extvwU->PuthImageList(ilItems, HandleToLong(controls.imageList.m_hImageList));
	int cImages = controls.imageList.GetImageCount();

	CComPtr<ITreeViewItems> pItems = controls.extvwU->GetTreeItems();
	if(pItems) {
		_variant_t emptyVariant;
		emptyVariant.Clear();
		int iIcon = 0;

		CComPtr<ITreeViewItem> pRootItem = pItems->Add(OLESTR("Level 1, Item 1"), NULL, emptyVariant, heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;

		CComPtr<ITreeViewItems> pItemsLevel2 = pRootItem->GetSubItems();
		pItemsLevel2->Add(OLESTR("Level 2, Item 1"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pRootItem->Expand(VARIANT_FALSE);

		CComPtr<ITreeViewItem> pItem = pItemsLevel2->Add(OLESTR("Level 2, Item 2"), NULL, emptyVariant, heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		CComPtr<ITreeViewItems> pItemsLevel3 = pItem->GetSubItems();
		pItemsLevel3->Add(OLESTR("Level 3, Item 1"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 2"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 3"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItem->Expand(VARIANT_FALSE);

		pItemsLevel2->Add(OLESTR("Level 2, Item 3"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel2->Add(OLESTR("Level 2, Item 4"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;

		pItem = pItemsLevel2->Add(OLESTR("Level 2, Item 5"), NULL, emptyVariant, heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3 = pItem->GetSubItems();
		pItemsLevel3->Add(OLESTR("Level 3, Item 4"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 5"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItem->Expand(VARIANT_FALSE);


		pRootItem = pItems->Add(OLESTR("Level 1, Item 1"), NULL, emptyVariant, heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;

		pItemsLevel2 = pRootItem->GetSubItems();
		pItemsLevel2->Add(OLESTR("Level 2, Item 1"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pRootItem->Expand(VARIANT_FALSE);

		pItem = pItemsLevel2->Add(OLESTR("Level 2, Item 2"), NULL, emptyVariant, heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3 = pItem->GetSubItems();
		pItemsLevel3->Add(OLESTR("Level 3, Item 1"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 2"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 3"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItem->Expand(VARIANT_FALSE);

		pItemsLevel2->Add(OLESTR("Level 2, Item 3"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel2->Add(OLESTR("Level 2, Item 4"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;

		pItem = pItemsLevel2->Add(OLESTR("Level 2, Item 5"), NULL, emptyVariant, heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3 = pItem->GetSubItems();
		pItemsLevel3->Add(OLESTR("Level 3, Item 4"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 5"), NULL, emptyVariant, heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItem->Expand(VARIANT_FALSE);
	}
}

void __stdcall CMainDlg::ClickExtvwu(LPDISPATCH treeItem, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long hitTestDetails)
{
	CComQIPtr<ITreeViewItem> pItem = treeItem;
	if(pItem) {
		if((hitTestDetails & (htItemIcon | htItemLabel)) == 0) {
			if(pItem->GetHandle() == controls.extvwU->GetCaretItem(VARIANT_TRUE)->GetHandle()) {
				if((pItem->GetExpansionState() & esIsExpanded) == 0) {
					pItem->Expand(VARIANT_FALSE);
				}
			} else {
				controls.extvwU->PutRefCaretItem(VARIANT_FALSE, pItem);
			}
		}
	}
}

void __stdcall CMainDlg::ContextMenuExtvwu(LPDISPATCH treeItem, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/, VARIANT_BOOL* /*showDefaultMenu*/)
{
	CComQIPtr<ITreeViewItem> pItem = treeItem;
	if(pItem) {
		HMENU hMenu = LoadMenu(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MENU));
		HMENU hPopupMenu = GetSubMenu(hMenu, 0);

		POINT pt = {x, y};
		::ClientToScreen(static_cast<HWND>(LongToHandle(controls.extvwU->GethWnd())), &pt);
		TrackPopupMenuEx(hPopupMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, static_cast<HWND>(LongToHandle(controls.extvwU->GethWnd())), NULL);
		DestroyMenu(hPopupMenu);
		DestroyMenu(hMenu);
	}
}

void __stdcall CMainDlg::KeyDownExtvwu(short* keyCode, short /*shift*/)
{
	if(*keyCode == VK_F2) {
		controls.extvwU->GetCaretItem(VARIANT_TRUE)->StartLabelEditing();
	}
}

void __stdcall CMainDlg::RecreatedControlWindowExtvwu(long /*hWnd*/)
{
	InsertItems();
}
