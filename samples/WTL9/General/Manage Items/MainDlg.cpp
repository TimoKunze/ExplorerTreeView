// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "EditItemDlg.h"
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
		PrepareTreeView();
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

void CMainDlg::PrepareTreeView(void)
{
	controls.extvwU->PuthImageList(ilItems, HandleToLong(controls.imageList.m_hImageList));
}

void __stdcall CMainDlg::ClickExtvwu(LPDISPATCH treeItem, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long hitTestDetails)
{
	CComQIPtr<ITreeViewItem> pItem = treeItem;
	if(!pItem) {
		if((hitTestDetails & (htItemStateImage | htItemExpando)) == 0) {
			// clear caret
			controls.extvwU->PutRefCaretItem(VARIANT_TRUE, NULL);
		}
	}
}

void __stdcall CMainDlg::ContextMenuExtvwu(LPDISPATCH treeItem, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/, VARIANT_BOOL* /*showDefaultMenu*/)
{
	CComQIPtr<ITreeViewItem> pItem = treeItem;
	UINT selectedCmd = 0;
	if(pItem) {
		HMENU hMenu = LoadMenu(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_ITEMACTIONSMENU));
		HMENU hPopupMenu = GetSubMenu(hMenu, 0);

		POINT pt = {x, y};
		::ClientToScreen(static_cast<HWND>(LongToHandle(controls.extvwU->GethWnd())), &pt);
		selectedCmd = TrackPopupMenuEx(hPopupMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, static_cast<HWND>(LongToHandle(controls.extvwU->GethWnd())), NULL);
		DestroyMenu(hPopupMenu);
		DestroyMenu(hMenu);
	} else {
		HMENU hMenu = LoadMenu(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_TREEVIEWACTIONSMENU));
		HMENU hPopupMenu = GetSubMenu(hMenu, 0);

		if(controls.extvwU->GetTreeItems()->Count() == 0) {
			EnableMenuItem(hPopupMenu, ID_TREEACTION_REMOVEALL, MF_DISABLED | MF_GRAYED | MF_BYCOMMAND);
		}

		POINT pt = {x, y};
		::ClientToScreen(static_cast<HWND>(LongToHandle(controls.extvwU->GethWnd())), &pt);
		selectedCmd = TrackPopupMenuEx(hPopupMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, static_cast<HWND>(LongToHandle(controls.extvwU->GethWnd())), NULL);
		DestroyMenu(hPopupMenu);
		DestroyMenu(hMenu);
	}

	switch(selectedCmd) {
		case ID_ITEMACTION_ADDSUBITEM: {
			ITEMPROPS itemProps;
			itemProps.text.Format(TEXT("Item %i, Level %i"), controls.extvwU->GetTreeItems()->Count() + 1, pItem->GetLevel() + 1);

			CEditItemDlg editItemDlg;
			INT_PTR ret = editItemDlg.DoModal(*this, reinterpret_cast<LPARAM>(&itemProps));
			if(ret == IDOK) {
				_variant_t emptyVariant;
				emptyVariant.Clear();
				CComPtr<ITreeViewItem> pNewItem = controls.extvwU->GetTreeItems()->Add(_bstr_t(itemProps.text), pItem, emptyVariant, (itemProps.hasExpando ? heYes : heNo), itemProps.iconIndex, itemProps.selectedIconIndex, itemProps.iconIndex, itemProps.itemData, VARIANT_FALSE, itemProps.heightIncrement);
				pNewItem->PutBold(itemProps.bold ? VARIANT_TRUE : VARIANT_FALSE);
				pNewItem->PutGhosted(itemProps.ghosted ? VARIANT_TRUE : VARIANT_FALSE);
				controls.extvwU->PutRefCaretItem(VARIANT_TRUE, pNewItem);
			}
			break;
		}
		case ID_ITEMACTION_ADDITEMABOVE: {
			ITEMPROPS itemProps;
			itemProps.text.Format(TEXT("Item %i, Level %i"), controls.extvwU->GetTreeItems()->Count() + 1, pItem->GetLevel());

			CEditItemDlg editItemDlg;
			INT_PTR ret = editItemDlg.DoModal(*this, reinterpret_cast<LPARAM>(&itemProps));
			if(ret == IDOK) {
				_variant_t ia;
				ia.vt = VT_DISPATCH;
				CComPtr<ITreeViewItem> pPrevItem = pItem->GetPreviousSiblingItem();
				if(pPrevItem) {
					pPrevItem->QueryInterface(IID_PPV_ARGS(&ia.pdispVal));
				} else {
					ia.pdispVal = NULL;
				}
				CComPtr<ITreeViewItem> pNewItem = controls.extvwU->GetTreeItems()->Add(_bstr_t(itemProps.text), pItem->GetParentItem(), ia, (itemProps.hasExpando ? heYes : heNo), itemProps.iconIndex, itemProps.selectedIconIndex, itemProps.iconIndex, itemProps.itemData, VARIANT_FALSE, itemProps.heightIncrement);
				pNewItem->PutBold(itemProps.bold ? VARIANT_TRUE : VARIANT_FALSE);
				pNewItem->PutGhosted(itemProps.ghosted ? VARIANT_TRUE : VARIANT_FALSE);
				controls.extvwU->PutRefCaretItem(VARIANT_TRUE, pNewItem);
			}
			break;
		}
		case ID_ITEMACTION_ADDITEMBELOW: {
			ITEMPROPS itemProps;
			itemProps.text.Format(TEXT("Item %i, Level %i"), controls.extvwU->GetTreeItems()->Count() + 1, pItem->GetLevel());

			CEditItemDlg editItemDlg;
			INT_PTR ret = editItemDlg.DoModal(*this, reinterpret_cast<LPARAM>(&itemProps));
			if(ret == IDOK) {
				_variant_t ia;
				ia.vt = VT_DISPATCH;
				pItem->QueryInterface(IID_PPV_ARGS(&ia.pdispVal));
				CComPtr<ITreeViewItem> pNewItem = controls.extvwU->GetTreeItems()->Add(_bstr_t(itemProps.text), pItem->GetParentItem(), ia, (itemProps.hasExpando ? heYes : heNo), itemProps.iconIndex, itemProps.selectedIconIndex, itemProps.iconIndex, itemProps.itemData, VARIANT_FALSE, itemProps.heightIncrement);
				pNewItem->PutBold(itemProps.bold ? VARIANT_TRUE : VARIANT_FALSE);
				pNewItem->PutGhosted(itemProps.ghosted ? VARIANT_TRUE : VARIANT_FALSE);
				controls.extvwU->PutRefCaretItem(VARIANT_TRUE, pNewItem);
			}
			break;
		}
		case ID_ITEMACTION_EDIT: {
			ITEMPROPS itemProps;
			itemProps.text = OLE2W(pItem->GetText());
			itemProps.bold = (pItem->GetBold() != VARIANT_FALSE);
			itemProps.ghosted = (pItem->GetGhosted() != VARIANT_FALSE);
			itemProps.hasExpando = (pItem->GetHasExpando() == heYes);
			itemProps.iconIndex = pItem->GetIconIndex();
			itemProps.itemData = pItem->GetItemData();
			itemProps.selectedIconIndex = pItem->GetSelectedIconIndex();
			itemProps.heightIncrement = pItem->GetHeightIncrement();

			CEditItemDlg editItemDlg;
			INT_PTR ret = editItemDlg.DoModal(*this, reinterpret_cast<LPARAM>(&itemProps));
			if(ret == IDOK) {
				pItem->PutText(_bstr_t(itemProps.text));
				pItem->PutBold(itemProps.bold ? VARIANT_TRUE : VARIANT_FALSE);
				pItem->PutGhosted(itemProps.ghosted ? VARIANT_TRUE : VARIANT_FALSE);
				pItem->PutHasExpando(itemProps.hasExpando ? heYes : heNo);
				pItem->PutIconIndex(itemProps.iconIndex);
				pItem->PutItemData(itemProps.itemData);
				pItem->PutSelectedIconIndex(itemProps.selectedIconIndex);
				pItem->PutHeightIncrement(itemProps.heightIncrement);
			}
			break;
		}
		case ID_ITEMACTION_REMOVEITEM:
			controls.extvwU->GetTreeItems()->Remove(pItem->GetHandle(), iitHandle);
			break;
		case ID_TREEACTION_ADDITEM: {
			ITEMPROPS itemProps;
			itemProps.text.Format(TEXT("Item %i, Level 0"), controls.extvwU->GetTreeItems()->Count() + 1);

			CEditItemDlg editItemDlg;
			INT_PTR ret = editItemDlg.DoModal(*this, reinterpret_cast<LPARAM>(&itemProps));
			if(ret == IDOK) {
				_variant_t emptyVariant;
				emptyVariant.Clear();
				CComPtr<ITreeViewItem> pNewItem = controls.extvwU->GetTreeItems()->Add(_bstr_t(itemProps.text), NULL, emptyVariant, (itemProps.hasExpando ? heYes : heNo), itemProps.iconIndex, itemProps.selectedIconIndex, itemProps.iconIndex, itemProps.itemData, VARIANT_FALSE, itemProps.heightIncrement);
				pNewItem->PutBold(itemProps.bold ? VARIANT_TRUE : VARIANT_FALSE);
				pNewItem->PutGhosted(itemProps.ghosted ? VARIANT_TRUE : VARIANT_FALSE);
				controls.extvwU->PutRefCaretItem(VARIANT_TRUE, pNewItem);
			}
			break;
		}
		case ID_TREEACTION_REMOVEALL:
			controls.extvwU->GetTreeItems()->RemoveAll();
			break;
	}
}

void __stdcall CMainDlg::DblClickExtvwu(LPDISPATCH treeItem, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/)
{
	CComQIPtr<ITreeViewItem> pItem = treeItem;
	if(pItem) {
		ITEMPROPS itemProps;
		itemProps.text = OLE2W(pItem->GetText());
		itemProps.bold = (pItem->GetBold() != VARIANT_FALSE);
		itemProps.ghosted = (pItem->GetGhosted() != VARIANT_FALSE);
		itemProps.hasExpando = (pItem->GetHasExpando() == heYes);
		itemProps.iconIndex = pItem->GetIconIndex();
		itemProps.itemData = pItem->GetItemData();
		itemProps.selectedIconIndex = pItem->GetSelectedIconIndex();
		itemProps.heightIncrement = pItem->GetHeightIncrement();

		CEditItemDlg editItemDlg;
		INT_PTR ret = editItemDlg.DoModal(*this, reinterpret_cast<LPARAM>(&itemProps));
		if(ret == IDOK) {
			pItem->PutText(_bstr_t(itemProps.text));
			pItem->PutBold(itemProps.bold ? VARIANT_TRUE : VARIANT_FALSE);
			pItem->PutGhosted(itemProps.ghosted ? VARIANT_TRUE : VARIANT_FALSE);
			pItem->PutHasExpando(itemProps.hasExpando ? heYes : heNo);
			pItem->PutIconIndex(itemProps.iconIndex);
			pItem->PutItemData(itemProps.itemData);
			pItem->PutSelectedIconIndex(itemProps.selectedIconIndex);
			pItem->PutHeightIncrement(itemProps.heightIncrement);
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowExtvwu(long /*hWnd*/)
{
	PrepareTreeView();
}
