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
	if(controls.extvwU1) {
		IDispEventImpl<IDC_EXTVWU1, CMainDlg, &__uuidof(_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 6>::DispEventUnadvise(controls.extvwU1);
		controls.extvwU1.Release();
	}
	if(controls.extvwU2) {
		IDispEventImpl<IDC_EXTVWU2, CMainDlg, &__uuidof(_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 6>::DispEventUnadvise(controls.extvwU2);
		controls.extvwU2.Release();
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

	controls.imageList.CreateFromImage(IDB_ICONS, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.enabledCheck = GetDlgItem(IDC_ENABLEDCHK);

	extvwU1Wnd.SubclassWindow(GetDlgItem(IDC_EXTVWU1));
	extvwU1Wnd.QueryControl(__uuidof(IExplorerTreeView), reinterpret_cast<LPVOID*>(&controls.extvwU1));
	if(controls.extvwU1) {
		IDispEventImpl<IDC_EXTVWU1, CMainDlg, &__uuidof(_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 6>::DispEventAdvise(controls.extvwU1);
		InsertItems1();
	}
	extvwU2Wnd.SubclassWindow(GetDlgItem(IDC_EXTVWU2));
	extvwU2Wnd.QueryControl(__uuidof(IExplorerTreeView), reinterpret_cast<LPVOID*>(&controls.extvwU2));
	if(controls.extvwU2) {
		IDispEventImpl<IDC_EXTVWU2, CMainDlg, &__uuidof(_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 6>::DispEventAdvise(controls.extvwU2);
		InsertItems2();
	}

	if(CTheme::IsThemingSupported() && IsComctl32Version600OrNewer()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed && pfnIsAppThemed()) {
				SetWindowTheme(static_cast<HWND>(LongToHandle(controls.extvwU1->GethWnd())), L"explorer", NULL);
				SetWindowTheme(static_cast<HWND>(LongToHandle(controls.extvwU2->GethWnd())), L"explorer", NULL);
			}
			FreeLibrary(hThemeDLL);
		}
	}

	return TRUE;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.extvwU1) {
		controls.extvwU1->About();
	}
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void CMainDlg::InsertItems1(void)
{
	controls.extvwU1->PuthImageList(ilItems, HandleToLong(controls.imageList.m_hImageList));
	int cImages = controls.imageList.GetImageCount();

	CComPtr<ITreeViewItems> pItems = controls.extvwU1->GetTreeItems();
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
	}
}

void CMainDlg::InsertItems2(void)
{
	controls.extvwU2->PuthImageList(ilItems, HandleToLong(controls.imageList.m_hImageList));
	int cImages = controls.imageList.GetImageCount();

	CComPtr<ITreeViewItems> pItems = controls.extvwU2->GetTreeItems();
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
		pRootItem->PutItemData(1);

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
		pItem->PutItemData(1);

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
		pItem->PutItemData(1);
	}
}

LRESULT CMainDlg::OnBnClickedEnabledchk(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.enabledCheck.GetCheck() == BST_CHECKED) {
		CComPtr<ITreeViewItem> pItem = controls.extvwU1->GetCaretItem(VARIANT_TRUE);
		pItem->PutItemData(0);
		controls.extvwU1->Refresh();
	} else if(controls.enabledCheck.GetCheck() == BST_UNCHECKED) {
		CComPtr<ITreeViewItem> pItem = controls.extvwU1->GetCaretItem(VARIANT_TRUE);
		pItem->PutItemData(1);
		controls.extvwU1->Refresh();
	}
	return 0;
}

void __stdcall CMainDlg::CaretChangedExtvwu1(LPDISPATCH /*previousCaretItem*/, LPDISPATCH newCaretItem, long /*caretChangeReason*/)
{
	CComQIPtr<ITreeViewItem> pNewCaretItem = newCaretItem;
	if(pNewCaretItem) {
		if(pNewCaretItem->GetItemData() == 0) {
			controls.enabledCheck.SetCheck(BST_CHECKED);
		} else if(pNewCaretItem->GetItemData() == 1) {
			controls.enabledCheck.SetCheck(BST_UNCHECKED);
		}
	}
}

void __stdcall CMainDlg::CaretChangingExtvwu2(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, long /*caretChangeReason*/, VARIANT_BOOL* cancelCaretChange)
{
	CComQIPtr<ITreeViewItem> pPrevCaretItem = previousCaretItem;
	CComQIPtr<ITreeViewItem> pNewCaretItem = newCaretItem;
	if(pNewCaretItem) {
		if(pNewCaretItem->GetItemData() == 1) {
			// this is a disabled item
			// select its first child instead
			BOOL bFirstChild = FALSE;
			CComPtr<ITreeViewItem> pFirstSubItem = pNewCaretItem->GetFirstSubItem();

			if(!pPrevCaretItem) {
				bFirstChild = TRUE;
			} else if(pFirstSubItem) {
				if(pFirstSubItem->GetHandle() != pPrevCaretItem->GetHandle()) {
					bFirstChild = TRUE;
				}
			}

			if(bFirstChild) {
				if(!pFirstSubItem) {
					// select the next sibling item
					CComPtr<ITreeViewItem> pNextSiblItem = pNewCaretItem->GetNextSiblingItem();
					if(pNextSiblItem) {
						controls.extvwU2->PutRefCaretItem(VARIANT_TRUE, pNextSiblItem);
					}
				} else {
					controls.extvwU2->PutRefCaretItem(VARIANT_TRUE, pFirstSubItem);
				}
			}
			*cancelCaretChange = VARIANT_TRUE;
		}
	}
}

void __stdcall CMainDlg::CollapsingItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL /*removingAllSubItems*/, VARIANT_BOOL* cancelCollapse)
{
	CComQIPtr<ITreeViewItem> pItem = treeItem;
	if(pItem) {
		if(pItem->GetItemData() == 1) {
			*cancelCollapse = VARIANT_TRUE;
		}
	}
}

void __stdcall CMainDlg::CustomDrawExtvwu(LPDISPATCH treeItem, long* /*itemLevel*/, OLE_COLOR* textColor, OLE_COLOR* textBackColor, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants itemState, long /*hDC*/, RECTANGLE* /*drawingRectangle*/, CustomDrawReturnValuesConstants* furtherProcessing)
{
	switch(drawStage) {
		case cdsPrePaint:
			*furtherProcessing = cdrvNotifyItemDraw;
			break;
		case cdsItemPrePaint: {
			CComQIPtr<ITreeViewItem> pItem = treeItem;
			if(pItem->GetItemData() == 1) {
				// this is a disabled item
				if(itemState & cdisSelected) {
				// use the following line instead of the last line, if you've set AllowDragDrop to True
				//if(*textBackColor == GetSysColor(COLOR_HIGHLIGHT)) {
					if(itemState & cdisFocus) {
						*textBackColor = GetSysColor(COLOR_3DFACE);
						*textColor = GetSysColor(COLOR_WINDOWTEXT);
					}
					// use the following code instead of the last If block, if you've set AllowDragDrop to True
					//*textBackColor = GetSysColor(COLOR_3DFACE);
					//*textColor = GetSysColor(COLOR_WINDOWTEXT);
				} else {
					*textColor = GetSysColor(COLOR_GRAYTEXT);
				}

				*furtherProcessing = cdrvNewFont;
			}
			break;
		}
	}
}

void __stdcall CMainDlg::ExpandingItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL /*expandingPartially*/, VARIANT_BOOL* cancelExpansion)
{
	CComQIPtr<ITreeViewItem> pItem = treeItem;
	if(pItem) {
		if(pItem->GetItemData() == 1) {
			*cancelExpansion = VARIANT_TRUE;
		}
	}
}

void __stdcall CMainDlg::ItemGetInfoTipTextExtvwu(LPDISPATCH treeItem, long /*maxInfoTipLength*/, BSTR* /*infoTipText*/, VARIANT_BOOL* abortToolTip)
{
	CComQIPtr<ITreeViewItem> pItem = treeItem;
	if(pItem) {
		if(pItem->GetItemData() == 1) {
			*abortToolTip = VARIANT_TRUE;
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowExtvwu1(long /*hWnd*/)
{
	InsertItems1();
}

void __stdcall CMainDlg::RecreatedControlWindowExtvwu2(long /*hWnd*/)
{
	InsertItems2();
}

void __stdcall CMainDlg::StartingLabelEditingExtvwu(LPDISPATCH treeItem, VARIANT_BOOL* cancelEditing)
{
	CComQIPtr<ITreeViewItem> pItem = treeItem;
	if(pItem) {
		if(pItem->GetItemData() == 1) {
			*cancelEditing = VARIANT_TRUE;
		}
	}
}
