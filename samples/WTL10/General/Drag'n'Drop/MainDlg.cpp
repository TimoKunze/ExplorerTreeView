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

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	controls.oleDDCheck = GetDlgItem(IDC_OLEDDCHECK);
	controls.oleDDCheck.SetCheck(BST_CHECKED);

	extvwUWnd.SubclassWindow(GetDlgItem(IDC_EXTVWU));
	extvwUWnd.QueryControl(__uuidof(IExplorerTreeView), reinterpret_cast<LPVOID*>(&controls.extvwU));
	if(controls.extvwU) {
		DispEventAdvise(controls.extvwU);
	}

	HMODULE hMod = LoadLibrary(TEXT("shell32.dll"));
	if(hMod) {
		if(IsComctl32Version600OrNewer()) {
			controls.imageList.Create(16, 16, ILC_COLOR32 | ILC_MASK, 10, 0);
			/* WORKAROUND: SysTreeView32 seems to have problems with imagelists that have a back color of
			               CLR_NONE. In this case it may happen that the icon's background isn't filled with
			               the control's background color when drawing the icon. If the icon doesn't have an
			               alpha channel, this doesn't really matter. But if it has, the semi-transparent pixels
			               become darker and darker with each paint cycle.
			               So ensure the imagelist has a background color. */
			//controls.imageList.SetBkColor(CLR_NONE);
			OLE_COLOR clr = controls.extvwU->GetBackColor();
			COLORREF bkClr;
			OleTranslateColor(clr, NULL, &bkClr);
			controls.imageList.SetBkColor(bkClr);

			int resIDsXP[10] = {1489, 516, 26, 185, 266, 51, 275, 126, 37, 1029};
			int resIDsVista[10] = {1013, 20, 37, 180, 301, 60, 207, 133, 1128, 1593};
			int* resIDs;
			if(RunTimeHelper::IsVista()) {
				resIDs = resIDsVista;
			} else {
				resIDs = resIDsXP;
			}

			for(int i = 0; i < 10; i++) {
				HRSRC hResource = FindResource(hMod, MAKEINTRESOURCE(resIDs[i]), RT_ICON);
				if(hResource) {
					HGLOBAL hMem = LoadResource(hMod, hResource);
					if(hMem) {
						LPVOID pIconData = LockResource(hMem);
						if(pIconData) {
							HICON hIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hMod, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR | LR_CREATEDIBSECTION);
							if(hIcon) {
								controls.imageList.AddIcon(hIcon);
								DestroyIcon(hIcon);
							}
						}
					}
				}
			}
		} else {
			controls.imageList.Create(16, 16, ILC_COLOR24 | ILC_MASK, 10, 0);
			controls.imageList.SetBkColor(CLR_NONE);

			int resIDsXP[10] = {1486, 513, 22, 181, 263, 48, 272, 122, 34, 1026};
			int resIDsVista[10] = {1010, 17, 34, 177, 298, 57, 204, 129, 1124, 1590};
			int* resIDs;
			if(RunTimeHelper::IsVista()) {
				resIDs = resIDsVista;
			} else {
				resIDs = resIDsXP;
			}

			for(int i = 0; i < 10; i++) {
				HRSRC hResource = FindResource(hMod, MAKEINTRESOURCE(resIDs[i]), RT_ICON);
				if(hResource) {
					HGLOBAL hMem = LoadResource(hMod, hResource);
					if(hMem) {
						LPVOID pIconData = LockResource(hMem);
						if(pIconData) {
							HICON hIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hMod, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR | LR_CREATEDIBSECTION);
							if(hIcon) {
								controls.imageList.AddIcon(hIcon);
								DestroyIcon(hIcon);
							}
						}
					}
				}
			}
		}
		FreeLibrary(hMod);
	}

	if(controls.extvwU) {
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

LRESULT CMainDlg::OnOptionsAeroDragImages(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(GetMenuState(hMenu, ID_OPTIONS_AERODRAGIMAGES, MF_BYCOMMAND) & MF_CHECKED) {
		controls.extvwU->PutOLEDragImageStyle(odistClassic);
	} else {
		controls.extvwU->PutOLEDragImageStyle(odistAeroIfAvailable);
	}
	UpdateMenu();
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

InsertMarkPositionConstants CMainDlg::CalculateInsertMarkPosition(ITreeViewItem* pDropTarget, long yToItemTop)
{
	if(pDropTarget) {
		OLE_YPOS_PIXELS yTop = 0, yBottom = 0;
		pDropTarget->GetRectangle(irtItem, NULL, &yTop, NULL, &yBottom);
		LONG cy = yBottom - yTop;
		if(yToItemTop < (cy / 3)) {
			return impBefore;
		} else if(yToItemTop > (cy * 2 / 3)) {
			return impAfter;
		}
	}
	return impNowhere;
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

BOOL CMainDlg::IsSameItem(ITreeViewItem* pItem1, ITreeViewItem* pItem2)
{
	if(!pItem1) {
		return (pItem2 == NULL);
	} else if(!pItem2) {
		return FALSE;
	} else {
		return (pItem1->GetHandle() == pItem2->GetHandle());
	}
}

void CMainDlg::UpdateMenu(void)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(controls.extvwU->GetOLEDragImageStyle() == odistAeroIfAvailable) {
		CheckMenuItem(hMenu, ID_OPTIONS_AERODRAGIMAGES, MF_BYCOMMAND | MF_CHECKED);
	} else {
		CheckMenuItem(hMenu, ID_OPTIONS_AERODRAGIMAGES, MF_BYCOMMAND | MF_UNCHECKED);
	}
}

void __stdcall CMainDlg::AbortedDragExtvwu()
{
	controls.extvwU->PutRefDropHilitedItem(NULL);
	controls.extvwU->SetInsertMarkPosition(impNowhere, NULL);
	hLastDropTarget = NULL;
}

void __stdcall CMainDlg::DragMouseMoveExtvwu(LPDISPATCH* dropTarget, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long yToItemTop, long /*hitTestDetails*/, VARIANT_BOOL* /*autoExpandItem*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/)
{
	CComQIPtr<ITreeViewItem> pDropTarget = *dropTarget;
	OLE_HANDLE hNewDropTarget;
	if(!pDropTarget) {
		hNewDropTarget = NULL;
		//controls.extvwU->PutShowDragImage(VARIANT_FALSE);
	} else {
		hNewDropTarget = pDropTarget->GetHandle();
		//controls.extvwU->PutShowDragImage(VARIANT_TRUE);
	}
	if(hNewDropTarget != hLastDropTarget) {
		controls.extvwU->PutRefDropHilitedItem(pDropTarget);
		hLastDropTarget = hNewDropTarget;
	}

	InsertMarkPositionConstants newInsertMarkRelativePosition = CalculateInsertMarkPosition(pDropTarget, yToItemTop);
	InsertMarkPositionConstants lastInsertMarkRelativePosition;
	CComPtr<ITreeViewItem> pItem = NULL;
	controls.extvwU->GetInsertMarkPosition(&lastInsertMarkRelativePosition, &pItem);

	if((newInsertMarkRelativePosition != lastInsertMarkRelativePosition) || !IsSameItem(pItem, pDropTarget)) {
		controls.extvwU->SetInsertMarkPosition(newInsertMarkRelativePosition, pDropTarget);
	}
}

void __stdcall CMainDlg::DropExtvwu(LPDISPATCH dropTarget, short /*button*/, short /*shift*/, long x, long y, long /*yToItemTop*/, long /*hitTestDetails*/)
{
	if(bRightDrag) {
		HMENU hMenu = LoadMenu(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_DRAGMENU));
		HMENU hPopupMenu = GetSubMenu(hMenu, 0);

		POINT pt = {x, y};
		::ClientToScreen(static_cast<HWND>(LongToHandle(controls.extvwU->GethWnd())), &pt);
		UINT selectedCmd = TrackPopupMenuEx(hPopupMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, static_cast<HWND>(LongToHandle(controls.extvwU->GethWnd())), NULL);
		DestroyMenu(hPopupMenu);
		DestroyMenu(hMenu);

		switch(selectedCmd) {
			case ID_ACTIONS_CANCEL:
				AbortedDragExtvwu();
				return;
				break;
			case ID_ACTIONS_COPYITEM:
				MessageBox(TEXT("TODO: Copy the item."), TEXT("Command not implemented"), MB_ICONINFORMATION);
				AbortedDragExtvwu();
				return;
				break;
			case ID_ACTIONS_MOVEITEM:
				// fall through
				break;
		}
	}

	// move the item
	InsertMarkPositionConstants insertionMark;
	CComPtr<ITreeViewItem> pItem = NULL;
	controls.extvwU->GetInsertMarkPosition(&insertionMark, &pItem);

	CComQIPtr<ITreeViewItem> pDropTarget = dropTarget;
	CComPtr<ITreeViewItem> pParentItem = NULL;
	CComPtr<ITreeViewItem> pInsertAfter = NULL;
	switch(insertionMark) {
		case impAfter:
			pParentItem = pDropTarget->GetParentItem();
			pInsertAfter = pDropTarget;
			break;
		case impBefore:
			pParentItem = pDropTarget->GetParentItem();
			pInsertAfter = pDropTarget->GetPreviousSiblingItem();
			break;
		case impNowhere:
			pParentItem = pDropTarget;
			break;
	}

	if(pParentItem) {
		pParentItem->PutHasExpando(heYes);
	}
	/* If we move an item together with its sub-items, it may happen that we move the parent before
	   the children, which will raise an error when the children shall be moved. It's the easiest
	   solution to ignore this error.
	   The same goes for the situation that the 'DraggedItems' object contains items that can't be
	   moved to the target (e. g. because they're parent items of the target). */
	CComQIPtr<IEnumVARIANT> pEnum = controls.extvwU->GetDraggedItems();
	pEnum->Reset();     // NOTE: It's important to call Reset()!
	ULONG ul = 0;
	_variant_t v;
	v.Clear();
	while(pEnum->Next(1, &v, &ul) == S_OK) {
		CComQIPtr<ITreeViewItem> pItem = v.pdispVal;
		if(insertionMark == impNowhere) {
			_variant_t ia;
			ia.vt = VT_I4;
			ia.lVal = iaLast;
			pItem->Move(pParentItem, ia);
		} else {
			_variant_t ia;
			ia.vt = VT_DISPATCH;
			if(pInsertAfter) {
				pInsertAfter->QueryInterface(IID_PPV_ARGS(&ia.pdispVal));
			} else {
				ia.pdispVal = NULL;
			}
			pItem->Move(pParentItem, ia);
		}
	}

	controls.extvwU->PutRefDropHilitedItem(NULL);
	controls.extvwU->SetInsertMarkPosition(impNowhere, NULL);
	hLastDropTarget = NULL;
}

void __stdcall CMainDlg::ItemBeginDragExtvwu(LPDISPATCH treeItem, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/)
{
	bRightDrag = FALSE;
	hLastDropTarget = NULL;
	if(controls.oleDDCheck.GetCheck() == BST_CHECKED) {
		IOLEDataObject* p = NULL;
		controls.extvwU->OLEDrag(reinterpret_cast<long*>(&p), odeCopyOrMove, static_cast<OLE_HANDLE>(-1), controls.extvwU->CreateItemContainer(_variant_t(treeItem)), -1);
	} else {
		controls.extvwU->BeginDrag(controls.extvwU->CreateItemContainer(_variant_t(treeItem)), static_cast<OLE_HANDLE>(-1), NULL, NULL);
	}
}

void __stdcall CMainDlg::ItemBeginRDragExtvwu(LPDISPATCH treeItem, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/)
{
	bRightDrag = TRUE;
	hLastDropTarget = NULL;
	if(controls.oleDDCheck.GetCheck() == BST_CHECKED) {
		IOLEDataObject* p = NULL;
		controls.extvwU->OLEDrag(reinterpret_cast<long*>(&p), odeCopyOrMove, static_cast<OLE_HANDLE>(-1), controls.extvwU->CreateItemContainer(_variant_t(treeItem)), -1);
	} else {
		controls.extvwU->BeginDrag(controls.extvwU->CreateItemContainer(_variant_t(treeItem)), static_cast<OLE_HANDLE>(-1), NULL, NULL);
	}
}

void __stdcall CMainDlg::KeyDownExtvwu(short* keyCode, short /*shift*/)
{
	if(*keyCode == VK_ESCAPE) {
		if(controls.extvwU->GetDraggedItems()) {
			controls.extvwU->EndDrag(VARIANT_TRUE);
		}
	}
}

void __stdcall CMainDlg::MouseUpExtvwu(LPDISPATCH /*treeItem*/, short button, short /*shift*/, long /*x*/, long /*y*/, long hitTestDetails)
{
	if(controls.oleDDCheck.GetCheck() == BST_UNCHECKED) {
		if(((button == 1) && !bRightDrag) || ((button == 2) && bRightDrag)) {
			if(controls.extvwU->GetDraggedItems()) {
				// Are we within the client area?
				if(((htAbove | htBelow | htToLeft | htToRight) & hitTestDetails) == 0) {
					// yes
					controls.extvwU->EndDrag(VARIANT_FALSE);
				} else {
					// no
					controls.extvwU->EndDrag(VARIANT_TRUE);
				}
			}
		}
	}
}

void __stdcall CMainDlg::OLECompleteDragExtvwu(LPDISPATCH data, long /*performedEffect*/)
{
	CComQIPtr<IDataObject> pData = data;
	if(pData) {
		FORMATETC format = {CF_TARGETCLSID, NULL, 1, -1, TYMED_HGLOBAL};
		if(pData->QueryGetData(&format) == S_OK) {
			STGMEDIUM storageMedium = {0};
			if(pData->GetData(&format, &storageMedium) == S_OK) {
				LPGUID pCLSIDOfTarget = reinterpret_cast<LPGUID>(GlobalLock(storageMedium.hGlobal));
				if(*pCLSIDOfTarget == CLSID_RecycleBin) {
					// remove the dragged items
					controls.extvwU->GetDraggedItems()->RemoveAll(VARIANT_TRUE);
				}
				GlobalUnlock(storageMedium.hGlobal);
				ReleaseStgMedium(&storageMedium);
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragDropExtvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long /*x*/, long /*y*/, long /*yToItemTop*/, long /*hitTestDetails*/)
{
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		if(pData->GetFormat(CF_DIB, -1, 1) != VARIANT_FALSE) {
			pDraggedPicture = pData->GetData(CF_DIB, -1, 1);
		} else {
			CComBSTR str = NULL;
			if(pData->GetFormat(CF_UNICODETEXT, -1, 1) != VARIANT_FALSE) {
				str = pData->GetData(CF_UNICODETEXT, -1, 1).bstrVal;
			} else if(pData->GetFormat(CF_TEXT, -1, 1) != VARIANT_FALSE) {
				str = pData->GetData(CF_TEXT, -1, 1).bstrVal;
			} else if(pData->GetFormat(CF_OEMTEXT, -1, 1) != VARIANT_FALSE) {
				str = pData->GetData(CF_OEMTEXT, -1, 1).bstrVal;
			}

			if(str != TEXT("")) {
				// insert a new item
				InsertMarkPositionConstants insertionMark;
				CComPtr<ITreeViewItem> pItem = NULL;
				controls.extvwU->GetInsertMarkPosition(&insertionMark, &pItem);

				CComQIPtr<ITreeViewItem> pDropTarget = *dropTarget;
				CComPtr<ITreeViewItem> pParentItem = NULL;
				CComPtr<ITreeViewItem> pInsertAfter = NULL;
				switch(insertionMark) {
					case impAfter:
						pParentItem = pDropTarget->GetParentItem();
						pInsertAfter = pDropTarget;
						break;
					case impBefore:
						pParentItem = pDropTarget->GetParentItem();
						pInsertAfter = pDropTarget->GetPreviousSiblingItem();
						break;
					case impNowhere:
						pParentItem = pDropTarget;
						break;
				}

				if(pParentItem) {
					pParentItem->PutHasExpando(heYes);
				}
				if(insertionMark == impNowhere) {
					_variant_t ia;
					ia.vt = VT_I4;
					ia.lVal = iaLast;
					controls.extvwU->GetTreeItems()->Add(_bstr_t(str), pParentItem, ia, heNo, 1, 1, 1, 0, VARIANT_FALSE, 1);
				} else {
					_variant_t ia;
					ia.vt = VT_DISPATCH;
					if(pInsertAfter) {
						pInsertAfter->QueryInterface(IID_PPV_ARGS(&ia.pdispVal));
					} else {
						ia.pdispVal = NULL;
					}
					controls.extvwU->GetTreeItems()->Add(_bstr_t(str), pParentItem, ia, heNo, 1, 1, 1, 0, VARIANT_FALSE, 1);
				}

			} else if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
				// insert a new item for each file/folder
				InsertMarkPositionConstants insertionMark;
				CComPtr<ITreeViewItem> pItem = NULL;
				controls.extvwU->GetInsertMarkPosition(&insertionMark, &pItem);

				CComQIPtr<ITreeViewItem> pDropTarget = *dropTarget;
				CComPtr<ITreeViewItem> pParentItem = NULL;
				CComPtr<ITreeViewItem> pInsertAfter = NULL;
				switch(insertionMark) {
					case impAfter:
						pParentItem = pDropTarget->GetParentItem();
						pInsertAfter = pDropTarget;
						break;
					case impBefore:
						pParentItem = pDropTarget->GetParentItem();
						pInsertAfter = pDropTarget->GetPreviousSiblingItem();
						break;
					case impNowhere:
						pParentItem = pDropTarget;
						break;
				}

				if(pParentItem) {
					pParentItem->PutHasExpando(heYes);
				}
				_variant_t v = pData->GetData(CF_HDROP, -1, 1);
				CComSafeArray<BSTR> arr;
				arr.Attach(v.parray);
				for(LONG i = arr.GetLowerBound(); i <= arr.GetUpperBound(); i++) {
					if(insertionMark == impNowhere) {
						_variant_t ia;
						ia.vt = VT_I4;
						ia.lVal = iaLast;
						controls.extvwU->GetTreeItems()->Add(_bstr_t(arr.GetAt(i)), pParentItem, ia, heNo, 1, 1, 1, 0, VARIANT_FALSE, 1);
					} else {
						_variant_t ia;
						ia.vt = VT_DISPATCH;
						if(pInsertAfter) {
							pInsertAfter->QueryInterface(IID_PPV_ARGS(&ia.pdispVal));
						} else {
							ia.pdispVal = NULL;
						}
						pInsertAfter = controls.extvwU->GetTreeItems()->Add(_bstr_t(arr.GetAt(i)), pParentItem, ia, heNo, 1, 1, 1, 0, VARIANT_FALSE, 1);
					}
				}
				arr.Detach();
			}
		}
	}

	controls.extvwU->PutRefDropHilitedItem(NULL);
	controls.extvwU->SetInsertMarkPosition(impNowhere, NULL);
	hLastDropTarget = NULL;

	if(shift & 2/*vbCtrlMask*/) {
		*effect = odeCopy;
	}
	if(shift & 4/*vbAltMask*/) {
		*effect = odeLink;
	}
}

void __stdcall CMainDlg::OLEDragEnterExtvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long /*x*/, long /*y*/, long yToItemTop, long /*hitTestDetails*/, VARIANT_BOOL* /*autoExpandItem*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/)
{
	_bstr_t dropActionDescription;
	DropDescriptionIconConstants dropDescriptionIcon = ddiNoDrop;
	_bstr_t dropTargetDescription;

	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		CComQIPtr<ITreeViewItem> pDropTarget = *dropTarget;
		OLE_HANDLE hNewDropTarget;
		if(!pDropTarget) {
			hNewDropTarget = NULL;
		} else {
			dropTargetDescription = pDropTarget->GetText();
			hNewDropTarget = pDropTarget->GetHandle();
		}
		if(hNewDropTarget != hLastDropTarget) {
			controls.extvwU->PutRefDropHilitedItem(pDropTarget);
			hLastDropTarget = hNewDropTarget;
		}

		InsertMarkPositionConstants newInsertMarkRelativePosition = CalculateInsertMarkPosition(pDropTarget, yToItemTop);
		InsertMarkPositionConstants lastInsertMarkRelativePosition;
		CComPtr<ITreeViewItem> pItem = NULL;
		controls.extvwU->GetInsertMarkPosition(&lastInsertMarkRelativePosition, &pItem);

		if((newInsertMarkRelativePosition != lastInsertMarkRelativePosition) || !IsSameItem(pItem, pDropTarget)) {
			controls.extvwU->SetInsertMarkPosition(newInsertMarkRelativePosition, pDropTarget);
		}

		*effect = odeCopy;
		if(shift & 1/*vbShiftMask*/) {
			*effect = odeMove;
		}
		if(shift & 2/*vbCtrlMask*/) {
			*effect = odeCopy;
		}
		if(shift & 4/*vbAltMask*/) {
			*effect = odeLink;
		}

		if(!pDropTarget) {
			dropDescriptionIcon = ddiNoDrop;
			dropActionDescription = TEXT("Cannot place here");
		} else {
			switch(*effect) {
				case odeMove:
					switch(newInsertMarkRelativePosition) {
						case impAfter:
							dropActionDescription = TEXT("Move after %1");
							break;
						case impBefore:
							dropActionDescription = TEXT("Move before %1");
							break;
						default:
							dropActionDescription = TEXT("Move to %1");
							break;
					}
					dropDescriptionIcon = ddiMove;
					break;
				case odeLink:
					switch(newInsertMarkRelativePosition) {
						case impAfter:
							dropActionDescription = TEXT("Create link after %1");
							break;
						case impBefore:
							dropActionDescription = TEXT("Create link before %1");
							break;
						default:
							dropActionDescription = TEXT("Create link in %1");
							break;
					}
					dropDescriptionIcon = ddiLink;
					break;
				default:
					switch(newInsertMarkRelativePosition) {
						case impAfter:
							dropActionDescription = TEXT("Insert copy after %1");
							break;
						case impBefore:
							dropActionDescription = TEXT("Insert copy before %1");
							break;
						default:
							dropActionDescription = TEXT("Copy to %1");
							break;
					}
					dropDescriptionIcon = ddiCopy;
					break;
			}
		}
		pData->SetDropDescription(dropTargetDescription, dropActionDescription, dropDescriptionIcon);
	}
}

void __stdcall CMainDlg::OLEDragLeaveExtvwu(LPDISPATCH /*data*/, LPDISPATCH /*dropTarget*/, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*yToItemTop*/, long /*hitTestDetails*/)
{
	controls.extvwU->PutRefDropHilitedItem(NULL);
	controls.extvwU->SetInsertMarkPosition(impNowhere, NULL);
	hLastDropTarget = NULL;
}

void __stdcall CMainDlg::OLEDragMouseMoveExtvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long /*x*/, long /*y*/, long yToItemTop, long /*hitTestDetails*/, VARIANT_BOOL* /*autoExpandItem*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/)
{
	_bstr_t dropActionDescription;
	DropDescriptionIconConstants dropDescriptionIcon = ddiNoDrop;
	_bstr_t dropTargetDescription;

	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		CComQIPtr<ITreeViewItem> pDropTarget = *dropTarget;
		OLE_HANDLE hNewDropTarget;
		if(!pDropTarget) {
			hNewDropTarget = NULL;
			//controls.extvwU->PutShowDragImage(VARIANT_FALSE);
		} else {
			dropTargetDescription = pDropTarget->GetText();
			hNewDropTarget = pDropTarget->GetHandle();
			//controls.extvwU->PutShowDragImage(VARIANT_TRUE);
		}
		if(hNewDropTarget != hLastDropTarget) {
			controls.extvwU->PutRefDropHilitedItem(pDropTarget);
			hLastDropTarget = hNewDropTarget;
		}

		InsertMarkPositionConstants newInsertMarkRelativePosition = CalculateInsertMarkPosition(pDropTarget, yToItemTop);
		InsertMarkPositionConstants lastInsertMarkRelativePosition;
		CComPtr<ITreeViewItem> pItem = NULL;
		controls.extvwU->GetInsertMarkPosition(&lastInsertMarkRelativePosition, &pItem);

		if((newInsertMarkRelativePosition != lastInsertMarkRelativePosition) || !IsSameItem(pItem, pDropTarget)) {
			controls.extvwU->SetInsertMarkPosition(newInsertMarkRelativePosition, pDropTarget);
		}

		*effect = odeCopy;
		if(shift & 1/*vbShiftMask*/) {
			*effect = odeMove;
		}
		if(shift & 2/*vbCtrlMask*/) {
			*effect = odeCopy;
		}
		if(shift & 4/*vbAltMask*/) {
			*effect = odeLink;
		}

		if(!pDropTarget) {
			dropDescriptionIcon = ddiNoDrop;
			dropActionDescription = TEXT("Cannot place here");
		} else {
			switch(*effect) {
				case odeMove:
					switch(newInsertMarkRelativePosition) {
						case impAfter:
							dropActionDescription = TEXT("Move after %1");
							break;
						case impBefore:
							dropActionDescription = TEXT("Move before %1");
							break;
						default:
							dropActionDescription = TEXT("Move to %1");
							break;
					}
					dropDescriptionIcon = ddiMove;
					break;
				case odeLink:
					switch(newInsertMarkRelativePosition) {
						case impAfter:
							dropActionDescription = TEXT("Create link after %1");
							break;
						case impBefore:
							dropActionDescription = TEXT("Create link before %1");
							break;
						default:
							dropActionDescription = TEXT("Create link in %1");
							break;
					}
					dropDescriptionIcon = ddiLink;
					break;
				default:
					switch(newInsertMarkRelativePosition) {
						case impAfter:
							dropActionDescription = TEXT("Insert copy after %1");
							break;
						case impBefore:
							dropActionDescription = TEXT("Insert copy before %1");
							break;
						default:
							dropActionDescription = TEXT("Copy to %1");
							break;
					}
					dropDescriptionIcon = ddiCopy;
					break;
			}
		}
		pData->SetDropDescription(dropTargetDescription, dropActionDescription, dropDescriptionIcon);
	}
}

void __stdcall CMainDlg::OLESetDataExtvwu(LPDISPATCH data, long formatID, long /*index*/, long /*dataOrViewAspect*/)
{
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		switch(formatID) {
			case CF_TEXT:
			case CF_OEMTEXT:
			case CF_UNICODETEXT: {
				CComQIPtr<IEnumVARIANT> pEnum = controls.extvwU->GetDraggedItems();
				pEnum->Reset();     // NOTE: It's important to call Reset()!
				ULONG ul = 0;
				_variant_t v;
				v.Clear();
				CComBSTR str = L"";
				while(pEnum->Next(1, &v, &ul) == S_OK) {
					CComQIPtr<ITreeViewItem> pItem = v.pdispVal;
					str.AppendBSTR(pItem->GetText());
					str.Append(TEXT("\r\n"));
				}
				pData->SetData(formatID, _variant_t(str), -1, 1);
				break;
			}

			case CF_BITMAP:
			case CF_DIB:
			case CF_DIBV5: {
				_variant_t v;
				v.vt = VT_DISPATCH;
				if(pDraggedPicture) {
					pDraggedPicture->QueryInterface(IID_PPV_ARGS(&v.pdispVal));
				} else {
					v.pdispVal = NULL;
				}
				pData->SetData(formatID, v, -1, 1);
				break;
			}

			case CF_PALETTE:
				if(pDraggedPicture) {
					OLE_HANDLE h = NULL;
					pDraggedPicture->get_hPal(&h);
					pData->SetData(formatID, _variant_t(h), -1, 1);
				}
				break;
		}
	}
}

void __stdcall CMainDlg::OLEStartDragExtvwu(LPDISPATCH data)
{
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		_variant_t invalidVariant;
		invalidVariant.vt = VT_ERROR;

		pData->SetData(CF_TEXT, invalidVariant, -1, 1);
		pData->SetData(CF_OEMTEXT, invalidVariant, -1, 1);
		pData->SetData(CF_UNICODETEXT, invalidVariant, -1, 1);
		pData->SetData(CF_HDROP, invalidVariant, -1, 1);
		if(pDraggedPicture) {
			pData->SetData(CF_PALETTE, invalidVariant, -1, 1);
			pData->SetData(CF_BITMAP, invalidVariant, -1, 1);
			pData->SetData(CF_DIB, invalidVariant, -1, 1);
			pData->SetData(CF_DIBV5, invalidVariant, -1, 1);
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowExtvwu(long /*hWnd*/)
{
	InsertItems();
}
