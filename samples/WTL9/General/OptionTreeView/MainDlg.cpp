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
	controls.stateImageList.CreateFromImage(IDB_STATEIMAGES, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);

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
	controls.extvwU->PuthImageList(ilState, HandleToLong(controls.stateImageList.m_hImageList));

	CComPtr<ITreeViewItems> pItems = controls.extvwU->GetTreeItems();
	if(pItems) {
		_variant_t emptyVariant;
		emptyVariant.Clear();

		CComPtr<ITreeViewItem> pCheckBoxItem = pItems->Add(OLESTR("Check 1"), NULL, emptyVariant, heYes, 0, 0, 0, 1, VARIANT_FALSE, 1);
		pCheckBoxItem->PutStateImageIndex(1);

		CComPtr<ITreeViewItems> pOptionButtonItems = pCheckBoxItem->GetSubItems();
		pOptionButtonItems->Add(OLESTR("Option 1"), NULL, emptyVariant, heNo, 1, 1, 1, 2, VARIANT_FALSE, 1)->PutStateImageIndex(3);
		pOptionButtonItems->Add(OLESTR("Option 2"), NULL, emptyVariant, heNo, 2, 2, 2, 2, VARIANT_FALSE, 1)->PutStateImageIndex(3);
		pOptionButtonItems->Add(OLESTR("Option 3"), NULL, emptyVariant, heNo, 3, 3, 3, 2, VARIANT_FALSE, 1)->PutStateImageIndex(3);
		pCheckBoxItem->Expand(VARIANT_FALSE);

		pCheckBoxItem = pItems->Add(OLESTR("Check 2"), NULL, emptyVariant, heYes, 4, 4, 4, 1, VARIANT_FALSE, 1);
		pCheckBoxItem->PutStateImageIndex(1);

		pOptionButtonItems = pCheckBoxItem->GetSubItems();
		pOptionButtonItems->Add(OLESTR("Option 1"), NULL, emptyVariant, heNo, 5, 5, 5, 2, VARIANT_FALSE, 1)->PutStateImageIndex(3);
		pOptionButtonItems->Add(OLESTR("Option 2"), NULL, emptyVariant, heNo, 6, 6, 6, 2, VARIANT_FALSE, 1)->PutStateImageIndex(3);
		pOptionButtonItems->Add(OLESTR("Option 3"), NULL, emptyVariant, heNo, 7, 7, 7, 2, VARIANT_FALSE, 1)->PutStateImageIndex(3);
		pCheckBoxItem->Expand(VARIANT_FALSE);
	}
}

void __stdcall CMainDlg::ItemStateImageChangingExtvwu(LPDISPATCH treeItem, long previousStateImageIndex, long* newStateImageIndex, long causedBy, VARIANT_BOOL* /*cancelChange*/)
{
	if(causedBy != siccbCode) {
		CComQIPtr<ITreeViewItem> pItem = treeItem;
		switch(pItem->GetItemData()) {
			case 1:
				// CheckBox
				if(previousStateImageIndex == 1) {
					// check the item
					*newStateImageIndex = 2;
				} else {
					// un-check the item
					*newStateImageIndex = 1;
				}
				break;
			case 2:
				// OptionButton
				if(previousStateImageIndex == 3) {
					// uncheck all items that have the same parent
					CComPtr<ITreeViewItems> pItems = pItem->GetParentItem()->GetSubItems();
					pItems->PutFilterType(fpStateImageIndex, ftIncluding);
					_variant_t filter;
					filter.Clear();
					CComSafeArray<VARIANT> arr;
					arr.Create(1, 1);
					arr.SetAt(1, _variant_t(4));
					filter.parray = arr.Detach();
					filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: ExplorerTreeView expects an array of VARIANTs!
					pItems->PutFilter(fpStateImageIndex, filter);

					CComQIPtr<IEnumVARIANT> pEnum = pItems;
					pEnum->Reset();     // NOTE: It's important to call Reset()!
					ULONG ul = 0;
					_variant_t v;
					v.Clear();
					while(pEnum->Next(1, &v, &ul) == S_OK) {
						CComQIPtr<ITreeViewItem> pItem = v.pdispVal;
						pItem->PutStateImageIndex(3);
					}

					// check the item
					*newStateImageIndex = 4;
				} else {
					// leave the item checked
					*newStateImageIndex = 4;
				}
				break;
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowExtvwu(long /*hWnd*/)
{
	InsertItems();
}
