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
		IDispEventImpl<IDC_EXTVWU, CMainDlg, &__uuidof(ExTVwLibU::_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 6>::DispEventUnadvise(controls.extvwU);
		controls.extvwU.Release();
	}
	if(controls.extvwA) {
		IDispEventImpl<IDC_EXTVWA, CMainDlg, &__uuidof(ExTVwLibA::_IExplorerTreeViewEvents), &LIBID_ExTVwLibA, 2, 6>::DispEventUnadvise(controls.extvwA);
		controls.extvwA.Release();
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

	controls.logEdit = GetDlgItem(IDC_EDITLOG);

	extvwUContainerWnd.SubclassWindow(GetDlgItem(IDC_EXTVWU));
	extvwUContainerWnd.QueryControl(__uuidof(ExTVwLibU::IExplorerTreeView), reinterpret_cast<LPVOID*>(&controls.extvwU));
	if(controls.extvwU) {
		IDispEventImpl<IDC_EXTVWU, CMainDlg, &__uuidof(ExTVwLibU::_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 6>::DispEventAdvise(controls.extvwU);
		extvwUWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.extvwU->GethWnd())));
		InsertItemsU();
	}

	extvwAContainerWnd.SubclassWindow(GetDlgItem(IDC_EXTVWA));
	extvwAContainerWnd.QueryControl(__uuidof(ExTVwLibA::IExplorerTreeView), reinterpret_cast<LPVOID*>(&controls.extvwA));
	if(controls.extvwA) {
		IDispEventImpl<IDC_EXTVWA, CMainDlg, &__uuidof(ExTVwLibA::_IExplorerTreeViewEvents), &LIBID_ExTVwLibA, 2, 6>::DispEventAdvise(controls.extvwA);
		extvwAWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.extvwA->GethWnd())));
		InsertItemsA();
	}

	if(CTheme::IsThemingSupported() && IsComctl32Version600OrNewer()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed && pfnIsAppThemed()) {
				SetWindowTheme(static_cast<HWND>(LongToHandle(controls.extvwU->GethWnd())), L"explorer", NULL);
				SetWindowTheme(static_cast<HWND>(LongToHandle(controls.extvwA->GethWnd())), L"explorer", NULL);
			}
			FreeLibrary(hThemeDLL);
		}
	}

	return TRUE;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(extvwaIsFocused) {
		if(controls.extvwA) {
			controls.extvwA->About();
		}
	} else {
		if(controls.extvwU) {
			controls.extvwU->About();
		}
	}
	return 0;
}

LRESULT CMainDlg::OnSetFocusU(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	extvwaIsFocused = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusA(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	extvwaIsFocused = TRUE;
	return 0;
}

void CMainDlg::AddLogEntry(CAtlString text)
{
	static int cLines = 0;
	static CAtlString oldText;

	cLines++;
	if(cLines > 50) {
		// delete the first line
		int pos = oldText.Find(TEXT("\r\n"));
		oldText = oldText.Mid(pos + lstrlen(TEXT("\r\n")), oldText.GetLength());
		cLines--;
	}
	oldText += text;
	oldText += TEXT("\r\n");

	controls.logEdit.SetWindowText(oldText);
	int l = oldText.GetLength();
	controls.logEdit.SetSel(l, l);
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void CMainDlg::InsertItemsA(void)
{
	controls.extvwA->PuthImageList(ExTVwLibA::ilItems, HandleToLong(controls.imageList.m_hImageList));
	int cImages = controls.imageList.GetImageCount();

	CComPtr<ExTVwLibA::ITreeViewItems> pItems = controls.extvwA->GetTreeItems();
	if(pItems) {
		_variant_t emptyVariant;
		emptyVariant.Clear();
		int iIcon = 0;

		CComPtr<ExTVwLibA::ITreeViewItem> pRootItem = pItems->Add(OLESTR("Level 1, Item 1"), NULL, emptyVariant, ExTVwLibA::heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;

		CComPtr<ExTVwLibA::ITreeViewItems> pItemsLevel2 = pRootItem->GetSubItems();
		pItemsLevel2->Add(OLESTR("Level 2, Item 1"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pRootItem->Expand(VARIANT_FALSE);

		CComPtr<ExTVwLibA::ITreeViewItem> pItem = pItemsLevel2->Add(OLESTR("Level 2, Item 2"), NULL, emptyVariant, ExTVwLibA::heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		CComPtr<ExTVwLibA::ITreeViewItems> pItemsLevel3 = pItem->GetSubItems();
		pItemsLevel3->Add(OLESTR("Level 3, Item 1"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 2"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 3"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItem->Expand(VARIANT_FALSE);

		pItemsLevel2->Add(OLESTR("Level 2, Item 3"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel2->Add(OLESTR("Level 2, Item 4"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;

		pItem = pItemsLevel2->Add(OLESTR("Level 2, Item 5"), NULL, emptyVariant, ExTVwLibA::heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3 = pItem->GetSubItems();
		pItemsLevel3->Add(OLESTR("Level 3, Item 4"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 5"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItem->Expand(VARIANT_FALSE);


		pRootItem = pItems->Add(OLESTR("Level 1, Item 1"), NULL, emptyVariant, ExTVwLibA::heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;

		pItemsLevel2 = pRootItem->GetSubItems();
		pItemsLevel2->Add(OLESTR("Level 2, Item 1"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pRootItem->Expand(VARIANT_FALSE);

		pItem = pItemsLevel2->Add(OLESTR("Level 2, Item 2"), NULL, emptyVariant, ExTVwLibA::heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3 = pItem->GetSubItems();
		pItemsLevel3->Add(OLESTR("Level 3, Item 1"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 2"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 3"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItem->Expand(VARIANT_FALSE);

		pItemsLevel2->Add(OLESTR("Level 2, Item 3"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel2->Add(OLESTR("Level 2, Item 4"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;

		pItem = pItemsLevel2->Add(OLESTR("Level 2, Item 5"), NULL, emptyVariant, ExTVwLibA::heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3 = pItem->GetSubItems();
		pItemsLevel3->Add(OLESTR("Level 3, Item 4"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 5"), NULL, emptyVariant, ExTVwLibA::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItem->Expand(VARIANT_FALSE);
	}
}

void CMainDlg::InsertItemsU(void)
{
	controls.extvwU->PuthImageList(ExTVwLibU::ilItems, HandleToLong(controls.imageList.m_hImageList));
	int cImages = controls.imageList.GetImageCount();

	CComPtr<ExTVwLibU::ITreeViewItems> pItems = controls.extvwU->GetTreeItems();
	if(pItems) {
		_variant_t emptyVariant;
		emptyVariant.Clear();
		int iIcon = 0;

		CComPtr<ExTVwLibU::ITreeViewItem> pRootItem = pItems->Add(OLESTR("Level 1, Item 1"), NULL, emptyVariant, ExTVwLibU::heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;

		CComPtr<ExTVwLibU::ITreeViewItems> pItemsLevel2 = pRootItem->GetSubItems();
		pItemsLevel2->Add(OLESTR("Level 2, Item 1"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pRootItem->Expand(VARIANT_FALSE);

		CComPtr<ExTVwLibU::ITreeViewItem> pItem = pItemsLevel2->Add(OLESTR("Level 2, Item 2"), NULL, emptyVariant, ExTVwLibU::heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		CComPtr<ExTVwLibU::ITreeViewItems> pItemsLevel3 = pItem->GetSubItems();
		pItemsLevel3->Add(OLESTR("Level 3, Item 1"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 2"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 3"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItem->Expand(VARIANT_FALSE);

		pItemsLevel2->Add(OLESTR("Level 2, Item 3"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel2->Add(OLESTR("Level 2, Item 4"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;

		pItem = pItemsLevel2->Add(OLESTR("Level 2, Item 5"), NULL, emptyVariant, ExTVwLibU::heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3 = pItem->GetSubItems();
		pItemsLevel3->Add(OLESTR("Level 3, Item 4"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 5"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItem->Expand(VARIANT_FALSE);


		pRootItem = pItems->Add(OLESTR("Level 1, Item 1"), NULL, emptyVariant, ExTVwLibU::heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;

		pItemsLevel2 = pRootItem->GetSubItems();
		pItemsLevel2->Add(OLESTR("Level 2, Item 1"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pRootItem->Expand(VARIANT_FALSE);

		pItem = pItemsLevel2->Add(OLESTR("Level 2, Item 2"), NULL, emptyVariant, ExTVwLibU::heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3 = pItem->GetSubItems();
		pItemsLevel3->Add(OLESTR("Level 3, Item 1"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 2"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 3"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItem->Expand(VARIANT_FALSE);

		pItemsLevel2->Add(OLESTR("Level 2, Item 3"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel2->Add(OLESTR("Level 2, Item 4"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;

		pItem = pItemsLevel2->Add(OLESTR("Level 2, Item 5"), NULL, emptyVariant, ExTVwLibU::heYes, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3 = pItem->GetSubItems();
		pItemsLevel3->Add(OLESTR("Level 3, Item 4"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItemsLevel3->Add(OLESTR("Level 3, Item 5"), NULL, emptyVariant, ExTVwLibU::heNo, iIcon, -2, -2, 0, VARIANT_FALSE, 1);
		iIcon = (iIcon + 1) % cImages;
		pItem->Expand(VARIANT_FALSE);
	}
}

void __stdcall CMainDlg::AbortedDragExtvwu()
{
	AddLogEntry(CAtlString(TEXT("ExTVwU_AbortedDrag")));

	controls.extvwU->PutRefDropHilitedItem(NULL);
}

void __stdcall CMainDlg::CaretChangedExtvwu(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, long caretChangeReason)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pPrevItem = previousCaretItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ExTVwU_CaretChanged: previousCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_CaretChanged: previousCaretItem=NULL");
	}
	CComQIPtr<ExTVwLibU::ITreeViewItem> pNewItem = newCaretItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newCaretItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", caretChangeReason=%i"), caretChangeReason);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CaretChangingExtvwu(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, long caretChangeReason, VARIANT_BOOL* cancelCaretChange)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pPrevItem = previousCaretItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ExTVwU_CaretChanging: previousCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_CaretChanging: previousCaretItem=NULL");
	}
	CComQIPtr<ExTVwLibU::ITreeViewItem> pNewItem = newCaretItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newCaretItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", caretChangeReason=%i, cancelCaretChange=%i"), caretChangeReason, *cancelCaretChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangedSortOrderExtvwu(long previousSortOrder, long newSortOrder)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_ChangedSortOrder: previousSortOrder=%i, newSortOrder=%i"), previousSortOrder, newSortOrder);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangingSortOrderExtvwu(long previousSortOrder, long newSortOrder, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_ChangingSortOrder: previousSortOrder=%i, newSortOrder=%i, cancelChange=%i"), previousSortOrder, newSortOrder, *cancelChange);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_Click: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_Click: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CollapsedItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL removedAllSubItems)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_CollapsedItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_CollapsedItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", removedAllSubItems=%i"), removedAllSubItems);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CollapsingItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL removingAllSubItems, VARIANT_BOOL* cancelCollapse)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_CollapsingItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_CollapsingItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", removingAllSubItems=%i, cancelCollapse=%i"), removingAllSubItems, *cancelCollapse);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CompareItemsExtvwu(LPDISPATCH firstItem, LPDISPATCH secondItem, long* result)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pFirstItem = firstItem;
	if(pFirstItem) {
		BSTR text = pFirstItem->GetText();
		str += TEXT("ExTVwU_CompareItems: firstItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_CompareItems: firstItem=NULL");
	}
	CComQIPtr<ExTVwLibU::ITreeViewItem> pSecondItem = secondItem;
	if(pSecondItem) {
		BSTR text = pSecondItem->GetText();
		str += TEXT(", secondItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", secondItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", result=%i"), *result);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_ContextMenu: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_ContextMenu: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedEditControlWindowExtvwu(long hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_CreatedEditControlWindow: hWndEdit=0x%X"), hWndEdit);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDrawExtvwu(LPDISPATCH treeItem, long* itemLevel, OLE_COLOR* textColor, OLE_COLOR* textBackColor, ExTVwLibU::CustomDrawStageConstants drawStage, ExTVwLibU::CustomDrawItemStateConstants itemState, long hDC, ExTVwLibU::RECTANGLE* drawingRectangle, ExTVwLibU::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_CustomDraw: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_CustomDraw: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemLevel=%i, textColor=0x%X, textBackColor=0x%X, drawStage=0x%X, itemState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), *itemLevel, *textColor, *textBackColor, drawStage, itemState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_DblClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_DblClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowExtvwu(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedEditControlWindowExtvwu(long hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_DestroyedEditControlWindow: hWndEdit=0x%X"), hWndEdit);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DragMouseMoveExtvwu(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails, VARIANT_BOOL* autoExpandItem, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_DragMouseMove: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_DragMouseMove: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=%i, autoExpandItem=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoExpandItem, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropExtvwu(LPDISPATCH dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_Drop: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_Drop: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=%i"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditChangeExtvwu()
{
	CAtlString str = TEXT("ExTVwU_EditChange - ");
	BSTR text = controls.extvwU->GetEditText().Detach();
	str += text;
	SysFreeString(text);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditClickExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditContextMenuExtvwu(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditContextMenu: button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditDblClickExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditGotFocusExtvwu()
{
	AddLogEntry(CAtlString(TEXT("ExTVwU_EditGotFocus")));
}

void __stdcall CMainDlg::EditKeyDownExtvwu(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditKeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditKeyPressExtvwu(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditKeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditKeyUpExtvwu(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditKeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditLostFocusExtvwu()
{
	AddLogEntry(CAtlString(TEXT("ExTVwU_EditLostFocus")));
}

void __stdcall CMainDlg::EditMClickExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditMClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMDblClickExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditMDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseDownExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditMouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseEnterExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditMouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseHoverExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditMouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseLeaveExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditMouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseMoveExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditMouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseUpExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditMouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseWheelExtvwu(short button, short shift, long x, long y, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditMouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditRClickExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditRClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditRDblClickExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditRDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditXClickExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditXClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditXDblClickExtvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_EditXDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ExpandedItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL expandedPartially)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_ExpandedItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_ExpandedItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", expandedPartially=%i"), expandedPartially);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ExpandingItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL expandingPartially, VARIANT_BOOL* cancelExpansion)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_ExpandingItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_ExpandingItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", expandingPartially=%i, cancelExpansion=%i"), expandingPartially, *cancelExpansion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeItemDataExtvwu(LPDISPATCH treeItem)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_FreeItemData: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_FreeItemData: treeItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::IncrementalSearchStringChangingExtvwu(BSTR currentSearchString, short keyCodeOfCharToBeAdded, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_IncrementalSearchStringChanging: currentSearchString=%s, keyCodeOfCharToBeAdded=%i, cancelChange=%i"), OLE2W(currentSearchString), keyCodeOfCharToBeAdded, *cancelChange);

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedItemExtvwu(LPDISPATCH treeItem)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_InsertedItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_InsertedItem: treeItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_InsertingItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_InsertingItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemAsynchronousDrawFailedExtvwu(LPDISPATCH treeItem, ExTVwLibU::FAILEDIMAGEDETAILS* imageDetails, ExTVwLibU::ImageDrawingFailureReasonConstants failureReason, ExTVwLibU::FailedAsyncDrawReturnValuesConstants* furtherProcessing, long* newImageToDraw)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_ItemAsynchronousDrawFailed: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_ItemAsynchronousDrawFailed: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", failureReason=0x%X, furtherProcessing=0x%X, newImageToDraw=%i"), failureReason, *furtherProcessing, *newImageToDraw);
	str += tmp;

	AddLogEntry(str);

	str.Format(TEXT("     imageDetails.hImageList=0x%X"), imageDetails->hImageList);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.hDC=0x%X"), imageDetails->hDC);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.IconIndex=%i"), imageDetails->IconIndex);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.OverlayIndex=%i"), imageDetails->OverlayIndex);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.DrawingStyle=0x%X"), imageDetails->DrawingStyle);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.DrawingEffects=0x%X"), imageDetails->DrawingEffects);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.BackColor=0x%X"), imageDetails->BackColor);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.ForeColor=0x%X"), imageDetails->ForeColor);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.EffectColor=0x%X"), imageDetails->EffectColor);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginDragExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_ItemBeginDrag: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_ItemBeginDrag: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	controls.extvwU->BeginDrag(controls.extvwU->CreateItemContainer(_variant_t(treeItem)), static_cast<OLE_HANDLE>(-1), NULL, NULL);
}

void __stdcall CMainDlg::ItemBeginRDragExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_ItemBeginRDrag: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_ItemBeginRDrag: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetDisplayInfoExtvwu(LPDISPATCH treeItem, long requestedInfo, long* iconIndex, long* selectedIconIndex, VARIANT_BOOL* hasExpando, long maxItemTextLength, BSTR* itemText, VARIANT_BOOL* dontAskAgain)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		CAtlString tmp;
		tmp.Format(TEXT("ExTVwU_ItemGetDisplayInfo: treeItem=0x%X"), pItem->GetHandle());
		str += tmp;
	} else {
		str += TEXT("ExTVwU_ItemGetDisplayInfo: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requestedInfo=%i, iconIndex=%i, selectedIconIndex=%i, hasExpando=%i, maxItemTextLength=%i, itemText=%s, dontAskAgain=%i"), requestedInfo, *iconIndex, *selectedIconIndex, *hasExpando, maxItemTextLength, OLE2W(*itemText), *dontAskAgain);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetInfoTipTextExtvwu(LPDISPATCH treeItem, long maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_ItemGetInfoTipText: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_ItemGetInfoTipText: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", maxInfoTipLength=%i, infoTipText=%s, abortToolTip=%i"), maxInfoTipLength, OLE2W(*infoTipText), *abortToolTip);
	str += tmp;

	AddLogEntry(str);

	CComBSTR infoTip(*infoTipText);
	if(infoTip.Length() > 0) {
		infoTip.Append(TEXT("\r\n"));
	}
	tmp.Format(TEXT("%sID: %i\r\nItemData: 0x%X"), infoTip, pItem->GetID(), pItem->GetItemData());
	*infoTipText = _bstr_t(tmp).Detach();
}

void __stdcall CMainDlg::ItemMouseEnterExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_ItemMouseEnter: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_ItemMouseEnter: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseLeaveExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_ItemMouseLeave: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_ItemMouseLeave: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemSelectionChangedExtvwu(LPDISPATCH treeItem)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_ItemSelectionChanged: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_ItemSelectionChanged: treeItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemSelectionChangingExtvwu(LPDISPATCH treeItem, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_ItemSelectionChanging: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_ItemSelectionChanging: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelChange=%i"), *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemSetTextExtvwu(LPDISPATCH treeItem, BSTR itemText)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		CAtlString tmp;
		tmp.Format(TEXT("ExTVwU_ItemSetText: treeItem=0x%X"), pItem->GetHandle());
		str += tmp;
	} else {
		str += TEXT("ExTVwU_ItemSetText: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemText=%s"), OLE2W(itemText));
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemStateImageChangedExtvwu(LPDISPATCH treeItem, long previousStateImageIndex, long newStateImageIndex, long causedBy)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_ItemStateImageChanged: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_ItemStateImageChanged: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousStateImageIndex=%i, newStateImageIndex=%i, causedBy=%i"), previousStateImageIndex, newStateImageIndex, causedBy);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemStateImageChangingExtvwu(LPDISPATCH treeItem, long previousStateImageIndex, long* newStateImageIndex, long causedBy, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_ItemStateImageChanging: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_ItemStateImageChanging: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousStateImageIndex=%i, newStateImageIndex=%i, causedBy=%i, cancelChange=%i"), previousStateImageIndex, *newStateImageIndex, causedBy, *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownExtvwu(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);

	if(*keyCode == VK_F2) {
		controls.extvwU->GetCaretItem(VARIANT_TRUE)->StartLabelEditing();
	} else if(*keyCode == VK_ESCAPE) {
		if(controls.extvwU->GetDraggedItems()) {
			controls.extvwU->EndDrag(VARIANT_TRUE);
		}
	}
}

void __stdcall CMainDlg::KeyPressExtvwu(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpExtvwu(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_MClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_MClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_MDblClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_MDblClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_MouseDown: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_MouseDown: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_MouseEnter: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_MouseEnter: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_MouseHover: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_MouseHover: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_MouseLeave: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_MouseLeave: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_MouseMove: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_MouseMove: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_MouseUp: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_MouseUp: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(button == 1/*vbLeftButton*/) {
		if(controls.extvwU->GetDraggedItems()) {
			// Are we within the client area?
			if(((ExTVwLibU::htAbove | ExTVwLibU::htBelow | ExTVwLibU::htToLeft | ExTVwLibU::htToRight) & hitTestDetails) == 0) {
				// yes
				controls.extvwU->EndDrag(VARIANT_FALSE);
			} else {
				// no
				controls.extvwU->EndDrag(VARIANT_TRUE);
			}
		}
	}
}

void __stdcall CMainDlg::MouseWheelExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_MouseWheel: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_MouseWheel: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, hitTestDetails, scrollAxis, wheelDelta);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragExtvwu(LPDISPATCH data, long performedEffect)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwU_OLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwU_OLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropExtvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=%i"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.extvwU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterExtvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails, VARIANT_BOOL* autoExpandItem, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=%i, autoExpandItem=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoExpandItem, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetExtvwu(long hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveExtvwu(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwU_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=%i"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetExtvwu(void)
{
	AddLogEntry(TEXT("ExTVwU_OLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::OLEDragMouseMoveExtvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails, VARIANT_BOOL* autoExpandItem, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=%i, autoExpandItem=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoExpandItem, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackExtvwu(long effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_OLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragExtvwu(VARIANT_BOOL pressedEscape, short button, short shift, long* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_OLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataExtvwu(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwU_OLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwU_OLEReceivedNewData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataExtvwu(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwU_OLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwU_OLESetData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEStartDragExtvwu(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwU_OLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwU_OLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_RClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_RClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_RDblClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_RDblClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowExtvwu(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ExTVwU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
	InsertItemsU();
}

void __stdcall CMainDlg::RemovedItemExtvwu(LPDISPATCH treeItem)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_RemovedItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_RemovedItem: treeItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_RemovingItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_RemovingItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RenamedItemExtvwu(LPDISPATCH treeItem, BSTR previousItemText, BSTR newItemText)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_RenamedItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_RenamedItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousItemText=%s, newItemText=%s"), OLE2W(previousItemText), OLE2W(newItemText));
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RenamingItemExtvwu(LPDISPATCH treeItem, BSTR previousItemText, BSTR newItemText, VARIANT_BOOL* cancelRenaming)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_RenamingItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_RenamingItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousItemText=%s, newItemText=%s, cancelRenaming=%i"), OLE2W(previousItemText), OLE2W(newItemText), *cancelRenaming);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowExtvwu()
{
	AddLogEntry(CAtlString(TEXT("ExTVwU_ResizedControlWindow")));
}

void __stdcall CMainDlg::SingleExpandingExtvwu(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, VARIANT_BOOL* dontChangePreviousItem, VARIANT_BOOL* dontChangeNewItem)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pPrevItem = previousCaretItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ExTVwU_SingleExpanding: previousCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_SingleExpanding: previousCaretItem=NULL");
	}
	CComQIPtr<ExTVwLibU::ITreeViewItem> pNewItem = newCaretItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newCaretItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", dontChangePreviousItem=%i, dontChangeNewItem=%i"), *dontChangePreviousItem, *dontChangeNewItem);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::StartingLabelEditingExtvwu(LPDISPATCH treeItem, VARIANT_BOOL* cancelEditing)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_StartingLabelEditing: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_StartingLabelEditing: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelEditing=%i"), *cancelEditing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_XClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_XClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibU::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwU_XDblClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwU_XDblClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}


void __stdcall CMainDlg::AbortedDragExtvwa()
{
	AddLogEntry(CAtlString(TEXT("ExTVwA_AbortedDrag")));

	controls.extvwA->PutRefDropHilitedItem(NULL);
}

void __stdcall CMainDlg::CaretChangedExtvwa(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, long caretChangeReason)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pPrevItem = previousCaretItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ExTVwA_CaretChanged: previousCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_CaretChanged: previousCaretItem=NULL");
	}
	CComQIPtr<ExTVwLibA::ITreeViewItem> pNewItem = newCaretItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newCaretItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", caretChangeReason=%i"), caretChangeReason);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CaretChangingExtvwa(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, long caretChangeReason, VARIANT_BOOL* cancelCaretChange)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pPrevItem = previousCaretItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ExTVwA_CaretChanging: previousCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_CaretChanging: previousCaretItem=NULL");
	}
	CComQIPtr<ExTVwLibA::ITreeViewItem> pNewItem = newCaretItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newCaretItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", caretChangeReason=%i, cancelCaretChange=%i"), caretChangeReason, *cancelCaretChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangedSortOrderExtvwa(long previousSortOrder, long newSortOrder)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_ChangedSortOrder: previousSortOrder=%i, newSortOrder=%i"), previousSortOrder, newSortOrder);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangingSortOrderExtvwa(long previousSortOrder, long newSortOrder, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_ChangingSortOrder: previousSortOrder=%i, newSortOrder=%i, cancelChange=%i"), previousSortOrder, newSortOrder, *cancelChange);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_Click: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_Click: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CollapsedItemExtvwa(LPDISPATCH treeItem, VARIANT_BOOL removedAllSubItems)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_CollapsedItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_CollapsedItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", removedAllSubItems=%i"), removedAllSubItems);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CollapsingItemExtvwa(LPDISPATCH treeItem, VARIANT_BOOL removingAllSubItems, VARIANT_BOOL* cancelCollapse)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_CollapsingItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_CollapsingItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", removingAllSubItems=%i, cancelCollapse=%i"), removingAllSubItems, *cancelCollapse);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CompareItemsExtvwa(LPDISPATCH firstItem, LPDISPATCH secondItem, long* result)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pFirstItem = firstItem;
	if(pFirstItem) {
		BSTR text = pFirstItem->GetText();
		str += TEXT("ExTVwA_CompareItems: firstItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_CompareItems: firstItem=NULL");
	}
	CComQIPtr<ExTVwLibA::ITreeViewItem> pSecondItem = secondItem;
	if(pSecondItem) {
		BSTR text = pSecondItem->GetText();
		str += TEXT(", secondItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", secondItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", result=%i"), *result);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_ContextMenu: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_ContextMenu: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedEditControlWindowExtvwa(long hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_CreatedEditControlWindow: hWndEdit=0x%X"), hWndEdit);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDrawExtvwa(LPDISPATCH treeItem, long* itemLevel, OLE_COLOR* textColor, OLE_COLOR* textBackColor, ExTVwLibA::CustomDrawStageConstants drawStage, ExTVwLibA::CustomDrawItemStateConstants itemState, long hDC, ExTVwLibA::RECTANGLE* drawingRectangle, ExTVwLibA::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_CustomDraw: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_CustomDraw: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemLevel=%i, textColor=0x%X, textBackColor=0x%X, drawStage=0x%X, itemState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), *itemLevel, *textColor, *textBackColor, drawStage, itemState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_DblClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_DblClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowExtvwa(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedEditControlWindowExtvwa(long hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_DestroyedEditControlWindow: hWndEdit=0x%X"), hWndEdit);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DragMouseMoveExtvwa(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails, VARIANT_BOOL* autoExpandItem, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_DragMouseMove: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_DragMouseMove: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=%i, autoExpandItem=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoExpandItem, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropExtvwa(LPDISPATCH dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_Drop: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_Drop: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=%i"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditChangeExtvwa()
{
	CAtlString str = TEXT("ExTVwA_EditChange - ");
	BSTR text = controls.extvwA->GetEditText().Detach();
	str += text;
	SysFreeString(text);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditClickExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditContextMenuExtvwa(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditContextMenu: button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditDblClickExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditGotFocusExtvwa()
{
	AddLogEntry(CAtlString(TEXT("ExTVwA_EditGotFocus")));
}

void __stdcall CMainDlg::EditKeyDownExtvwa(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditKeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditKeyPressExtvwa(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditKeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditKeyUpExtvwa(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditKeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditLostFocusExtvwa()
{
	AddLogEntry(CAtlString(TEXT("ExTVwA_EditLostFocus")));
}

void __stdcall CMainDlg::EditMClickExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditMClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMDblClickExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditMDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseDownExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditMouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseEnterExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditMouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseHoverExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditMouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseLeaveExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditMouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseMoveExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditMouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseUpExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditMouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseWheelExtvwa(short button, short shift, long x, long y, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditMouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditRClickExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditRClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditRDblClickExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditRDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditXClickExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditXClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditXDblClickExtvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_EditXDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ExpandedItemExtvwa(LPDISPATCH treeItem, VARIANT_BOOL expandedPartially)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_ExpandedItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_ExpandedItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", expandedPartially=%i"), expandedPartially);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ExpandingItemExtvwa(LPDISPATCH treeItem, VARIANT_BOOL expandingPartially, VARIANT_BOOL* cancelExpansion)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_ExpandingItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_ExpandingItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", expandingPartially=%i, cancelExpansion=%i"), expandingPartially, *cancelExpansion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeItemDataExtvwa(LPDISPATCH treeItem)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_FreeItemData: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_FreeItemData: treeItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::IncrementalSearchStringChangingExtvwa(BSTR currentSearchString, short keyCodeOfCharToBeAdded, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_IncrementalSearchStringChanging: currentSearchString=%s, keyCodeOfCharToBeAdded=%i, cancelChange=%i"), OLE2W(currentSearchString), keyCodeOfCharToBeAdded, *cancelChange);

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedItemExtvwa(LPDISPATCH treeItem)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_InsertedItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_InsertedItem: treeItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingItemExtvwa(LPDISPATCH treeItem, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_InsertingItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_InsertingItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemAsynchronousDrawFailedExtvwa(LPDISPATCH treeItem, ExTVwLibA::FAILEDIMAGEDETAILS* imageDetails, ExTVwLibA::ImageDrawingFailureReasonConstants failureReason, ExTVwLibA::FailedAsyncDrawReturnValuesConstants* furtherProcessing, long* newImageToDraw)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_ItemAsynchronousDrawFailed: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_ItemAsynchronousDrawFailed: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", failureReason=0x%X, furtherProcessing=0x%X, newImageToDraw=%i"), failureReason, *furtherProcessing, *newImageToDraw);
	str += tmp;

	AddLogEntry(str);

	str.Format(TEXT("     imageDetails.hImageList=0x%X"), imageDetails->hImageList);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.hDC=0x%X"), imageDetails->hDC);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.IconIndex=%i"), imageDetails->IconIndex);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.OverlayIndex=%i"), imageDetails->OverlayIndex);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.DrawingStyle=0x%X"), imageDetails->DrawingStyle);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.DrawingEffects=0x%X"), imageDetails->DrawingEffects);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.BackColor=0x%X"), imageDetails->BackColor);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.ForeColor=0x%X"), imageDetails->ForeColor);
	AddLogEntry(str);
	str.Format(TEXT("     imageDetails.EffectColor=0x%X"), imageDetails->EffectColor);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginDragExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_ItemBeginDrag: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_ItemBeginDrag: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	controls.extvwA->BeginDrag(controls.extvwA->CreateItemContainer(_variant_t(treeItem)), static_cast<OLE_HANDLE>(-1), NULL, NULL);
}

void __stdcall CMainDlg::ItemBeginRDragExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_ItemBeginRDrag: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_ItemBeginRDrag: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetDisplayInfoExtvwa(LPDISPATCH treeItem, long requestedInfo, long* iconIndex, long* selectedIconIndex, VARIANT_BOOL* hasExpando, long maxItemTextLength, BSTR* itemText, VARIANT_BOOL* dontAskAgain)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		CAtlString tmp;
		tmp.Format(TEXT("ExTVwA_ItemGetDisplayInfo: treeItem=0x%X"), pItem->GetHandle());
		str += tmp;
	} else {
		str += TEXT("ExTVwA_ItemGetDisplayInfo: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requestedInfo=%i, iconIndex=%i, selectedIconIndex=%i, hasExpando=%i, maxItemTextLength=%i, itemText=%s, dontAskAgain=%i"), requestedInfo, *iconIndex, *selectedIconIndex, *hasExpando, maxItemTextLength, OLE2W(*itemText), *dontAskAgain);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetInfoTipTextExtvwa(LPDISPATCH treeItem, long maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_ItemGetInfoTipText: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_ItemGetInfoTipText: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", maxInfoTipLength=%i, infoTipText=%s, abortToolTip=%i"), maxInfoTipLength, OLE2W(*infoTipText), *abortToolTip);
	str += tmp;

	AddLogEntry(str);

	CComBSTR infoTip(*infoTipText);
	if(infoTip.Length() > 0) {
		infoTip.Append(TEXT("\r\n"));
	}
	tmp.Format(TEXT("%sID: %i\r\nItemData: 0x%X"), infoTip, pItem->GetID(), pItem->GetItemData());
	*infoTipText = _bstr_t(tmp).Detach();
}

void __stdcall CMainDlg::ItemMouseEnterExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_ItemMouseEnter: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_ItemMouseEnter: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseLeaveExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_ItemMouseLeave: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_ItemMouseLeave: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemSelectionChangedExtvwa(LPDISPATCH treeItem)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_ItemSelectionChanged: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_ItemSelectionChanged: treeItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemSelectionChangingExtvwa(LPDISPATCH treeItem, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_ItemSelectionChanging: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_ItemSelectionChanging: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelChange=%i"), *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemSetTextExtvwa(LPDISPATCH treeItem, BSTR itemText)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		CAtlString tmp;
		tmp.Format(TEXT("ExTVwA_ItemSetText: treeItem=0x%X"), pItem->GetHandle());
		str += tmp;
	} else {
		str += TEXT("ExTVwA_ItemSetText: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemText=%s"), OLE2W(itemText));
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemStateImageChangedExtvwa(LPDISPATCH treeItem, long previousStateImageIndex, long newStateImageIndex, long causedBy)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_ItemStateImageChanged: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_ItemStateImageChanged: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousStateImageIndex=%i, newStateImageIndex=%i, causedBy=%i"), previousStateImageIndex, newStateImageIndex, causedBy);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemStateImageChangingExtvwa(LPDISPATCH treeItem, long previousStateImageIndex, long* newStateImageIndex, long causedBy, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_ItemStateImageChanging: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_ItemStateImageChanging: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousStateImageIndex=%i, newStateImageIndex=%i, causedBy=%i, cancelChange=%i"), previousStateImageIndex, *newStateImageIndex, causedBy, *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownExtvwa(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);

	if(*keyCode == VK_F2) {
		controls.extvwA->GetCaretItem(VARIANT_TRUE)->StartLabelEditing();
	} else if(*keyCode == VK_ESCAPE) {
		if(controls.extvwA->GetDraggedItems()) {
			controls.extvwA->EndDrag(VARIANT_TRUE);
		}
	}
}

void __stdcall CMainDlg::KeyPressExtvwa(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpExtvwa(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_MClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_MClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_MDblClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_MDblClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_MouseDown: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_MouseDown: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_MouseEnter: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_MouseEnter: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_MouseHover: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_MouseHover: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_MouseLeave: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_MouseLeave: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_MouseMove: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_MouseMove: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_MouseUp: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_MouseUp: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(button == 1/*vbLeftButton*/) {
		if(controls.extvwA->GetDraggedItems()) {
			// Are we within the client area?
			if(((ExTVwLibA::htAbove | ExTVwLibA::htBelow | ExTVwLibA::htToLeft | ExTVwLibA::htToRight) & hitTestDetails) == 0) {
				// yes
				controls.extvwA->EndDrag(VARIANT_FALSE);
			} else {
				// no
				controls.extvwA->EndDrag(VARIANT_TRUE);
			}
		}
	}
}

void __stdcall CMainDlg::MouseWheelExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_MouseWheel: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_MouseWheel: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, hitTestDetails, scrollAxis, wheelDelta);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragExtvwa(LPDISPATCH data, long performedEffect)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwA_OLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwA_OLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropExtvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=%i"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.extvwA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterExtvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails, VARIANT_BOOL* autoExpandItem, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=%i, autoExpandItem=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoExpandItem, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetExtvwa(long hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveExtvwa(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwA_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=%i"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetExtvwa(void)
{
	AddLogEntry(TEXT("ExTVwA_OLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::OLEDragMouseMoveExtvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails, VARIANT_BOOL* autoExpandItem, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=%i, autoExpandItem=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoExpandItem, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackExtvwa(long effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_OLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragExtvwa(VARIANT_BOOL pressedEscape, short button, short shift, long* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_OLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataExtvwa(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwA_OLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwA_OLEReceivedNewData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataExtvwa(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwA_OLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwA_OLESetData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEStartDragExtvwa(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExTVwA_OLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExTVwA_OLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_RClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_RClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_RDblClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_RDblClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowExtvwa(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ExTVwA_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
	InsertItemsA();
}

void __stdcall CMainDlg::RemovedItemExtvwa(LPDISPATCH treeItem)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_RemovedItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_RemovedItem: treeItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingItemExtvwa(LPDISPATCH treeItem, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_RemovingItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_RemovingItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RenamedItemExtvwa(LPDISPATCH treeItem, BSTR previousItemText, BSTR newItemText)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_RenamedItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_RenamedItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousItemText=%s, newItemText=%s"), OLE2W(previousItemText), OLE2W(newItemText));
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RenamingItemExtvwa(LPDISPATCH treeItem, BSTR previousItemText, BSTR newItemText, VARIANT_BOOL* cancelRenaming)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_RenamingItem: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_RenamingItem: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousItemText=%s, newItemText=%s, cancelRenaming=%i"), OLE2W(previousItemText), OLE2W(newItemText), *cancelRenaming);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowExtvwa()
{
	AddLogEntry(CAtlString(TEXT("ExTVwA_ResizedControlWindow")));
}

void __stdcall CMainDlg::SingleExpandingExtvwa(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, VARIANT_BOOL* dontChangePreviousItem, VARIANT_BOOL* dontChangeNewItem)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pPrevItem = previousCaretItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ExTVwA_SingleExpanding: previousCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_SingleExpanding: previousCaretItem=NULL");
	}
	CComQIPtr<ExTVwLibA::ITreeViewItem> pNewItem = newCaretItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newCaretItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", dontChangePreviousItem=%i, dontChangeNewItem=%i"), *dontChangePreviousItem, *dontChangeNewItem);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::StartingLabelEditingExtvwa(LPDISPATCH treeItem, VARIANT_BOOL* cancelEditing)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_StartingLabelEditing: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_StartingLabelEditing: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelEditing=%i"), *cancelEditing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_XClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_XClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExTVwLibA::ITreeViewItem> pItem = treeItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExTVwA_XDblClick: treeItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExTVwA_XDblClick: treeItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}
