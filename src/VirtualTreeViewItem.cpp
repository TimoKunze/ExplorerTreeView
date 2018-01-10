// VirtualTreeViewItem.cpp: A wrapper for non-existing treeview items.

#include "stdafx.h"
#include "VirtualTreeViewItem.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualTreeViewItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualTreeViewItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


VirtualTreeViewItem::Properties::~Properties()
{
	if(settings.itemex.pszText != LPSTR_TEXTCALLBACK) {
		SECUREFREE(settings.itemex.pszText);
	}
	if(pOwnerExTvw) {
		pOwnerExTvw->Release();
	}
}


HWND VirtualTreeViewItem::Properties::GetExTvwHWnd(void)
{
	ATLASSUME(pOwnerExTvw);

	OLE_HANDLE handle = NULL;
	pOwnerExTvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


HTREEITEM VirtualTreeViewItem::GetParentItem(void)
{
	return properties.settings.hParent;
}

HTREEITEM VirtualTreeViewItem::GetPreviousItem(void)
{
	return properties.settings.hInsertAfter;
}

HRESULT VirtualTreeViewItem::LoadState(LPTVINSERTSTRUCT pSource)
{
	ATLASSERT_POINTER(pSource, TVINSERTSTRUCT);

	SECUREFREE(properties.settings.itemex.pszText);
	properties.settings = *pSource;
	if(properties.settings.itemex.mask & TVIF_TEXT) {
		// duplicate the item's text
		if(properties.settings.itemex.pszText != LPSTR_TEXTCALLBACK) {
			properties.settings.itemex.cchTextMax = lstrlen(pSource->itemex.pszText);
			properties.settings.itemex.pszText = reinterpret_cast<LPTSTR>(malloc((properties.settings.itemex.cchTextMax + 1) * sizeof(TCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopy(properties.settings.itemex.pszText, properties.settings.itemex.cchTextMax + 1, pSource->itemex.pszText)));
		}
	}

	return S_OK;
}

HRESULT VirtualTreeViewItem::SaveState(LPTVINSERTSTRUCT pTarget)
{
	ATLASSERT_POINTER(pTarget, TVINSERTSTRUCT);

	SECUREFREE(pTarget->itemex.pszText);
	*pTarget = properties.settings;
	if(pTarget->itemex.mask & TVIF_TEXT) {
		// duplicate the item's text
		if(pTarget->itemex.pszText != LPSTR_TEXTCALLBACK) {
			pTarget->itemex.pszText = reinterpret_cast<LPTSTR>(malloc((pTarget->itemex.cchTextMax + 1) * sizeof(TCHAR)));
			ATLASSERT(pTarget->itemex.pszText);
			if(pTarget->itemex.pszText) {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pTarget->itemex.pszText, pTarget->itemex.cchTextMax + 1, properties.settings.itemex.pszText)));
			}
		}
	}

	return S_OK;
}

void VirtualTreeViewItem::SetOwner(ExplorerTreeView* pOwner)
{
	if(properties.pOwnerExTvw) {
		properties.pOwnerExTvw->Release();
	}
	properties.pOwnerExTvw = pOwner;
	if(properties.pOwnerExTvw) {
		properties.pOwnerExTvw->AddRef();
	}
}


STDMETHODIMP VirtualTreeViewItem::get_Bold(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.itemex.state & TVIS_BOLD);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_DropHilited(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.itemex.state & TVIS_DROPHILITED);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_ExpandedIconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_EXPANDEDIMAGE) {
		*pValue = properties.settings.itemex.iExpandedImage;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_ExpansionState(ExpansionStateConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ExpansionStateConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_STATE) {
		*pValue = static_cast<ExpansionStateConstants>(properties.settings.itemex.state & (TVIS_EXPANDED | TVIS_EXPANDEDONCE | TVIS_EXPANDPARTIAL));
	} else {
		*pValue = esNeverWasExpanded;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_Ghosted(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.itemex.state & TVIS_CUT);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_Grayed(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_STATEEX) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.itemex.uStateEx & TVIS_EX_DISABLED);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_HasExpando(HasExpandoConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HasExpandoConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_CHILDREN) {
		*pValue = static_cast<HasExpandoConstants>(properties.settings.itemex.cChildren);
	} else {
		*pValue = heNo;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_HeightIncrement(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_INTEGRAL) {
		*pValue = properties.settings.itemex.iIntegral;
	} else {
		*pValue = 1;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_IMAGE) {
		*pValue = properties.settings.itemex.iImage;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_ItemData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_PARAM) {
		*pValue = static_cast<LONG>(properties.settings.itemex.lParam);
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_Level(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	// get the parent item's level
	CComPtr<ITreeViewItem> pParentItem = NULL;
	get_ParentItem(&pParentItem);
	if(pParentItem) {
		pParentItem->get_Level(pValue);
		pParentItem = NULL;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_NextSiblingItem(ITreeViewItem** ppNextItem)
{
	ATLASSERT_POINTER(ppNextItem, ITreeViewItem*);
	if(!ppNextItem) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	HTREEITEM hNextItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(properties.settings.hInsertAfter)));
	ClassFactory::InitTreeItem(hNextItem, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppNextItem));
	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_OverlayIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_STATE) {
		*pValue = ((properties.settings.itemex.state & TVIS_OVERLAYMASK) >> 8);
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_ParentItem(ITreeViewItem** ppParentItem)
{
	ATLASSERT_POINTER(ppParentItem, ITreeViewItem*);
	if(!ppParentItem) {
		return E_POINTER;
	}

	if(properties.pOwnerExTvw) {
		ClassFactory::InitTreeItem(properties.settings.hParent, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppParentItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP VirtualTreeViewItem::get_PreviousSiblingItem(ITreeViewItem** ppPreviousItem)
{
	ATLASSERT_POINTER(ppPreviousItem, ITreeViewItem*);
	if(!ppPreviousItem) {
		return E_POINTER;
	}

	if(properties.pOwnerExTvw) {
		ClassFactory::InitTreeItem(properties.settings.hInsertAfter, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppPreviousItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP VirtualTreeViewItem::get_Selected(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.itemex.state & TVIS_SELECTED);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_SelectedIconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_SELECTEDIMAGE) {
		*pValue = properties.settings.itemex.iSelectedImage;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_StateImageIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_STATE) {
		*pValue = ((properties.settings.itemex.state & TVIS_STATEIMAGEMASK) >> 12);
	} else {
		*pValue = 1;
	}

	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_TEXT) {
		if(properties.settings.itemex.pszText == LPSTR_TEXTCALLBACK) {
			*pValue = NULL;
		} else {
			*pValue = _bstr_t(properties.settings.itemex.pszText).Detach();
		}
	} else {
		*pValue = SysAllocString(L"");
	}
	return S_OK;
}

STDMETHODIMP VirtualTreeViewItem::get_Virtual(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.itemex.mask & TVIF_STATEEX) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.itemex.uStateEx & TVIS_EX_FLAT);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}


STDMETHODIMP VirtualTreeViewItem::GetPath(BSTR* pSeparator, BSTR* pPath)
{
	ATLASSERT_POINTER(pPath, BSTR);
	if(pPath== NULL) {
		return E_POINTER;
	}

	// retrieve the parent item's path
	CComPtr<ITreeViewItem> pParentItem = NULL;
	get_ParentItem(&pParentItem);
	if(pParentItem) {
		pParentItem->GetPath(pSeparator, pPath);
		pParentItem = NULL;
	}

	// now attach our text
	CAtlString path = TEXT("");
	if(properties.settings.itemex.mask & TVIF_TEXT) {
		if(properties.settings.itemex.pszText != LPSTR_TEXTCALLBACK) {
			path = CAtlString(properties.settings.itemex.pszText);
		}
	}
	CAtlString separator = TEXT("\\");
	if(pSeparator) {
		separator = *pSeparator;
	}
	if((*pPath != NULL) && (lstrcmp(COLE2CT(*pPath), TEXT("")) != 0)) {
		path = CAtlString(*pPath) + separator + path;
	}

	*pPath = path.AllocSysString();
	return S_OK;
}