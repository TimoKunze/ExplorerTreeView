// TreeViewItem.cpp: A wrapper for existing treeview items.

#include "stdafx.h"
#include "TreeViewItem.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP TreeViewItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_ITreeViewItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


TreeViewItem::Properties::~Properties()
{
	if(pOwnerExTvw) {
		pOwnerExTvw->Release();
	}
}

HWND TreeViewItem::Properties::GetExTvwHWnd(void)
{
	ATLASSUME(pOwnerExTvw);

	OLE_HANDLE handle = NULL;
	pOwnerExTvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


BOOL TreeViewItem::CouldMoveTo(HTREEITEM hItem, HTREEITEM hNewParentItem, HTREEITEM hNewPreviousItem, HWND hWndTvw/* = NULL*/)
{
	if(!hWndTvw) {
		hWndTvw = properties.GetExTvwHWnd();
	}
	ATLASSERT(IsWindow(hWndTvw));

	if(hNewParentItem == TVI_ROOT) {
		hNewParentItem = NULL;
	}

	if((hNewParentItem == hItem) || (hNewPreviousItem == hItem)) {
		return FALSE;
	}
	if((hNewPreviousItem != NULL) && !ISPREDEFINEDHTREEITEM(hNewPreviousItem)) {
		HTREEITEM h = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hNewPreviousItem)));
		if(h != hNewParentItem) {
			// hNewPreviousItem is no immediate sub-item of hNewParentItem
			return FALSE;
		}
	}
	if(hNewParentItem != NULL) {
		// hNewParentItem mustn't be a sub-item of hItem
		HTREEITEM h = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hNewParentItem)));
		while(h) {
			if(h == hItem) {
				return FALSE;
			}
			h = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(h)));
		}
	}

	return TRUE;
}

BOOL TreeViewItem::CouldMoveTo(HTREEITEM hNewParentItem, HTREEITEM hNewPreviousItem, HWND hWndTvw/* = NULL*/)
{
	return CouldMoveTo(properties.itemHandle, hNewParentItem, hNewPreviousItem, hWndTvw);
}

BOOL TreeViewItem::IsExpanded(HTREEITEM hItem, BOOL allowPartialExpansion/* = TRUE*/, HWND hWndTvw/* = NULL*/)
{
	if(!hWndTvw) {
		hWndTvw = properties.GetExTvwHWnd();
	}
	ATLASSERT(IsWindow(hWndTvw));

	if(!hItem || (hItem == TVI_ROOT)) {
		return TRUE;
	}

	TVITEMEX item = {0};
	item.hItem = hItem;
	item.stateMask = TVIS_EXPANDED | TVIS_EXPANDPARTIAL;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		if(allowPartialExpansion) {
			return (item.state & (TVIS_EXPANDED | TVIS_EXPANDPARTIAL) ? TRUE : FALSE);
		} else {
			return ((item.state & TVIS_EXPANDED) == TVIS_EXPANDED);
		}
	}

	return FALSE;
}

BOOL TreeViewItem::IsSelected(HTREEITEM hItem, HWND hWndTvw/* = NULL*/)
{
	if(!hWndTvw) {
		hWndTvw = properties.GetExTvwHWnd();
	}
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = hItem;
	item.stateMask = TVIS_SELECTED;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return ((item.state & TVIS_SELECTED) == TVIS_SELECTED);
	}

	return FALSE;
}

BOOL TreeViewItem::IsSelected(HWND hWndTvw/* = NULL*/)
{
	return IsSelected(properties.itemHandle, hWndTvw);
}

BOOL TreeViewItem::IsTruncated(HWND hWndTvw/* = NULL*/)
{
	if(!hWndTvw) {
		hWndTvw = properties.GetExTvwHWnd();
	}
	return IsTreeViewItemTruncated(hWndTvw, properties.itemHandle);
}

HRESULT TreeViewItem::Move(HTREEITEM hItem, HTREEITEM hNewParentItem, HTREEITEM hNewPreviousItem, HWND hWndTvw/* = NULL*/, BOOL skipErrorChecking/* = TRUE*/)
{
	if(!hWndTvw) {
		hWndTvw = properties.GetExTvwHWnd();
	}
	ATLASSERT(IsWindow(hWndTvw));

	if(!skipErrorChecking) {
		if(!CouldMoveTo(hItem, hNewParentItem, hNewPreviousItem, hWndTvw)) {
			return E_INVALIDARG;
		}
	}

	// we re-insert this item at the new location and do the same for its children
	TVINSERTSTRUCT item = {0};
	if(SUCCEEDED(SaveState(hItem, &item, hWndTvw))) {
		BOOL selected = ((item.itemex.mask & TVIF_STATE) && ((item.itemex.state & TVIS_SELECTED) == TVIS_SELECTED));

		item.hInsertAfter = hNewPreviousItem;
		item.hParent = hNewParentItem;
		// make item ID management easier
		// NOTE: Should we extend SaveState so that it may save either the ID or the lParam?
		item.itemex.lParam = properties.pOwnerExTvw->ItemHandleToID(hItem);
		item.itemex.mask |= TVIF_NOINTERCEPTION;
		// the selection state can't be copied
		item.itemex.state &= ~TVIS_SELECTED;
		item.itemex.stateMask &= ~TVIS_SELECTED;
		item.itemex.stateMask = TVIS_BOLD | TVIS_CUT | TVIS_DROPHILITED | TVIS_EXPANDED | TVIS_EXPANDEDONCE | TVIS_EXPANDPARTIAL | TVIS_OVERLAYMASK | TVIS_STATEIMAGEMASK;

		properties.pOwnerExTvw->EnterSilentItemInsertionSection();
		HTREEITEM hNewItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&item)));
		// free the caption string
		SECUREFREE(item.itemex.pszText);

		// sync our internal item map
		properties.pOwnerExTvw->ReplaceItemHandle(hItem, hNewItem);

		BOOL caret = (hItem == flags.caretItemChange.hPreviousItem);
		if(caret) {
			// we're gonna removing the current caret, so store the just inserted item as the new caret
			flags.caretItemChange.hNewItem = hNewItem;
			flags.caretItemChange.hPreviousItem = NULL;
		}
		// restore flag
		properties.pOwnerExTvw->LeaveSilentItemInsertionSection();

		// if we're selected, ensure the item is visible
		if(selected && !IsExpanded(hNewParentItem, TRUE, hWndTvw)) {
			SendMessage(hWndTvw, TVM_ENSUREVISIBLE, 0, reinterpret_cast<LPARAM>(hNewItem));
		}

		// move all sub-items
		HRESULT hr = MoveSubItems(hItem, hNewItem, hWndTvw);

		if(SUCCEEDED(hr)) {
			// prepare the item's deletion
			properties.pOwnerExTvw->AddSilentItemDeletion(hItem);
		} else {
			// the sub-items couldn't be moved - try to rollback
			// remove the new item
			properties.pOwnerExTvw->AddSilentItemDeletion(hNewItem);
			SendMessage(hWndTvw, TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(hNewItem));
			// sync our internal item vector
			properties.pOwnerExTvw->ReplaceItemHandle(hNewItem, hItem);
		}

		return hr;
	} else {     // SUCCEEDED(SaveState(hItem, &item, hWndTvw))
		return E_FAIL;
	}
}

HRESULT TreeViewItem::Move(HTREEITEM hNewParentItem, HTREEITEM hNewPreviousItem, HWND hWndTvw/* = NULL*/, BOOL skipErrorChecking/* = TRUE*/)
{
	if(!hWndTvw) {
		hWndTvw = properties.GetExTvwHWnd();
	}
	ATLASSERT(IsWindow(hWndTvw));

	if(!skipErrorChecking) {
		if(!CouldMoveTo(hNewParentItem, hNewPreviousItem, hWndTvw)) {
			return E_INVALIDARG;
		}
	}

	// we re-insert this item at the new location, do the same for its children and remove this item
	TVINSERTSTRUCT item = {0};
	if(SUCCEEDED(SaveState(&item, hWndTvw))) {
		CWindowEx(hWndTvw).InternalSetRedraw(FALSE);
		BOOL selected = ((item.itemex.mask & TVIF_STATE) && ((item.itemex.state & TVIS_SELECTED) == TVIS_SELECTED));

		flags.caretItemChange.hNewItem = NULL;
		flags.caretItemChange.hPreviousItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0));

		item.hInsertAfter = hNewPreviousItem;
		item.hParent = hNewParentItem;
		// make item ID management easier
		// NOTE: Should we extend SaveState so that it may save either the ID or the lParam?
		item.itemex.lParam = properties.pOwnerExTvw->ItemHandleToID(properties.itemHandle);
		item.itemex.mask |= TVIF_NOINTERCEPTION;
		// the selection state can't be copied
		item.itemex.state &= ~TVIS_SELECTED;
		item.itemex.stateMask &= ~TVIS_SELECTED;
		item.itemex.stateMask = TVIS_BOLD | TVIS_CUT | TVIS_DROPHILITED | TVIS_EXPANDED | TVIS_EXPANDEDONCE | TVIS_EXPANDPARTIAL | TVIS_OVERLAYMASK | TVIS_STATEIMAGEMASK;

		properties.pOwnerExTvw->EnterSilentItemInsertionSection();
		HTREEITEM hNewItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&item)));
		// free the caption string
		SECUREFREE(item.itemex.pszText);

		// sync our internal item map
		properties.pOwnerExTvw->ReplaceItemHandle(properties.itemHandle, hNewItem);

		BOOL caret = (properties.itemHandle == flags.caretItemChange.hPreviousItem);
		if(caret) {
			// we're gonna removing the current caret, so store the just inserted item as the new caret
			flags.caretItemChange.hNewItem = hNewItem;
			flags.caretItemChange.hPreviousItem = NULL;
		}
		// restore flag
		properties.pOwnerExTvw->LeaveSilentItemInsertionSection();

		// move all sub-items
		MoveSubItems(properties.itemHandle, hNewItem, hWndTvw);

		// if we're selected, ensure the item is visible
		if(selected && !IsExpanded(hNewParentItem, TRUE, hWndTvw)) {
			SendMessage(hWndTvw, TVM_ENSUREVISIBLE, 0, reinterpret_cast<LPARAM>(hNewItem));
		}

		// select the new caret
		if(flags.caretItemChange.hNewItem) {
			// prepare the control for selecting the new caret
			properties.pOwnerExTvw->AddInternalCaretChange(flags.caretItemChange.hNewItem);
			properties.pOwnerExTvw->EnterNoSingleExpandSection();
			// ensure it's visible
			properties.pOwnerExTvw->AddCaretChangeTask(TRUE, FALSE, new CCEnsureVisibleTask(flags.caretItemChange.hNewItem));
			// finally select the new caret
			SendMessage(hWndTvw, TVM_SELECTITEM, static_cast<WPARAM>(TVGN_CARET), reinterpret_cast<LPARAM>(flags.caretItemChange.hNewItem));
			// clean up
			properties.pOwnerExTvw->LeaveNoSingleExpandSection();
			properties.pOwnerExTvw->RemoveInternalCaretChange(flags.caretItemChange.hNewItem);
		}

		// prepare the item's deletion
		properties.pOwnerExTvw->AddSilentItemDeletion(properties.itemHandle);
		// finally remove the old item
		SendMessage(hWndTvw, TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(properties.itemHandle));
		properties.itemHandle = hNewItem;

		CWindowEx(hWndTvw).InternalSetRedraw(TRUE);
		return S_OK;
	} else {     // SUCCEEDED(SaveState(&item, hWndTvw))
		return E_FAIL;
	}
}

HRESULT TreeViewItem::MoveSubItems(HTREEITEM hParentItem, HTREEITEM hNewParentItem, HWND hWndTvw/* = NULL*/)
{
	if(!hWndTvw) {
		hWndTvw = properties.GetExTvwHWnd();
	}
	ATLASSERT(IsWindow(hWndTvw));

	// iterate each sub-item of hParentItem and duplicate it
	HTREEITEM hSubItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(hParentItem)));
	if(!hSubItem) {
		return S_OK;
	}

	// get the item's proceeding item before moving it
	HTREEITEM hNextItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hSubItem)));
	while(hSubItem) {
		Move(hSubItem, hNewParentItem, TVI_LAST, hWndTvw);

		// get the next item to process
		hSubItem = hNextItem;
		if(hNextItem) {
			hNextItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hNextItem)));
		}
	}

	return S_OK;
}

HRESULT TreeViewItem::SaveState(HTREEITEM hItem, LPTVINSERTSTRUCT pTarget, HWND hWndTvw/* = NULL*/)
{
	if(!hWndTvw) {
		hWndTvw = properties.GetExTvwHWnd();
	}
	ATLASSERT(IsWindow(hWndTvw));
	ATLASSERT(hItem != NULL);
	if(!hItem) {
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pTarget, TVINSERTSTRUCT);
	if(!pTarget) {
		return E_POINTER;
	}

	pTarget->hParent = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hItem)));
	pTarget->hInsertAfter = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PREVIOUS), reinterpret_cast<LPARAM>(hItem)));

	ZeroMemory(&pTarget->itemex, sizeof(TVITEMEX));
	pTarget->itemex.hItem = hItem;
	pTarget->itemex.mask = TVIF_CHILDREN | TVIF_HANDLE | TVIF_IMAGE | TVIF_INTEGRAL | TVIF_PARAM | TVIF_SELECTEDIMAGE | TVIF_STATE;
	if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
		pTarget->itemex.mask |= TVIF_STATEEX | TVIF_EXPANDEDIMAGE;
	}
	/*TODO: if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
		pTarget->itemex.mask |= 0x0400;
	}*/

	LRESULT lr = SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&pTarget->itemex));
	if(lr) {
		// TODO: for some reason we can't get the text using pTarget->itemex - spooky
		pTarget->itemex.cchTextMax = MAX_ITEMTEXTLENGTH;
		pTarget->itemex.pszText = reinterpret_cast<LPTSTR>(malloc((pTarget->itemex.cchTextMax + 1) * sizeof(TCHAR)));
		pTarget->itemex.mask |= TVIF_TEXT;
		TVITEMEX tmp = pTarget->itemex;
		tmp.mask = TVIF_HANDLE | TVIF_TEXT;

		lr = SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&tmp));
	}

	return static_cast<HRESULT>(lr);
}


void TreeViewItem::Attach(HTREEITEM hItem)
{
	properties.itemHandle = hItem;
}

void TreeViewItem::Detach(void)
{
	properties.itemHandle = NULL;
}

HRESULT TreeViewItem::LoadState(LPTVINSERTSTRUCT pSource)
{
	ATLASSERT_POINTER(pSource, TVINSERTSTRUCT);
	if(!pSource) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	HTREEITEM hParentItem = pSource->hParent;
	if(hParentItem == TVI_ROOT) {
		hParentItem = NULL;
	}

	// maybe the item must be moved
	HTREEITEM hCurParentItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(properties.itemHandle)));
	HTREEITEM hCurPreviousItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PREVIOUS), reinterpret_cast<LPARAM>(properties.itemHandle)));
	if((hCurParentItem != hParentItem) || (hCurPreviousItem != pSource->hInsertAfter)) {
		Move(hParentItem, pSource->hInsertAfter, hWndTvw);
	}

	if(pSource->itemex.mask & TVIF_STATE) {
		TVITEMEX item = {0};
		item.hItem = properties.itemHandle;
		if(pSource->itemex.state & TVIS_DROPHILITED) {
			item.state = TVIS_DROPHILITED;
		}
		item.stateMask = TVIS_DROPHILITED;
		item.mask = TVIF_HANDLE | TVIF_STATE;
		SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
	}

	// now apply the settings
	VARIANT_BOOL b = VARIANT_FALSE;
	if(pSource->itemex.mask & TVIF_STATE) {
		b = BOOL2VARIANTBOOL(pSource->itemex.state & TVIS_BOLD);
		put_Bold(b);
	}

	b = VARIANT_FALSE;
	if(pSource->itemex.mask & TVIF_STATE) {
		b = BOOL2VARIANTBOOL(pSource->itemex.state & TVIS_CUT);
		put_Ghosted(b);
	}
	if(pSource->itemex.mask & TVIF_STATEEX) {
		b = BOOL2VARIANTBOOL(pSource->itemex.uStateEx & TVIS_EX_DISABLED);
		put_Grayed(b);
	}

	BSTR text = NULL;
	if(pSource->itemex.mask & TVIF_TEXT) {
		if(pSource->itemex.pszText != LPSTR_TEXTCALLBACK) {
			text = _bstr_t(pSource->itemex.pszText).Detach();
		}
		put_Text(text);
	}
	SysFreeString(text);

	ExpansionStateConstants es = esNeverWasExpanded;
	if(pSource->itemex.mask & TVIF_STATE) {
		es = static_cast<ExpansionStateConstants>(pSource->itemex.state & (TVIS_EXPANDED | TVIS_EXPANDEDONCE | TVIS_EXPANDPARTIAL));
		put_ExpansionState(es);
	}

	HasExpandoConstants hasExpando = heNo;
	if(pSource->itemex.mask & TVIF_CHILDREN) {
		hasExpando = static_cast<HasExpandoConstants>(pSource->itemex.cChildren);
		put_HasExpando(hasExpando);
	}

	LONG l = 0;
	if(pSource->itemex.mask & TVIF_EXPANDEDIMAGE) {
		l = pSource->itemex.iExpandedImage;
		put_ExpandedIconIndex(l);
	}
	l = 1;
	if(pSource->itemex.mask & TVIF_INTEGRAL) {
		l = pSource->itemex.iIntegral;
		put_HeightIncrement(l);
	}
	l = 0;
	if(pSource->itemex.mask & TVIF_IMAGE) {
		l = pSource->itemex.iImage;
		put_IconIndex(l);
	}
	l = 0;
	if(pSource->itemex.mask & TVIF_PARAM) {
		l = static_cast<LONG>(pSource->itemex.lParam);
		put_ItemData(l);
	}
	l = 0;
	if(pSource->itemex.mask & TVIF_STATE) {
		l = ((pSource->itemex.state & TVIS_OVERLAYMASK) >> 8);
		put_OverlayIndex(l);
	}
	/*TODO: l = 0;
	if(pSource->itemex.mask & 0x0400) {
		l = pSource->itemex.iPadding;
		put_Padding(l);
	}*/
	l = 0;
	if(pSource->itemex.mask & TVIF_SELECTEDIMAGE) {
		l = pSource->itemex.iSelectedImage;
		put_SelectedIconIndex(l);
	}
	l = 1;
	if(pSource->itemex.mask & TVIF_STATE) {
		l = ((pSource->itemex.state & TVIS_STATEIMAGEMASK) >> 12);
		put_StateImageIndex(l);
	}

	return S_OK;
}

HRESULT TreeViewItem::LoadState(VirtualTreeViewItem* pSource)
{
	ATLASSUME(pSource);
	if(!pSource) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	HTREEITEM hParentItem = pSource->GetParentItem();
	if(hParentItem == TVI_ROOT) {
		hParentItem = NULL;
	}

	HTREEITEM hPrevItem = pSource->GetPreviousItem();
	// maybe the item must be moved
	HTREEITEM hCurParentItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(properties.itemHandle)));
	HTREEITEM hCurPreviousItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PREVIOUS), reinterpret_cast<LPARAM>(properties.itemHandle)));
	if((hCurParentItem != hParentItem) || (hCurPreviousItem != hPrevItem)) {
		Move(hParentItem, hPrevItem, hWndTvw);
	}

	VARIANT_BOOL b = VARIANT_FALSE;
	pSource->get_DropHilited(&b);
	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	if(b != VARIANT_FALSE) {
		item.state = TVIS_DROPHILITED;
	}
	item.stateMask = TVIS_DROPHILITED;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));

	// now apply the settings
	b = VARIANT_FALSE;
	pSource->get_Bold(&b);
	put_Bold(b);
	b = VARIANT_FALSE;
	pSource->get_Ghosted(&b);
	put_Ghosted(b);
	b = VARIANT_TRUE;
	pSource->get_Grayed(&b);
	put_Grayed(b);

	BSTR text = NULL;
	pSource->get_Text(&text);
	put_Text(text);
	SysFreeString(text);

	ExpansionStateConstants es = esNeverWasExpanded;
	pSource->get_ExpansionState(&es);
	put_ExpansionState(es);

	HasExpandoConstants hasExpando = heNo;
	pSource->get_HasExpando(&hasExpando);
	put_HasExpando(hasExpando);

	LONG l = 0;
	pSource->get_ExpandedIconIndex(&l);
	put_ExpandedIconIndex(l);
	l = 1;
	pSource->get_HeightIncrement(&l);
	put_HeightIncrement(l);
	l = 0;
	pSource->get_IconIndex(&l);
	put_IconIndex(l);
	l = 0;
	pSource->get_ItemData(&l);
	put_ItemData(l);
	l = 0;
	pSource->get_OverlayIndex(&l);
	put_OverlayIndex(l);
	l = 0;
	pSource->get_SelectedIconIndex(&l);
	put_SelectedIconIndex(l);
	l = 1;
	pSource->get_StateImageIndex(&l);
	put_StateImageIndex(l);

	return S_OK;
}

HRESULT TreeViewItem::SaveState(LPTVINSERTSTRUCT pTarget, HWND hWndTvw/* = NULL*/)
{
	return SaveState(properties.itemHandle, pTarget, hWndTvw);
}

HRESULT TreeViewItem::SaveState(VirtualTreeViewItem* pTarget)
{
	ATLASSUME(pTarget);
	if(!pTarget) {
		return E_POINTER;
	}

	pTarget->SetOwner(properties.pOwnerExTvw);
	TVINSERTSTRUCT item = {0};
	HRESULT hr = SaveState(&item);
	pTarget->LoadState(&item);
	SECUREFREE(item.itemex.pszText);

	return hr;
}


void TreeViewItem::SetOwner(ExplorerTreeView* pOwner)
{
	if(properties.pOwnerExTvw) {
		properties.pOwnerExTvw->Release();
	}
	properties.pOwnerExTvw = pOwner;
	if(properties.pOwnerExTvw) {
		properties.pOwnerExTvw->AddRef();
	}
}


STDMETHODIMP TreeViewItem::get_AccessibilityID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	if(RunTimeHelper::IsCommCtrl6()) {
		*pValue = static_cast<UINT>(SendMessage(hWndTvw, TVM_MAPHTREEITEMTOACCID, reinterpret_cast<WPARAM>(properties.itemHandle), 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_Bold(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.stateMask = TVIS_BOLD;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = BOOL2VARIANTBOOL(item.state & TVIS_BOLD);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::put_Bold(VARIANT_BOOL newValue)
{
	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	if(newValue != VARIANT_FALSE) {
		item.state = TVIS_BOLD;
	}
	item.stateMask = TVIS_BOLD;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_Caret(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	*pValue = BOOL2VARIANTBOOL(properties.itemHandle == reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0)));
	return S_OK;
}

STDMETHODIMP TreeViewItem::get_DropHilited(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.stateMask = TVIS_DROPHILITED;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = BOOL2VARIANTBOOL(item.state & TVIS_DROPHILITED);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_ExpandedIconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;

	if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
		HWND hWndTvw = properties.GetExTvwHWnd();
		ATLASSERT(IsWindow(hWndTvw));

		TVITEMEX item = {0};
		item.hItem = properties.itemHandle;
		item.mask = TVIF_HANDLE | TVIF_EXPANDEDIMAGE;
		if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			*pValue = item.iExpandedImage;
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP TreeViewItem::put_ExpandedIconIndex(LONG newValue)
{
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
		HWND hWndTvw = properties.GetExTvwHWnd();
		ATLASSERT(IsWindow(hWndTvw));

		TVITEMEX item = {0};
		item.hItem = properties.itemHandle;
		item.iExpandedImage = newValue;
		item.mask = TVIF_HANDLE | TVIF_EXPANDEDIMAGE;
		if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_ExpansionState(ExpansionStateConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ExpansionStateConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = static_cast<ExpansionStateConstants>(item.state & (TVIS_EXPANDED | TVIS_EXPANDEDONCE | TVIS_EXPANDPARTIAL));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::put_ExpansionState(ExpansionStateConstants newValue)
{
	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.state = newValue;
	item.stateMask = TVIS_EXPANDED | TVIS_EXPANDEDONCE | TVIS_EXPANDPARTIAL;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		InvalidateRect(hWndTvw, NULL, TRUE);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_FirstSubItem(ITreeViewItem** ppFirstSubItem)
{
	ATLASSERT_POINTER(ppFirstSubItem, ITreeViewItem*);
	if(!ppFirstSubItem) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	HTREEITEM hChild = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(properties.itemHandle)));
	ClassFactory::InitTreeItem(hChild, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppFirstSubItem));
	return S_OK;
}

STDMETHODIMP TreeViewItem::get_Ghosted(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.stateMask = TVIS_CUT;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = BOOL2VARIANTBOOL(item.state & TVIS_CUT);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::put_Ghosted(VARIANT_BOOL newValue)
{
	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	if(newValue != VARIANT_FALSE) {
		item.state = TVIS_CUT;
	}
	item.stateMask = TVIS_CUT;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_Grayed(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;

	if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
		HWND hWndTvw = properties.GetExTvwHWnd();
		ATLASSERT(IsWindow(hWndTvw));

		TVITEMEX item = {0};
		item.hItem = properties.itemHandle;
		item.mask = TVIF_HANDLE | TVIF_STATEEX;
		if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			*pValue = BOOL2VARIANTBOOL(item.uStateEx & TVIS_EX_DISABLED);
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP TreeViewItem::put_Grayed(VARIANT_BOOL newValue)
{
	if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
		HWND hWndTvw = properties.GetExTvwHWnd();
		ATLASSERT(IsWindow(hWndTvw));

		TVITEMEX item = {0};
		item.hItem = properties.itemHandle;
		item.mask = TVIF_HANDLE | TVIF_STATEEX;
		if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			if(newValue != VARIANT_FALSE) {
				item.uStateEx |= TVIS_EX_DISABLED;
			} else {
				item.uStateEx &= ~TVIS_EX_DISABLED;
			}
			if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
				return S_OK;
			}
		}
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_Handle(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(properties.itemHandle);
	return S_OK;
}

STDMETHODIMP TreeViewItem::get_HasExpando(HasExpandoConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HasExpandoConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.mask = TVIF_CHILDREN | TVIF_HANDLE;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = static_cast<HasExpandoConstants>(item.cChildren);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::put_HasExpando(HasExpandoConstants newValue)
{
	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.cChildren = newValue;
	item.mask = TVIF_CHILDREN | TVIF_HANDLE;
	if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_HeightIncrement(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.mask = TVIF_HANDLE | TVIF_INTEGRAL;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iIntegral;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::put_HeightIncrement(LONG newValue)
{
	if(newValue < 1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.iIntegral = newValue;
	item.mask = TVIF_HANDLE | TVIF_INTEGRAL;
	if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.mask = TVIF_HANDLE | TVIF_IMAGE;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iImage;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::put_IconIndex(LONG newValue)
{
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.iImage = newValue;
	item.mask = TVIF_HANDLE | TVIF_IMAGE;
	if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExTvw) {
		*pValue = properties.pOwnerExTvw->ItemHandleToID(properties.itemHandle);
		if(*pValue) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_ItemData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.mask = TVIF_HANDLE | TVIF_PARAM;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = static_cast<LONG>(item.lParam);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::put_ItemData(LONG newValue)
{
	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.lParam = newValue;
	item.mask = TVIF_HANDLE | TVIF_PARAM;
	if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_LastSubItem(ITreeViewItem** ppLastSubItem)
{
	ATLASSERT_POINTER(ppLastSubItem, ITreeViewItem*);
	if(!ppLastSubItem) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	HTREEITEM hChild = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(properties.itemHandle)));
	HTREEITEM hDummy = hChild;
	while(hDummy) {
		hChild = hDummy;
		hDummy = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hChild)));
	}
	ClassFactory::InitTreeItem(hChild, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppLastSubItem));
	return S_OK;
}

STDMETHODIMP TreeViewItem::get_Level(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	*pValue = 0;
	HTREEITEM hItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(properties.itemHandle)));
	while(hItem) {
		++(*pValue);
		hItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hItem)));
	}
	return S_OK;
}

STDMETHODIMP TreeViewItem::get_NextSiblingItem(ITreeViewItem** ppNextItem)
{
	ATLASSERT_POINTER(ppNextItem, ITreeViewItem*);
	if(!ppNextItem) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	HTREEITEM hNext = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(properties.itemHandle)));
	ClassFactory::InitTreeItem(hNext, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppNextItem));
	return S_OK;
}

STDMETHODIMP TreeViewItem::get_NextVisibleItem(ITreeViewItem** ppNextItem)
{
	ATLASSERT_POINTER(ppNextItem, ITreeViewItem*);
	if(!ppNextItem) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	HTREEITEM hNext = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXTVISIBLE), reinterpret_cast<LPARAM>(properties.itemHandle)));
	ClassFactory::InitTreeItem(hNext, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppNextItem));
	return S_OK;
}

STDMETHODIMP TreeViewItem::get_OverlayIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = (item.state & TVIS_OVERLAYMASK) >> 8;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::put_OverlayIndex(LONG newValue)
{
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.state = INDEXTOOVERLAYMASK(newValue);
	item.stateMask = TVIS_OVERLAYMASK;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return S_OK;
	}
	return E_FAIL;
}

/*TODO: STDMETHODIMP TreeViewItem::get_Padding(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;

	if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
		HWND hWndTvw = properties.GetExTvwHWnd();
		ATLASSERT(IsWindow(hWndTvw));

		TVITEMEX item = {0};
		item.hItem = properties.itemHandle;
		item.mask = TVIF_HANDLE | 0x0400;
		if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			*pValue = item.iPadding;
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP TreeViewItem::put_Padding(LONG newValue)
{
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		//return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
		HWND hWndTvw = properties.GetExTvwHWnd();
		ATLASSERT(IsWindow(hWndTvw));

		TVITEMEX item = {0};
		item.hItem = properties.itemHandle;
		item.iPadding = newValue;
		item.mask = TVIF_HANDLE | 0x0400;
		if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
	}
	return E_FAIL;
}*/

STDMETHODIMP TreeViewItem::get_ParentItem(ITreeViewItem** ppParentItem)
{
	ATLASSERT_POINTER(ppParentItem, ITreeViewItem*);
	if(!ppParentItem) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	HTREEITEM hParent = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(properties.itemHandle)));
	ClassFactory::InitTreeItem(hParent, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppParentItem));
	return S_OK;
}

STDMETHODIMP TreeViewItem::get_PreviousSiblingItem(ITreeViewItem** ppPreviousItem)
{
	ATLASSERT_POINTER(ppPreviousItem, ITreeViewItem*);
	if(!ppPreviousItem) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	HTREEITEM hPrev = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PREVIOUS), reinterpret_cast<LPARAM>(properties.itemHandle)));
	ClassFactory::InitTreeItem(hPrev, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppPreviousItem));
	return S_OK;
}

STDMETHODIMP TreeViewItem::putref_PreviousSiblingItem(ITreeViewItem* pNewPreviousItem)
{
	ITreeViewItem* pParentItem = NULL;
	if(SUCCEEDED(get_ParentItem(&pParentItem))) {
		if(pNewPreviousItem) {
			pNewPreviousItem->AddRef();
		}
		HRESULT hr = Move(pParentItem, _variant_t(static_cast<LPDISPATCH>(pNewPreviousItem), FALSE));
		// pNewPreviousItem seems to be auto-released when the _variant_t is destroyed
		if(pParentItem) {
			pParentItem->Release();
		}
		return hr;
	} else {
		return E_FAIL;
	}
}

STDMETHODIMP TreeViewItem::get_PreviousVisibleItem(ITreeViewItem** ppPreviousItem)
{
	ATLASSERT_POINTER(ppPreviousItem, ITreeViewItem*);
	if(!ppPreviousItem) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	HTREEITEM hPrev = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PREVIOUSVISIBLE), reinterpret_cast<LPARAM>(properties.itemHandle)));
	ClassFactory::InitTreeItem(hPrev, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppPreviousItem));
	return S_OK;
}

STDMETHODIMP TreeViewItem::get_Selected(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.stateMask = TVIS_SELECTED;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = BOOL2VARIANTBOOL(item.state & TVIS_SELECTED);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_SelectedIconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.mask = TVIF_HANDLE | TVIF_SELECTEDIMAGE;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iSelectedImage;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::put_SelectedIconIndex(LONG newValue)
{
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.iSelectedImage = newValue;
	item.mask = TVIF_HANDLE | TVIF_SELECTEDIMAGE;
	if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return S_OK;
	}
	return E_FAIL;
}

#ifdef INCLUDESHELLBROWSERINTERFACE
	STDMETHODIMP TreeViewItem::get_ShellTreeViewItemObject(IDispatch** ppItem)
	{
		ATLASSERT_POINTER(ppItem, LPDISPATCH);
		if(!ppItem) {
			return E_POINTER;
		}

		if(!properties.itemHandle) {
			return E_FAIL;
		}

		if(properties.pOwnerExTvw) {
			if(properties.pOwnerExTvw->shellBrowserInterface.pInternalMessageListener) {
				properties.pOwnerExTvw->shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_GETSHTREEVIEWITEM, reinterpret_cast<WPARAM>(properties.itemHandle), reinterpret_cast<LPARAM>(ppItem));
				return S_OK;
			}
		}
		return E_FAIL;
	}
#endif

STDMETHODIMP TreeViewItem::get_StateImageIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = (item.state & TVIS_STATEIMAGEMASK) >> 12;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::put_StateImageIndex(LONG newValue)
{
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.state = INDEXTOSTATEIMAGEMASK(newValue);
	item.stateMask = TVIS_STATEIMAGEMASK;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::get_SubItems(ITreeViewItems** ppSubItems)
{
	ATLASSERT_POINTER(ppSubItems, ITreeViewItems*);
	if(!ppSubItems) {
		return E_POINTER;
	}

	ClassFactory::InitTreeItems(properties.pOwnerExTvw, IID_ITreeViewItems, reinterpret_cast<LPUNKNOWN*>(ppSubItems), TRUE, properties.itemHandle);
	return S_OK;
}

STDMETHODIMP TreeViewItem::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	LPWSTR pBuffer = reinterpret_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, (MAX_ITEMTEXTLENGTH + 1) * sizeof(WCHAR)));
	if(!pBuffer) {
		return E_OUTOFMEMORY;
	}

	HRESULT hr = E_FAIL;

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.cchTextMax = MAX_ITEMTEXTLENGTH;
	item.pszText = reinterpret_cast<LPTSTR>(pBuffer);
	item.mask = TVIF_HANDLE | TVIF_TEXT;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = _bstr_t(item.pszText).Detach();
		hr = S_OK;
	}
	HeapFree(GetProcessHeap(), 0, pBuffer);
	return hr;
}

STDMETHODIMP TreeViewItem::put_Text(BSTR newValue)
{
	HRESULT hr = E_FAIL;

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVITEMEX item = {0};
	item.mask = TVIF_HANDLE | TVIF_TEXT;
	item.hItem = properties.itemHandle;
	COLE2T converter(newValue);
	if(newValue) {
		item.pszText = converter;
	} else {
		item.pszText = LPSTR_TEXTCALLBACK;
	}
	if(SendMessage(hWndTvw, TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP TreeViewItem::get_Virtual(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;

	if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
		HWND hWndTvw = properties.GetExTvwHWnd();
		ATLASSERT(IsWindow(hWndTvw));

		TVITEMEX item = {0};
		item.hItem = properties.itemHandle;
		item.mask = TVIF_HANDLE | TVIF_STATEEX;
		if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			*pValue = BOOL2VARIANTBOOL(item.uStateEx & TVIS_EX_FLAT);
			hr = S_OK;
		}
	}
	return hr;
}


STDMETHODIMP TreeViewItem::Collapse(VARIANT_BOOL removeSubItems/* = VARIANT_FALSE*/)
{
	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	if(SUCCEEDED(SendMessage(hWndTvw, TVM_EXPAND, (removeSubItems == VARIANT_FALSE ? TVE_COLLAPSE : TVE_COLLAPSE | TVE_COLLAPSERESET), reinterpret_cast<LPARAM>(properties.itemHandle)))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::CouldMoveTo(ITreeViewItem* pNewParentItem, VARIANT insertAfter, VARIANT_BOOL* pCouldMove)
{
	ATLASSERT_POINTER(pCouldMove, VARIANT_BOOL);
	if(!pCouldMove) {
		return E_POINTER;
	}

	HTREEITEM hNewParentItem = TVI_ROOT;
  HTREEITEM hNewPreviousSiblingItem = TVI_LAST;

	if(pNewParentItem) {
		OLE_HANDLE h = NULL;
		pNewParentItem->get_Handle(&h);
		hNewParentItem = static_cast<HTREEITEM>(LongToHandle(h));
	}

	switch(insertAfter.vt) {
		case VT_DISPATCH: {
			CComQIPtr<ITreeViewItem, &IID_ITreeViewItem> pTvwItem(insertAfter.pdispVal);
			if(pTvwItem) {
				OLE_HANDLE h = NULL;
				pTvwItem->get_Handle(&h);
				hNewPreviousSiblingItem = static_cast<HTREEITEM>(LongToHandle(h));
			} else {
				// VB's 'Nothing'
				hNewPreviousSiblingItem = TVI_FIRST;
			}
			break;
		}
		case VT_UNKNOWN: {
			CComQIPtr<ITreeViewItem, &IID_ITreeViewItem> pTvwItem(insertAfter.punkVal);
			if(pTvwItem) {
				OLE_HANDLE h = NULL;
				pTvwItem->get_Handle(&h);
				hNewPreviousSiblingItem = static_cast<HTREEITEM>(LongToHandle(h));
			} else {
				hNewPreviousSiblingItem = TVI_LAST;
			}
			break;
		}
		case VT_I4:
			hNewPreviousSiblingItem = static_cast<HTREEITEM>(LongToHandle(insertAfter.lVal));
			break;
		case VT_UI4:
			hNewPreviousSiblingItem = static_cast<HTREEITEM>(LongToHandle(insertAfter.ulVal));
			break;
		case VT_EMPTY:
			hNewPreviousSiblingItem = TVI_LAST;
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}

	*pCouldMove = BOOL2VARIANTBOOL(CouldMoveTo(hNewParentItem, hNewPreviousSiblingItem, static_cast<HWND>(NULL)));
	return S_OK;
}

STDMETHODIMP TreeViewItem::CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft/* = NULL*/, OLE_YPOS_PIXELS* pYUpperLeft/* = NULL*/, OLE_HANDLE* phImageList/* = NULL*/)
{
	ATLASSERT_POINTER(phImageList, OLE_HANDLE);
	if(!phImageList) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerExTvw);

	// ask the control for a drag image
	POINT upperLeftPoint = {0};
	*phImageList = HandleToLong(properties.pOwnerExTvw->CreateLegacyDragImage(properties.itemHandle, &upperLeftPoint, NULL));

	if(*phImageList) {
		if(pXUpperLeft) {
			*pXUpperLeft = upperLeftPoint.x;
		}
		if(pYUpperLeft) {
			*pYUpperLeft = upperLeftPoint.y;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::DisplayInfoTip(void)
{
	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
		if(SUCCEEDED(SendMessage(hWndTvw, TVM_SHOWINFOTIP, 0, reinterpret_cast<LPARAM>(properties.itemHandle)))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::EnsureVisible(VARIANT_BOOL* pExpandedItems)
{
	ATLASSERT_POINTER(pExpandedItems, VARIANT_BOOL);
	if(!pExpandedItems) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	LRESULT lr = SendMessage(hWndTvw, TVM_ENSUREVISIBLE, 0, reinterpret_cast<LPARAM>(properties.itemHandle));
	*pExpandedItems = BOOL2VARIANTBOOL(lr == 0);
	return S_OK;
}

STDMETHODIMP TreeViewItem::Expand(VARIANT_BOOL expandPartial/* = VARIANT_FALSE*/)
{
	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	if(SUCCEEDED(SendMessage(hWndTvw, TVM_EXPAND, (expandPartial == VARIANT_FALSE ? TVE_EXPAND : TVE_EXPAND | TVE_EXPANDPARTIAL), reinterpret_cast<LPARAM>(properties.itemHandle)))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItem::GetPath(BSTR* pSeparator, BSTR* pPath)
{
	ATLASSERT_POINTER(pPath, BSTR);
	if(!pPath) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	CAtlString path = TEXT("");
	CAtlString separator = TEXT("\\");
	if(pSeparator) {
		separator = *pSeparator;
	}

	// retrieve this item's text
	TVITEMEX item = {0};
	item.hItem = properties.itemHandle;
	item.cchTextMax = MAX_ITEMTEXTLENGTH;
	item.pszText = reinterpret_cast<LPTSTR>(malloc((item.cchTextMax + 1) * sizeof(TCHAR)));
	item.mask = TVIF_HANDLE | TVIF_TEXT;
	if(SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		path = item.pszText;

		// now go upwardly and complete the path with each level
		item.hItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(properties.itemHandle)));
		if(item.hItem) {
			LRESULT lr = SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item));
			if(lr) {
				path = CAtlString(item.pszText) + separator + path;
			}
			while((item.hItem != NULL) && SUCCEEDED(lr)) {
				item.hItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(item.hItem)));
				if(item.hItem) {
					lr = SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item));
					if(lr) {
						path = CAtlString(item.pszText) + separator + path;
					}
				}
			}
		}
	}
	SECUREFREE(item.pszText);

	*pPath = path.AllocSysString();
	return S_OK;
}

STDMETHODIMP TreeViewItem::GetRectangle(ItemRectangleTypeConstants rectangleType, OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	HRESULT hr = E_FAIL;
	RECT rc = {0};
	if(rectangleType < irtItemExpando) {
		*reinterpret_cast<HTREEITEM*>(&rc) = properties.itemHandle;
		if(SendMessage(hWndTvw, TVM_GETITEMRECT, (rectangleType == irtItem), reinterpret_cast<LPARAM>(&rc))) {
			hr = S_OK;
		}
	} else if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
		TVGETITEMPARTRECTINFO itemPartInfo = {0};
		itemPartInfo.hti = properties.itemHandle;
		itemPartInfo.partID = static_cast<TVITEMPART>(rectangleType - 1);
		itemPartInfo.prc = &rc;
		SendMessage(hWndTvw, TVM_GETITEMPARTRECT, 0, reinterpret_cast<LPARAM>(&itemPartInfo));
		hr = S_OK;
	}
	if(SUCCEEDED(hr)) {
		if(pXLeft) {
			*pXLeft = rc.left;
		}
		if(pYTop) {
			*pYTop = rc.top;
		}
		if(pXRight) {
			*pXRight = rc.right;
		}
		if(pYBottom) {
			*pYBottom = rc.bottom;
		}
	}
	return hr;
}

STDMETHODIMP TreeViewItem::IsTruncated(VARIANT_BOOL* pTruncated)
{
	ATLASSERT_POINTER(pTruncated, VARIANT_BOOL);
	if(!pTruncated) {
		return E_POINTER;
	}

	*pTruncated = BOOL2VARIANTBOOL(IsTruncated());
	return S_OK;
}

STDMETHODIMP TreeViewItem::Move(ITreeViewItem* pNewParentItem, VARIANT insertAfter)
{
	HTREEITEM hNewParentItem = TVI_ROOT;
  HTREEITEM hNewPreviousSiblingItem = TVI_LAST;

	if(pNewParentItem) {
		OLE_HANDLE h = NULL;
		pNewParentItem->get_Handle(&h);
		hNewParentItem = static_cast<HTREEITEM>(LongToHandle(h));
	}

	switch(insertAfter.vt) {
		case VT_DISPATCH: {
			CComQIPtr<ITreeViewItem, &IID_ITreeViewItem> pTvwItem(insertAfter.pdispVal);
			if(pTvwItem) {
				OLE_HANDLE h = NULL;
				pTvwItem->get_Handle(&h);
				hNewPreviousSiblingItem = static_cast<HTREEITEM>(LongToHandle(h));
			} else {
				// VB's 'Nothing'
				hNewPreviousSiblingItem = TVI_FIRST;
			}
			break;
		}
		case VT_UNKNOWN: {
			CComQIPtr<ITreeViewItem, &IID_ITreeViewItem> pTvwItem(insertAfter.punkVal);
			if(pTvwItem) {
				OLE_HANDLE h = NULL;
				pTvwItem->get_Handle(&h);
				hNewPreviousSiblingItem = static_cast<HTREEITEM>(LongToHandle(h));
			} else {
				hNewPreviousSiblingItem = TVI_LAST;
			}
			break;
		}
		case VT_I4:
			hNewPreviousSiblingItem = static_cast<HTREEITEM>(LongToHandle(insertAfter.lVal));
			break;
		case VT_UI4:
			hNewPreviousSiblingItem = static_cast<HTREEITEM>(LongToHandle(insertAfter.ulVal));
			break;
		case VT_EMPTY:
			hNewPreviousSiblingItem = TVI_LAST;
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}

	return Move(hNewParentItem, hNewPreviousSiblingItem, NULL, FALSE);
}

STDMETHODIMP TreeViewItem::SortSubItems(SortByConstants firstCriterion/* = sobShell*/, SortByConstants secondCriterion/* = sobText*/, SortByConstants thirdCriterion/* = sobNone*/, SortByConstants fourthCriterion/* = sobNone*/, SortByConstants fifthCriterion/* = sobNone*/, VARIANT_BOOL recurse/* = VARIANT_FALSE*/, VARIANT_BOOL caseSensitive/* = VARIANT_FALSE*/)
{
	if(properties.pOwnerExTvw) {
		return properties.pOwnerExTvw->SortSubItems(properties.itemHandle, firstCriterion, secondCriterion, thirdCriterion, fourthCriterion, fifthCriterion, VARIANTBOOL2BOOL(recurse), VARIANTBOOL2BOOL(caseSensitive));
	}

	return E_FAIL;
}

STDMETHODIMP TreeViewItem::StartLabelEditing(void)
{
	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	if(SendMessage(hWndTvw, TVM_EDITLABEL, 0, reinterpret_cast<LPARAM>(properties.itemHandle))) {
		return S_OK;
	}
	return E_FAIL;
}