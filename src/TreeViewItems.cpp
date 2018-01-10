// TreeViewItems.cpp: Manages a collection of TreeViewItem objects

#include "stdafx.h"
#include "TreeViewItems.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP TreeViewItems::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_ITreeViewItems, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP TreeViewItems::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<TreeViewItems>* pTvwItemsObj = NULL;
	CComObject<TreeViewItems>::CreateInstance(&pTvwItemsObj);
	pTvwItemsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pTvwItemsObj->properties);

	pTvwItemsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pTvwItemsObj->Release();
	return S_OK;
}

STDMETHODIMP TreeViewItems::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	if((properties.filterType[fpSelected] != ftDeactivated) && !properties.comparisonFunction[fpSelected]) {
		BOOL singleSelectionMode = TRUE;
		if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
			singleSelectionMode = ((SendMessage(hWndTvw, TVM_GETEXTENDEDSTYLE, 0, 0) & TVS_EX_MULTISELECT) == 0);
		}

		if(singleSelectionMode) {
			/* If the filter says either "selected items" only or "no unselected items" only, the caret item is
			   the only one in the collection. */
			LONG lBound = 0;
			SafeArrayGetLBound(properties.filter[fpSelected].parray, 1, &lBound);
			VARIANT element;
			SafeArrayGetElement(properties.filter[fpSelected].parray, &lBound, &element);
			BOOL requiredSelectionState = VARIANTBOOL2BOOL(element.boolVal);
			VariantClear(&element);

			if((requiredSelectionState && (properties.filterType[fpSelected] == ftIncluding)) || (!requiredSelectionState && (properties.filterType[fpSelected] == ftExcluding))) {
				ULONG numberOfItemsReturned = 0;
				// the caret item is the only one matching this filter - apply the other filters
				HTREEITEM hCaretItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0));
				if(IsPartOfCollection(hCaretItem, TRUE, hWndTvw)) {
					if(numberOfMaxItems >= 1) {
						VariantInit(&pItems[0]);
						ClassFactory::InitTreeItem(hCaretItem, properties.pOwnerExTvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[0].pdispVal));
						pItems[0].vt = VT_DISPATCH;
						numberOfItemsReturned = 1;
					}
				}
				if(pNumberOfItemsReturned) {
					*pNumberOfItemsReturned = numberOfItemsReturned;
				}

				return (numberOfItemsReturned == 0 ? S_FALSE : S_OK);
			}
		}
	}

	// check each item in the treeview
	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		do {
			if(properties.hLastEnumeratedItem) {
				properties.hLastEnumeratedItem = GetNextItemToProcess(properties.hLastEnumeratedItem, hWndTvw);
			} else {
				properties.hLastEnumeratedItem = GetFirstItemToProcess(hWndTvw);
			}
		} while((properties.hLastEnumeratedItem != NULL) && (!IsPartOfCollection(properties.hLastEnumeratedItem, FALSE, hWndTvw)));

		if(properties.hLastEnumeratedItem) {
			ClassFactory::InitTreeItem(properties.hLastEnumeratedItem, properties.pOwnerExTvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal));
			pItems[i].vt = VT_DISPATCH;
		} else {
			// there's nothing more to iterate
			break;
		}
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP TreeViewItems::Reset(void)
{
	properties.hLastEnumeratedItem = NULL;
	return S_OK;
}

STDMETHODIMP TreeViewItems::Skip(ULONG numberOfItemsToSkip)
{
	VARIANT dummy;
	ULONG numItemsReturned = 0;
	// we could skip all items at once, but it's easier to skip them one after the other
	for(ULONG i = 1; i <= numberOfItemsToSkip; ++i) {
		HRESULT hr = Next(1, &dummy, &numItemsReturned);
		VariantClear(&dummy);
		if(hr != S_OK || numItemsReturned == 0) {
			// there're no more items to skip, so don't try it anymore
			break;
		}
	}
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


BOOL TreeViewItems::IsValidBooleanFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = (element.vt == VT_BOOL);
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL TreeViewItems::IsValidIntegerFilter(VARIANT& filter, int minValue, int maxValue)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT)) && element.intVal >= minValue && element.intVal <= maxValue;
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL TreeViewItems::IsValidIntegerFilter(VARIANT& filter, int minValue)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT)) && element.intVal >= minValue;
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL TreeViewItems::IsValidIntegerFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_UI4));
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL TreeViewItems::IsValidStringFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = (element.vt == VT_BSTR);
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL TreeViewItems::IsValidTreeViewItemFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					if(element.vt == VT_DISPATCH && element.pdispVal) {
						CComQIPtr<ITreeViewItem, &IID_ITreeViewItem> pTvwItem(element.pdispVal);
						if(!pTvwItem) {
							// the element is invalid - abort
							isValid = FALSE;
						}
					} else {
						// we also support item handles
						switch(element.vt) {
							case VT_UI1:
							case VT_I1:
							case VT_UI2:
							case VT_I2:
							case VT_UI4:
							case VT_I4:
							case VT_UINT:
							case VT_INT:
								// the element is valid
								break;
							default:
								// the element is invalid - abort
								isValid = FALSE;
								break;
						}
					}

					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

HTREEITEM TreeViewItems::GetFirstItemToProcess(HWND hWndTvw)
{
	ATLASSERT(IsWindow(hWndTvw));

	if(properties.singleNodeAndLevel) {
		if(properties.hSimpleParentItem) {
			return reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(properties.hSimpleParentItem)));
		} else {
			return reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_ROOT), 0));
		}
	} else {
		return reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_ROOT), 0));
	}
}

HTREEITEM TreeViewItems::GetNextItemToProcess(HTREEITEM hPreviousItem, HWND hWndTvw)
{
	ATLASSERT(IsWindow(hWndTvw));

	if(properties.singleNodeAndLevel) {
		return reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hPreviousItem)));
	} else {
		if(properties.pOwnerExTvw->IsComctl32Version610OrNewer() && ((properties.filterType[fpSelected] == ftIncluding) && !properties.comparisonFunction[fpSelected])) {
			// jump directly to the next selected item
			return reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXTSELECTED), reinterpret_cast<LPARAM>(hPreviousItem)));
		} else {
			// try the first sub-item at first
			HTREEITEM hDummy = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(hPreviousItem)));
			if(!hDummy) {
				// the item doesn't have sub-items - try its successor
				hDummy = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hPreviousItem)));
				if(!hDummy) {
					/* There is no successor - try the parent item's successor and if we again don't have
						 success, go upwardly until we find an item with a successor. */
					hDummy = hPreviousItem;
					HTREEITEM hDummy2 = NULL;
					do {
						hDummy = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hDummy)));
						/*if(hDummy == properties.hSimpleParentItem) {
							hDummy = NULL;
						}*/
						if(hDummy) {
							hDummy2 = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hDummy)));
						}
					} while(hDummy && !hDummy2);
					hDummy = hDummy2;
				}
			}

			return hDummy;
		}
	}
}

BOOL TreeViewItems::IsBooleanInSafeArray(LPSAFEARRAY pSafeArray, VARIANT_BOOL value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(VARIANT_BOOL, VARIANT_BOOL);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, element.boolVal)) {
				foundMatch = TRUE;
			}
		} else {
			if(element.boolVal == value) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL TreeViewItems::IsIntegerInSafeArray(LPSAFEARRAY pSafeArray, int value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		int v = 0;
		if(SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT))) {
			v = element.intVal;
		}
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(LONG, LONG);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, v)) {
				foundMatch = TRUE;
			}
		} else {
			if(v == value) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL TreeViewItems::IsStringInSafeArray(LPSAFEARRAY pSafeArray, BSTR value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(BSTR, BSTR);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, element.bstrVal)) {
				foundMatch = TRUE;
			}
		} else {
			if(properties.caseSensitiveFilters) {
				if(lstrcmpW(OLE2W(element.bstrVal), OLE2W(value)) == 0) {
					foundMatch = TRUE;
				}
			} else {
				if(lstrcmpiW(OLE2W(element.bstrVal), OLE2W(value)) == 0) {
					foundMatch = TRUE;
				}
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL TreeViewItems::IsPartOfCollection(HTREEITEM hItem, BOOL skipFPSelected/* = FALSE*/, HWND hWndTvw/* = NULL*/, BOOL forceMultiLevel/* = FALSE*/)
{
	if(!hItem) {
		return FALSE;
	}

	BOOL isPart = FALSE;
	// we declare this one here to avoid compiler warnings
	TVITEMEX item = {0};

	if(properties.filterType[fpHandle] != ftDeactivated) {
		LONG lBound = 0;
		LONG uBound = 0;
		SafeArrayGetLBound(properties.filter[fpHandle].parray, 1, &lBound);
		SafeArrayGetUBound(properties.filter[fpHandle].parray, 1, &uBound);
		VARIANT element;
		BOOL foundMatch = FALSE;
		for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
			SafeArrayGetElement(properties.filter[fpHandle].parray, &i, &element);
			HTREEITEM hItemToCompareWith = NULL;
			if(SUCCEEDED(VariantChangeType(&element, &element, 0, VT_UI4))) {
				hItemToCompareWith = static_cast<HTREEITEM>(LongToHandle(element.ulVal));
			}
			if(properties.comparisonFunction[fpHandle]) {
				typedef BOOL WINAPI ComparisonFn(LONG, LONG);
				ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(properties.comparisonFunction[fpHandle]);
				if(pComparisonFn(HandleToLong(hItem), HandleToLong(hItemToCompareWith))) {
					foundMatch = TRUE;
				}
			} else {
				if(hItem == hItemToCompareWith) {
					foundMatch = TRUE;
				}
			}
			VariantClear(&element);
		}

		if(foundMatch) {
			if(properties.filterType[fpHandle] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpHandle] == ftIncluding) {
				goto Exit;
			}
		}
	}

	if(!hWndTvw) {
		hWndTvw = properties.GetExTvwHWnd();
	}
	ATLASSERT(IsWindow(hWndTvw));

	if(properties.singleNodeAndLevel && !forceMultiLevel) {
		HTREEITEM hParent = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hItem)));
		if(hParent != properties.hSimpleParentItem) {
			goto Exit;
		}
	} else if(properties.filterType[fpParentItem] != ftDeactivated) {
		HTREEITEM hParent = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hItem)));
		LONG lBound = 0;
		LONG uBound = 0;
		SafeArrayGetLBound(properties.filter[fpParentItem].parray, 1, &lBound);
		SafeArrayGetUBound(properties.filter[fpParentItem].parray, 1, &uBound);
		VARIANT element;
		BOOL foundMatch = FALSE;
		for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
			SafeArrayGetElement(properties.filter[fpParentItem].parray, &i, &element);
			HTREEITEM hItemToCompareWith = NULL;
			if(element.vt == VT_DISPATCH) {
				CComQIPtr<ITreeViewItem, &IID_ITreeViewItem> pTvwItem(element.pdispVal);
				if(pTvwItem) {
					OLE_HANDLE h = NULL;
					pTvwItem->get_Handle(&h);
					hItemToCompareWith = static_cast<HTREEITEM>(LongToHandle(h));
				}
			} else if(SUCCEEDED(VariantChangeType(&element, &element, 0, VT_UI4))) {
				hItemToCompareWith = static_cast<HTREEITEM>(LongToHandle(element.ulVal));
			}
			if(properties.comparisonFunction[fpParentItem]) {
				typedef BOOL WINAPI ComparisonFn(ITreeViewItem*, ITreeViewItem*);
				ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(properties.comparisonFunction[fpParentItem]);
				CComPtr<ITreeViewItem> pTvwItem1 = ClassFactory::InitTreeItem(hParent, properties.pOwnerExTvw);
				CComPtr<ITreeViewItem> pTvwItem2 = ClassFactory::InitTreeItem(hItemToCompareWith, properties.pOwnerExTvw);
				if(pComparisonFn(pTvwItem1, pTvwItem2)) {
					foundMatch = TRUE;
				}
			} else {
				if(hParent == hItemToCompareWith) {
					foundMatch = TRUE;
				}
			}
			VariantClear(&element);
		}

		if(foundMatch) {
			if(properties.filterType[fpParentItem] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpParentItem] == ftIncluding) {
				goto Exit;
			}
		}
	}

	if(properties.filterType[fpLevel] != ftDeactivated) {
		int level = 0;
		HTREEITEM hParent = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hItem)));
		while(hParent) {
			++level;
			hParent = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hItem)));
		}

		if(IsIntegerInSafeArray(properties.filter[fpLevel].parray, level, properties.comparisonFunction[fpLevel])) {
			if(properties.filterType[fpLevel] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpLevel] == ftIncluding) {
				goto Exit;
			}
		}
	}

	item.hItem = hItem;
	item.mask |= TVIF_HANDLE;
	BOOL mustRetrieveItemData = FALSE;
	if(properties.filterType[fpBold] != ftDeactivated) {
		item.mask |= TVIF_STATE;
		item.stateMask |= TVIS_BOLD;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpExpansionState] != ftDeactivated) {
		item.mask |= TVIF_STATE;
		item.stateMask |= TVIS_EXPANDED | TVIS_EXPANDEDONCE | TVIS_EXPANDPARTIAL;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpGhosted] != ftDeactivated) {
		item.mask |= TVIF_STATE;
		item.stateMask |= TVIS_CUT;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpHasExpando] != ftDeactivated) {
		item.mask |= TVIF_CHILDREN;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpHeightIncrement] != ftDeactivated) {
		item.mask |= TVIF_INTEGRAL;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpIconIndex] != ftDeactivated) {
		item.mask |= TVIF_IMAGE;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpItemData] != ftDeactivated) {
		item.mask |= TVIF_PARAM;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpOverlayIndex] != ftDeactivated) {
		item.mask |= TVIF_STATE;
		item.stateMask |= TVIS_OVERLAYMASK;
		mustRetrieveItemData = TRUE;
	}
	if(!skipFPSelected && (properties.filterType[fpSelected] != ftDeactivated)) {
		item.mask |= TVIF_STATE;
		item.stateMask |= TVIS_SELECTED;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpSelectedIconIndex] != ftDeactivated) {
		item.mask |= TVIF_SELECTEDIMAGE;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpStateImageIndex] != ftDeactivated) {
		item.mask |= TVIF_STATE;
		item.stateMask |= TVIS_STATEIMAGEMASK;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpText] != ftDeactivated) {
		item.mask |= TVIF_TEXT;
		item.cchTextMax = MAX_ITEMTEXTLENGTH;
		item.pszText = reinterpret_cast<LPTSTR>(malloc((item.cchTextMax + 1) * sizeof(TCHAR)));
		mustRetrieveItemData = TRUE;
	}
	if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
		if(properties.filterType[fpExpandedIconIndex] != ftDeactivated) {
			item.mask |= TVIF_EXPANDEDIMAGE;
			mustRetrieveItemData = TRUE;
		}
		if((properties.filterType[fpGrayed] != ftDeactivated) || (properties.filterType[fpVirtual] != ftDeactivated)) {
			item.mask |= TVIF_STATEEX;
			mustRetrieveItemData = TRUE;
		}
	}

	if(mustRetrieveItemData) {
		if(!SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			goto Exit;
		}

		// apply the filters

		if(!skipFPSelected && (properties.filterType[fpSelected] != ftDeactivated)) {
			if(IsBooleanInSafeArray(properties.filter[fpSelected].parray, BOOL2VARIANTBOOL((item.state & TVIS_SELECTED) == TVIS_SELECTED), properties.comparisonFunction[fpSelected])) {
				if(properties.filterType[fpSelected] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpSelected] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpStateImageIndex] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpStateImageIndex].parray, (item.state & TVIS_STATEIMAGEMASK) >> 12, properties.comparisonFunction[fpStateImageIndex])) {
				if(properties.filterType[fpStateImageIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpStateImageIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpItemData] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpItemData].parray, static_cast<int>(item.lParam), properties.comparisonFunction[fpItemData])) {
				if(properties.filterType[fpItemData] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpItemData] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
			if(properties.filterType[fpVirtual] != ftDeactivated) {
				if(IsBooleanInSafeArray(properties.filter[fpVirtual].parray, BOOL2VARIANTBOOL((item.uStateEx & TVIS_EX_FLAT) == TVIS_EX_FLAT), properties.comparisonFunction[fpVirtual])) {
					if(properties.filterType[fpVirtual] == ftExcluding) {
						goto Exit;
					}
				} else {
					if(properties.filterType[fpVirtual] == ftIncluding) {
						goto Exit;
					}
				}
			}

			if(properties.filterType[fpGrayed] != ftDeactivated) {
				if(IsBooleanInSafeArray(properties.filter[fpGrayed].parray, BOOL2VARIANTBOOL((item.uStateEx & TVIS_EX_DISABLED) == TVIS_EX_DISABLED), properties.comparisonFunction[fpGrayed])) {
					if(properties.filterType[fpGrayed] == ftExcluding) {
						goto Exit;
					}
				} else {
					if(properties.filterType[fpGrayed] == ftIncluding) {
						goto Exit;
					}
				}
			}

			if(properties.filterType[fpExpandedIconIndex] != ftDeactivated) {
				if(IsIntegerInSafeArray(properties.filter[fpExpandedIconIndex].parray, item.iExpandedImage, properties.comparisonFunction[fpExpandedIconIndex])) {
					if(properties.filterType[fpExpandedIconIndex] == ftExcluding) {
						goto Exit;
					}
				} else {
					if(properties.filterType[fpExpandedIconIndex] == ftIncluding) {
						goto Exit;
					}
				}
			}
		}

		if(properties.filterType[fpBold] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpBold].parray, BOOL2VARIANTBOOL((item.state & TVIS_BOLD) == TVIS_BOLD), properties.comparisonFunction[fpBold])) {
				if(properties.filterType[fpBold] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpBold] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpGhosted] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpGhosted].parray, BOOL2VARIANTBOOL((item.state & TVIS_CUT) == TVIS_CUT), properties.comparisonFunction[fpGhosted])) {
				if(properties.filterType[fpGhosted] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpGhosted] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpExpansionState] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpExpansionState].parray, item.state & (TVIS_EXPANDED | TVIS_EXPANDEDONCE | TVIS_EXPANDPARTIAL), properties.comparisonFunction[fpExpansionState])) {
				if(properties.filterType[fpExpansionState] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpExpansionState] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpHasExpando] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpHasExpando].parray, item.cChildren, properties.comparisonFunction[fpHasExpando])) {
				if(properties.filterType[fpHasExpando] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpHasExpando] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpOverlayIndex] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpOverlayIndex].parray, (item.state & TVIS_OVERLAYMASK) >> 8, properties.comparisonFunction[fpOverlayIndex])) {
				if(properties.filterType[fpOverlayIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpOverlayIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpIconIndex] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpIconIndex].parray, item.iImage, properties.comparisonFunction[fpIconIndex])) {
				if(properties.filterType[fpIconIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpIconIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpSelectedIconIndex] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpSelectedIconIndex].parray, item.iSelectedImage, properties.comparisonFunction[fpSelectedIconIndex])) {
				if(properties.filterType[fpSelectedIconIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpSelectedIconIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpHeightIncrement] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpHeightIncrement].parray, item.iIntegral, properties.comparisonFunction[fpHeightIncrement])) {
				if(properties.filterType[fpHeightIncrement] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpHeightIncrement] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpText] != ftDeactivated) {
			if(IsStringInSafeArray(properties.filter[fpText].parray, CComBSTR(item.pszText), properties.comparisonFunction[fpText])) {
				if(properties.filterType[fpText] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpText] == ftIncluding) {
					goto Exit;
				}
			}
		}
	}
	isPart = TRUE;

Exit:
	if(item.pszText) {
		SECUREFREE(item.pszText);
	}
	return isPart;
}

#ifdef INCLUDESHELLBROWSERINTERFACE
	HTREEITEM TreeViewItems::Add(LPTSTR pItemText, HTREEITEM hParentItem, HTREEITEM hInsertAfter, int hasExpando, int iconIndex, int selectedIconIndex, int expandedIconIndex, int overlayIndex, LPARAM itemData, BOOL isGhosted, BOOL isVirtual, int heightIncrement, BOOL setShellItemFlag)
#else
	HTREEITEM TreeViewItems::Add(LPTSTR pItemText, HTREEITEM hParentItem, HTREEITEM hInsertAfter, int hasExpando, int iconIndex, int selectedIconIndex, int expandedIconIndex, int overlayIndex, LPARAM itemData, BOOL isGhosted, BOOL isVirtual, int heightIncrement)
#endif
{
	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	TVINSERTSTRUCT insertionData = {0};
	insertionData.hParent = hParentItem;
	insertionData.hInsertAfter = hInsertAfter;
	insertionData.itemex.cChildren = hasExpando;
	insertionData.itemex.iIntegral = heightIncrement;
	insertionData.itemex.mask = TVIF_CHILDREN | TVIF_INTEGRAL | TVIF_PARAM | TVIF_TEXT;
	insertionData.itemex.lParam = itemData;

	if(iconIndex != -2) {
		insertionData.itemex.iImage = iconIndex;
		insertionData.itemex.mask |= TVIF_IMAGE;
	}
	if(selectedIconIndex != -2) {
		insertionData.itemex.iSelectedImage = selectedIconIndex;
		insertionData.itemex.mask |= TVIF_SELECTEDIMAGE;
	} else if(iconIndex != -2) {
		insertionData.itemex.iSelectedImage = insertionData.itemex.iImage;
		insertionData.itemex.mask |= TVIF_SELECTEDIMAGE;
	}
	if(overlayIndex > 0) {
		insertionData.itemex.state |= INDEXTOOVERLAYMASK(overlayIndex);
		insertionData.itemex.stateMask |= TVIS_OVERLAYMASK;
		insertionData.itemex.mask |= TVIF_STATE;
	}
	if(isGhosted) {
		insertionData.itemex.state |= TVIS_CUT;
		insertionData.itemex.stateMask |= TVIS_CUT;
		insertionData.itemex.mask |= TVIF_STATE;
	}
	if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
		if(expandedIconIndex != -2) {
			insertionData.itemex.iExpandedImage = expandedIconIndex;
			insertionData.itemex.mask |= TVIF_EXPANDEDIMAGE;
		} else if(iconIndex != -2) {
			insertionData.itemex.iExpandedImage = insertionData.itemex.iImage;
			insertionData.itemex.mask |= TVIF_EXPANDEDIMAGE;
		}
		if(isVirtual) {
			insertionData.itemex.uStateEx |= TVIS_EX_FLAT;
			insertionData.itemex.mask |= TVIF_STATEEX;
		}
	}

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(setShellItemFlag) {
			insertionData.itemex.mask |= TVIF_ISASHELLITEM;
		}
	#endif

	insertionData.itemex.pszText = pItemText;
	return reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData)));
}

BOOL TreeViewItems::IsSelected(HTREEITEM hItem, HWND hWndTvw/* = NULL*/)
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

#ifdef USE_STL
	HRESULT TreeViewItems::MoveItems(std::vector<HTREEITEM>& itemsToMove, HTREEITEM hNewParentItem, HWND hWndTvw/* = NULL*/)
#else
	HRESULT TreeViewItems::MoveItems(CAtlArray<HTREEITEM>& itemsToMove, HTREEITEM hNewParentItem, HWND hWndTvw/* = NULL*/)
#endif
{
	if(!hWndTvw) {
		hWndTvw = properties.GetExTvwHWnd();
	}
	ATLASSERT(IsWindow(hWndTvw));

	HRESULT hr = S_OK;
	CComPtr<ITreeViewItem> pTvwItemNewParent = ClassFactory::InitTreeItem(hNewParentItem, properties.pOwnerExTvw);
	CComObject<TreeViewItem>* pTvwItemToMove = NULL;
	CComObject<TreeViewItem>::CreateInstance(&pTvwItemToMove);
	pTvwItemToMove->AddRef();
	pTvwItemToMove->SetOwner(properties.pOwnerExTvw);
	pTvwItemToMove->flags.movingWholeBranch = TRUE;

	CWindowEx(hWndTvw).InternalSetRedraw(FALSE);
	#ifdef USE_STL
		for(std::vector<HTREEITEM>::reverse_iterator iter = itemsToMove.rbegin(); iter != itemsToMove.rend(); ++iter) {
			pTvwItemToMove->Attach(*iter);
	#else
		for(int i = itemsToMove.GetCount() - 1; i >= 0; --i) {
			pTvwItemToMove->Attach(itemsToMove[i]);
	#endif
		hr = pTvwItemToMove->Move(pTvwItemNewParent, _variant_t(HandleToLong(TVI_FIRST)));
		if(hr != S_OK) {
			break;
		}
	}

	pTvwItemToMove->flags.movingWholeBranch = FALSE;
	pTvwItemToMove->Release();
	pTvwItemToMove = NULL;

	if(SUCCEEDED(hr)) {
		properties.hSimpleParentItem = hNewParentItem;
		properties.hLastEnumeratedItem = NULL;
	}
	CWindowEx(hWndTvw).InternalSetRedraw(TRUE);

	return hr;
}

void TreeViewItems::OptimizeFilter(FilteredPropertyConstants filteredProperty)
{
	if(!(filteredProperty == fpBold || filteredProperty == fpGhosted || filteredProperty == fpGrayed || filteredProperty == fpSelected || filteredProperty == fpVirtual)) {
		// currently we optimize boolean filters only
		return;
	}

	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(properties.filter[filteredProperty].parray, 1, &lBound);
	SafeArrayGetUBound(properties.filter[filteredProperty].parray, 1, &uBound);

	// now that we have the bounds, iterate the array
	VARIANT element;
	UINT numberOfTrues = 0;
	UINT numberOfFalses = 0;
	for(LONG i = lBound; i <= uBound; ++i) {
		SafeArrayGetElement(properties.filter[filteredProperty].parray, &i, &element);
		if(element.boolVal == VARIANT_FALSE) {
			++numberOfFalses;
		} else {
			++numberOfTrues;
		}

		VariantClear(&element);
	}

	if(numberOfTrues > 0 && numberOfFalses > 0) {
		// we've something like True Or False Or True - we can deactivate this filter
		properties.filterType[filteredProperty] = ftDeactivated;
		VariantClear(&properties.filter[filteredProperty]);
	} else if(numberOfTrues == 0 && numberOfFalses > 1) {
		// False Or False Or False... is still False, so we need just one False
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_FALSE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	} else if(numberOfFalses == 0 && numberOfTrues > 1) {
		// True Or True Or True... is still True, so we need just one True
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_TRUE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	}
}

#ifdef USE_STL
	HRESULT TreeViewItems::RemoveItems(std::vector<HTREEITEM>& itemsToRemove, HWND hWndTvw)
#else
	HRESULT TreeViewItems::RemoveItems(CAtlArray<HTREEITEM>& itemsToRemove, HWND hWndTvw)
#endif
{
	ATLASSERT(IsWindow(hWndTvw));

	CWindowEx(hWndTvw).InternalSetRedraw(FALSE);
	#ifdef USE_STL
		for(std::vector<HTREEITEM>::reverse_iterator iter = itemsToRemove.rbegin(); iter < itemsToRemove.rend(); ++iter) {
			SendMessage(hWndTvw, TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(*iter));
		}
	#else
		for(int i = itemsToRemove.GetCount() - 1; i >= 0; --i) {
			SendMessage(hWndTvw, TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(itemsToRemove[i]));
		}
	#endif
	CWindowEx(hWndTvw).InternalSetRedraw(TRUE);

	return S_OK;
}


TreeViewItems::Properties::~Properties()
{
	for(int i = 0; i < NUMBEROFFILTERS; ++i) {
		VariantClear(&filter[i]);
	}
	if(pOwnerExTvw) {
		pOwnerExTvw->Release();
	}
}

void TreeViewItems::Properties::CopyTo(TreeViewItems::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerExTvw = this->pOwnerExTvw;
		if(pTarget->pOwnerExTvw) {
			pTarget->pOwnerExTvw->AddRef();
		}
		pTarget->hLastEnumeratedItem = this->hLastEnumeratedItem;
		pTarget->hSimpleParentItem = this->hSimpleParentItem;
		pTarget->caseSensitiveFilters = this->caseSensitiveFilters;
		pTarget->singleNodeAndLevel = this->singleNodeAndLevel;

		for(int i = 0; i < NUMBEROFFILTERS; ++i) {
			VariantCopy(&pTarget->filter[i], &this->filter[i]);
			pTarget->filterType[i] = this->filterType[i];
			pTarget->comparisonFunction[i] = this->comparisonFunction[i];
		}
		pTarget->usingFilters = this->usingFilters;
	}
}

HWND TreeViewItems::Properties::GetExTvwHWnd(void)
{
	ATLASSUME(pOwnerExTvw);

	OLE_HANDLE handle = NULL;
	pOwnerExTvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void TreeViewItems::SetOwner(ExplorerTreeView* pOwner)
{
	if(properties.pOwnerExTvw) {
		properties.pOwnerExTvw->Release();
	}
	properties.pOwnerExTvw = pOwner;
	if(properties.pOwnerExTvw) {
		properties.pOwnerExTvw->AddRef();
	}
}


STDMETHODIMP TreeViewItems::get_CaseSensitiveFilters(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.caseSensitiveFilters);
	return S_OK;
}

STDMETHODIMP TreeViewItems::put_CaseSensitiveFilters(VARIANT_BOOL newValue)
{
	properties.caseSensitiveFilters = VARIANTBOOL2BOOL(newValue);
	return S_OK;
}

STDMETHODIMP TreeViewItems::get_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(filteredProperty >= 0 && filteredProperty < NUMBEROFFILTERS) {
		*pValue = static_cast<LONG>(reinterpret_cast<LONG_PTR>(properties.comparisonFunction[filteredProperty]));
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP TreeViewItems::put_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG newValue)
{
	if(filteredProperty >= 0 && filteredProperty < NUMBEROFFILTERS) {
		properties.comparisonFunction[filteredProperty] = reinterpret_cast<LPVOID>(static_cast<LONG_PTR>(newValue));
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP TreeViewItems::get_Filter(FilteredPropertyConstants filteredProperty, VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	if(filteredProperty >= 0 && filteredProperty < NUMBEROFFILTERS) {
		VariantClear(pValue);
		if(properties.singleNodeAndLevel && filteredProperty == fpParentItem) {
			// hardcode the filter to properties.hSimpleParentItem

			// create the array
			pValue->vt = VT_ARRAY | VT_VARIANT;
			pValue->parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

			// store properties.hSimpleParentItem in the array (as object)
			VARIANT element;
			VariantInit(&element);
			element.vt = VT_DISPATCH;
			ClassFactory::InitTreeItem(properties.hSimpleParentItem, properties.pOwnerExTvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&element.pdispVal));
			LONG index = 1;
			SafeArrayPutElement(pValue->parray, &index, &element);
			VariantClear(&element);
		} else {
			VariantCopy(pValue, &properties.filter[filteredProperty]);
		}
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP TreeViewItems::put_Filter(FilteredPropertyConstants filteredProperty, VARIANT newValue)
{
	if(filteredProperty >= 0 && filteredProperty < NUMBEROFFILTERS) {
		if(!(properties.singleNodeAndLevel && filteredProperty == fpParentItem)) {
			// check 'newValue'
			switch(filteredProperty) {
				case fpBold:
				case fpGhosted:
				case fpSelected:
				case fpVirtual:
					if(!IsValidBooleanFilter(newValue)) {
						// invalid value - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					break;
				case fpExpansionState:
				case fpHandle:
				case fpItemData:
					if(!IsValidIntegerFilter(newValue)) {
						// invalid value - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					break;
				case fpHasExpando:
					if(!IsValidIntegerFilter(newValue, -2, 1)) {
						// invalid value - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					break;
				case fpHeightIncrement:
					if(!IsValidIntegerFilter(newValue, 1)) {
						// invalid value - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					break;
				case fpIconIndex:
				case fpExpandedIconIndex:
				case fpSelectedIconIndex:
					if(!IsValidIntegerFilter(newValue, -1)) {
						// invalid value - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					break;
				case fpLevel:
				case fpOverlayIndex:
				case fpStateImageIndex:
					if(!IsValidIntegerFilter(newValue, 0)) {
						// invalid value - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					break;
				case fpParentItem:
					if(!IsValidTreeViewItemFilter(newValue)) {
						// invalid value - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					break;
				case fpText:
					if(!IsValidStringFilter(newValue)) {
						// invalid value - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					break;
			}

			VariantClear(&properties.filter[filteredProperty]);
			VariantCopy(&properties.filter[filteredProperty], &newValue);
			OptimizeFilter(filteredProperty);
		}
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP TreeViewItems::get_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, FilterTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(filteredProperty >= 0 && filteredProperty < NUMBEROFFILTERS) {
		if(properties.singleNodeAndLevel && filteredProperty == fpParentItem) {
			// hardcode the filter to ftIncluding
			*pValue = ftIncluding;
		} else {
			*pValue = properties.filterType[filteredProperty];
		}
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP TreeViewItems::put_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants newValue)
{
	if(newValue < 0 || newValue > 2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(filteredProperty >= 0 && filteredProperty < NUMBEROFFILTERS) {
		if(!(properties.singleNodeAndLevel && filteredProperty == fpParentItem)) {
			properties.filterType[filteredProperty] = newValue;
			if(newValue != ftDeactivated) {
				properties.usingFilters = TRUE;
			} else {
				properties.usingFilters = FALSE;
				for(int i = 0; i < NUMBEROFFILTERS; ++i) {
					if(properties.filterType[i] != ftDeactivated) {
						properties.usingFilters = TRUE;
						break;
					}
				}
			}
		}
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP TreeViewItems::get_Item(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitHandle*/, ITreeViewItem** ppItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppItem, ITreeViewItem*);
	if(!ppItem) {
		return E_POINTER;
	}

	HTREEITEM hItem = NULL;
	switch(itemIdentifierType) {
		case iitID:
			if(properties.pOwnerExTvw) {
				hItem = properties.pOwnerExTvw->IDToItemHandle(itemIdentifier);
			}
			break;
		case iitHandle:
			hItem = static_cast<HTREEITEM>(LongToHandle(itemIdentifier));
			break;
		case iitAccessibilityID:
			HWND hWndTvw = properties.GetExTvwHWnd();
			ATLASSERT(IsWindow(hWndTvw));

			if(RunTimeHelper::IsCommCtrl6()) {
				hItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_MAPACCIDTOHTREEITEM, static_cast<UINT>(itemIdentifier), 0));
			}
			break;
	}

	if(hItem && IsPartOfCollection(hItem)) {
		ClassFactory::InitTreeItem(hItem, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppItem));
		return S_OK;
	}

	return E_INVALIDARG;
}

STDMETHODIMP TreeViewItems::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}

STDMETHODIMP TreeViewItems::get_ParentItem(ITreeViewItem** ppParentItem)
{
	ATLASSERT_POINTER(ppParentItem, ITreeViewItem*);
	if(!ppParentItem) {
		return E_POINTER;
	}

	if(properties.singleNodeAndLevel) {
		ClassFactory::InitTreeItem(properties.hSimpleParentItem, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppParentItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TreeViewItems::putref_ParentItem(ITreeViewItem* pNewParentItem)
{
	HTREEITEM hNewParentItem = NULL;
	if(pNewParentItem) {
		OLE_HANDLE h = NULL;
		pNewParentItem->get_Handle(&h);
		hNewParentItem = static_cast<HTREEITEM>(LongToHandle(h));
	}

	HRESULT hr = E_FAIL;
	if(properties.singleNodeAndLevel) {
		// move the immediate sub-items of properties.hSimpleParentItem only

		HWND hWndTvw = properties.GetExTvwHWnd();
		ATLASSERT(IsWindow(hWndTvw));

		// iterate each immediate sub-item of hParentItem
		#ifdef USE_STL
			std::vector<HTREEITEM> itemsToMove;
		#else
			CAtlArray<HTREEITEM> itemsToMove;
		#endif
		// get the first item to process
		HTREEITEM hItem = GetFirstItemToProcess(hWndTvw);
		while(hItem) {
			if(IsPartOfCollection(hItem, FALSE, hWndTvw)) {
				// move the item
				#ifdef USE_STL
					itemsToMove.push_back(hItem);
				#else
					itemsToMove.Add(hItem);
				#endif
			}
			hItem = GetNextItemToProcess(hItem, hWndTvw);
		}

		// now move all found items
		hr = MoveItems(itemsToMove, hNewParentItem, hWndTvw);
		if(hr == S_OK) {
			if(hNewParentItem) {
				properties.hSimpleParentItem = hNewParentItem;
			} else {
				properties.hSimpleParentItem = NULL;
			}
		}
	}

	return hr;
}

STDMETHODIMP TreeViewItems::get_SingleLevel(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.singleNodeAndLevel) {
		*pValue = VARIANT_TRUE;
	} else if(properties.filterType[fpParentItem] == ftIncluding) {
		if(IsValidTreeViewItemFilter(properties.filter[fpParentItem])) {
			*pValue = VARIANT_TRUE;
		} else {
			*pValue = VARIANT_FALSE;
		}
	} else {
		*pValue = VARIANT_FALSE;
	}
	return S_OK;
}


STDMETHODIMP TreeViewItems::Add(BSTR itemText, ITreeViewItem* pParentItem/* = NULL*/, VARIANT insertAfter/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, HasExpandoConstants hasExpando/* = heNo*/, LONG iconIndex/* = -2*/, LONG selectedIconIndex/* = -2*/, LONG expandedIconIndex/* = -2*/, LONG itemData/* = 0*/, VARIANT_BOOL isVirtual/* = VARIANT_FALSE*/, LONG heightIncrement/* = 1*/, ITreeViewItem** ppAddedItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedItem, ITreeViewItem*);
	if(!ppAddedItem) {
		return E_POINTER;
	}
	ATLASSERT(heightIncrement >= 1);
	if(heightIncrement < 1 || iconIndex < -2 || selectedIconIndex < -2 || expandedIconIndex < -2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;

	HTREEITEM hParentItem = NULL;
	HTREEITEM hInsertAfter = NULL;
	if(pParentItem) {
		OLE_HANDLE h = NULL;
		pParentItem->get_Handle(&h);
		hParentItem = static_cast<HTREEITEM>(LongToHandle(h));
	} else {
		if(properties.singleNodeAndLevel && (properties.hSimpleParentItem != NULL)) {
			hParentItem = properties.hSimpleParentItem;
		} else {
			hParentItem = TVI_ROOT;
		}
	}

	switch(insertAfter.vt) {
		case VT_DISPATCH: {
			CComQIPtr<ITreeViewItem, &IID_ITreeViewItem> pTvwItem(insertAfter.pdispVal);
			if(pTvwItem) {
				OLE_HANDLE h = NULL;
				pTvwItem->get_Handle(&h);
				hInsertAfter = static_cast<HTREEITEM>(LongToHandle(h));
			} else {
				// VB's 'Nothing'
				hInsertAfter = TVI_FIRST;
			}
			break;
		}
		case VT_UNKNOWN: {
			CComQIPtr<ITreeViewItem, &IID_ITreeViewItem> pTvwItem(insertAfter.punkVal);
			if(pTvwItem) {
				OLE_HANDLE h = NULL;
				pTvwItem->get_Handle(&h);
				hInsertAfter = static_cast<HTREEITEM>(LongToHandle(h));
			} else {
				hInsertAfter = TVI_LAST;
			}
			break;
		}
		case VT_EMPTY:
			hInsertAfter = TVI_LAST;
			break;
		case VT_ERROR:
			if(insertAfter.scode == DISP_E_PARAMNOTFOUND) {
				// the argument wasn't specified
				hInsertAfter = TVI_LAST;
				break;
			}
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
		default:
			VARIANT v;
			VariantInit(&v);
			if(SUCCEEDED(VariantChangeType(&v, &insertAfter, 0, VT_UI4))) {
				hInsertAfter = static_cast<HTREEITEM>(LongToHandle(v.ulVal));
				break;
			}
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}

	LPTSTR pItemText;
	#ifndef UNICODE
		COLE2T converter(itemText);
	#endif
	if(itemText) {
		#ifdef UNICODE
			pItemText = OLE2W(itemText);
		#else
			pItemText = converter;
		#endif
	} else {
		pItemText = LPSTR_TEXTCALLBACK;
	}

	// finally insert the item
	*ppAddedItem = NULL;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		HTREEITEM hItem = Add(pItemText, hParentItem, hInsertAfter, hasExpando, iconIndex, selectedIconIndex, expandedIconIndex, 0, itemData, FALSE, VARIANTBOOL2BOOL(isVirtual), heightIncrement, FALSE);
	#else
		HTREEITEM hItem = Add(pItemText, hParentItem, hInsertAfter, hasExpando, iconIndex, selectedIconIndex, expandedIconIndex, 0, itemData, FALSE, VARIANTBOOL2BOOL(isVirtual), heightIncrement);
	#endif
	if(hItem) {
		ClassFactory::InitTreeItem(hItem, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppAddedItem));
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP TreeViewItems::Contains(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitHandle*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	// get the item's handle
	HTREEITEM hItem = NULL;
	switch(itemIdentifierType) {
		case iitID:
			if(properties.pOwnerExTvw) {
				hItem = properties.pOwnerExTvw->IDToItemHandle(itemIdentifier);
				break;
			}
		case iitHandle:
			hItem = static_cast<HTREEITEM>(LongToHandle(itemIdentifier));
			break;
		case iitAccessibilityID:
			if(RunTimeHelper::IsCommCtrl6()) {
				hItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_MAPACCIDTOHTREEITEM, static_cast<UINT>(itemIdentifier), 0));
			}
			break;
	}

	*pValue = BOOL2VARIANTBOOL(IsPartOfCollection(hItem, FALSE, hWndTvw));
	return S_OK;
}

STDMETHODIMP TreeViewItems::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	if(!(properties.singleNodeAndLevel || properties.usingFilters)) {
		*pValue = static_cast<LONG>(SendMessage(hWndTvw, TVM_GETCOUNT, 0, 0));
		return S_OK;
	}

	if(properties.filterType[fpSelected] != ftDeactivated && !properties.comparisonFunction[fpSelected]) {
		// maybe we've to check the caret item only
		BOOL singleSelectionMode = TRUE;
		if(properties.pOwnerExTvw->IsComctl32Version610OrNewer()) {
			singleSelectionMode = ((SendMessage(hWndTvw, TVM_GETEXTENDEDSTYLE, 0, 0) & TVS_EX_MULTISELECT) == 0);
		}

		if(singleSelectionMode) {
			/* If the filter says either "selected items" only or "no unselected items" only, the caret item is
			   the only one in the collection. */
			LONG lBound = 0;
			SafeArrayGetLBound(properties.filter[fpSelected].parray, 1, &lBound);
			VARIANT element;
			SafeArrayGetElement(properties.filter[fpSelected].parray, &lBound, &element);
			BOOL requiredSelectionState = VARIANTBOOL2BOOL(element.boolVal);
			VariantClear(&element);

			if((requiredSelectionState && (properties.filterType[fpSelected] == ftIncluding)) || (!requiredSelectionState && (properties.filterType[fpSelected] == ftExcluding))) {
				// the caret item is the only one matching this filter - apply the other filters
				HTREEITEM hCaretItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0));
				if(IsPartOfCollection(hCaretItem, TRUE, hWndTvw)) {
					*pValue = 1;
				} else {
					*pValue = 0;
				}
				return S_OK;
			} else {
				// if fpSelected is the only active filter, only one item is not in the collection
				int numberOfActiveFilters = 0;
				for(int i = 0; i < NUMBEROFFILTERS; ++i) {
					if(properties.filterType[i] != ftDeactivated) {
						++numberOfActiveFilters;
					}
				}
				if(numberOfActiveFilters == 1) {
					*pValue = static_cast<LONG>(SendMessage(hWndTvw, TVM_GETCOUNT, 0, 0)) - 1;
					return S_OK;
				}
			}
		} else {
			int numberOfActiveFilters = 0;
			for(int i = 0; i < NUMBEROFFILTERS; i++) {
				if(properties.filterType[i] != ftDeactivated) {
					numberOfActiveFilters++;
				}
			}
			if(numberOfActiveFilters == 1) {
				// we can use TVM_GETSELECTEDCOUNT
				if(properties.filterType[fpSelected] == ftIncluding) {
					*pValue = static_cast<LONG>(SendMessage(hWndTvw, TVM_GETSELECTEDCOUNT, 0, 0));
					return S_OK;
				} else {
					*pValue = static_cast<LONG>(SendMessage(hWndTvw, TVM_GETCOUNT, 0, 0)) - static_cast<LONG>(SendMessage(hWndTvw, TVM_GETSELECTEDCOUNT, 0, 0));
					return S_OK;
				}
			}
		}
	}

	// count the items "by hand"
	*pValue = 0;
	HTREEITEM hItem = GetFirstItemToProcess(hWndTvw);
	while(hItem) {
		if(IsPartOfCollection(hItem, FALSE, hWndTvw)) {
			++(*pValue);
		}
		hItem = GetNextItemToProcess(hItem, hWndTvw);
	}
	return S_OK;
}

STDMETHODIMP TreeViewItems::FindByPath(BSTR path, VARIANT_BOOL returnDeepestMatch/* = VARIANT_FALSE*/, BSTR* pSeparator/* = NULL*/, ITreeViewItem** ppItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppItem, ITreeViewItem*);
	if(!ppItem) {
		return E_POINTER;
	}

	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	#ifdef UNICODE
		CAtlString fullPath = path;
	#else
		COLE2T converter(path);
		CAtlString fullPath = converter;
	#endif
	CAtlString separator = TEXT("\\");
	if(pSeparator) {
		separator = *pSeparator;
	}
	if(fullPath.Right(separator.GetLength()) != separator) {
		fullPath.Append(separator);
	}

	int pos = fullPath.Find(separator);
	CAtlString currentSegment = TEXT("");
	if(pos > 0) {
		currentSegment = fullPath.Left(pos);
	}
	HTREEITEM hBestMatch = NULL;
	BOOL foundExactMatch = FALSE;

	typedef BOOL WINAPI ComparisonFn(BSTR, BSTR);
	ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(properties.comparisonFunction[fpText]);

	TVITEMEX item = {0};
	item.mask = TVIF_HANDLE | TVIF_TEXT;
	item.cchTextMax = MAX_ITEMTEXTLENGTH;
	item.pszText = reinterpret_cast<LPTSTR>(malloc((item.cchTextMax + 1) * sizeof(TCHAR)));
	item.hItem = GetFirstItemToProcess(hWndTvw);
	while(item.hItem) {
		if(item.hItem && SendMessage(hWndTvw, TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			CAtlString itemText = item.pszText;

			BOOL itemTextMatches = FALSE;
			if(pComparisonFn) {
				if(pComparisonFn(itemText.AllocSysString(), currentSegment.AllocSysString())) {
					itemTextMatches = TRUE;
				}
			} else {
				if(properties.caseSensitiveFilters) {
					if(lstrcmp(itemText, currentSegment) == 0) {
						itemTextMatches = TRUE;
					}
				} else {
					if(lstrcmpi(itemText, currentSegment) == 0) {
						itemTextMatches = TRUE;
					}
				}
			}

			if(itemTextMatches && IsPartOfCollection(item.hItem, FALSE, hWndTvw, TRUE)) {
				hBestMatch = item.hItem;
				int start = pos + separator.GetLength();
				pos = fullPath.Find(separator, start);
				currentSegment = TEXT("");
				if(pos > 0) {
					currentSegment = fullPath.Mid(start, pos - start);
				} else if(pos == -1) {
					foundExactMatch = TRUE;
					break;
				}

				// go one level deeper
				item.hItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(hBestMatch)));
			} else {
				// NOTE: We can't use GetNextItemToProcess here, because we always have to search within the same level.
				item.hItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(item.hItem)));
			}
		}
	}
	SECUREFREE(item.pszText);

	if(foundExactMatch || VARIANTBOOL2BOOL(returnDeepestMatch)) {
		ClassFactory::InitTreeItem(hBestMatch, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppItem));
	} else {
		*ppItem = NULL;
	}
	return S_OK;
}

STDMETHODIMP TreeViewItems::Remove(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitHandle*/)
{
	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	// get the item's handle
	HTREEITEM hItem = NULL;
	switch(itemIdentifierType) {
		case iitID:
			if(properties.pOwnerExTvw) {
				hItem = properties.pOwnerExTvw->IDToItemHandle(itemIdentifier);
				break;
			}
		case iitHandle:
			hItem = static_cast<HTREEITEM>(LongToHandle(itemIdentifier));
			break;
		case iitAccessibilityID:
			if(RunTimeHelper::IsCommCtrl6()) {
				hItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_MAPACCIDTOHTREEITEM, static_cast<UINT>(itemIdentifier), 0));
			}
			break;
	}

	if(IsPartOfCollection(hItem, FALSE, hWndTvw)) {
		#ifdef USE_STL
			std::vector<HTREEITEM> itemsToRemove;
			itemsToRemove.push_back(hItem);
		#else
			CAtlArray<HTREEITEM> itemsToRemove;
			itemsToRemove.Add(hItem);
		#endif
		return RemoveItems(itemsToRemove, hWndTvw);
	}

	return E_INVALIDARG;
}

STDMETHODIMP TreeViewItems::RemoveAll(void)
{
	HWND hWndTvw = properties.GetExTvwHWnd();
	ATLASSERT(IsWindow(hWndTvw));

	if(!(properties.singleNodeAndLevel || properties.usingFilters)) {
		if(SendMessage(hWndTvw, TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(TVI_ROOT))) {
			return S_OK;
		} else {
			return E_FAIL;
		}
	}

	#ifdef USE_STL
		std::vector<HTREEITEM> itemsToRemove;
	#else
		CAtlArray<HTREEITEM> itemsToRemove;
	#endif
	if((properties.filterType[fpSelected] != ftDeactivated) && !properties.comparisonFunction[fpSelected]) {
		// maybe we've to check the caret item only
		if(properties.filterType[fpSelected] == ftIncluding) {
			// if the filter says "selected items" only, the caret item is the only one in the collection
			LONG lBound = 0;
			SafeArrayGetLBound(properties.filter[fpSelected].parray, 1, &lBound);
			VARIANT element;
			SafeArrayGetElement(properties.filter[fpSelected].parray, &lBound, &element);
			BOOL requiredSelectionState = VARIANTBOOL2BOOL(element.boolVal);
			VariantClear(&element);

			if(requiredSelectionState) {
				// the caret item is the only one matching this filter - apply the other filters
				HTREEITEM hCaretItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0));
				if(IsPartOfCollection(hCaretItem, TRUE, hWndTvw)) {
					#ifdef USE_STL
						itemsToRemove.push_back(hCaretItem);
					#else
						itemsToRemove.Add(hCaretItem);
					#endif
					return RemoveItems(itemsToRemove, hWndTvw);
				} else {
					return S_OK;
				}
			}
		} else {
			// if the filter says "no unselected items" only, the caret item is the only one in the collection
			LONG lBound = 0;
			SafeArrayGetLBound(properties.filter[fpSelected].parray, 1, &lBound);
			VARIANT element;
			SafeArrayGetElement(properties.filter[fpSelected].parray, &lBound, &element);
			BOOL requiredSelectionState = VARIANTBOOL2BOOL(element.boolVal);
			VariantClear(&element);

			if(!requiredSelectionState) {
				// the caret item is the only one not matching this filter - apply the other filters
				HTREEITEM hCaretItem = reinterpret_cast<HTREEITEM>(SendMessage(hWndTvw, TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0));
				if(IsPartOfCollection(hCaretItem, TRUE, hWndTvw)) {
					#ifdef USE_STL
						itemsToRemove.push_back(hCaretItem);
					#else
						itemsToRemove.Add(hCaretItem);
					#endif
					return RemoveItems(itemsToRemove, hWndTvw);
				} else {
					return S_OK;
				}
			}
		}
	}

	// find the items to remove "by hand"
	HTREEITEM hItem = GetFirstItemToProcess(hWndTvw);
	while(hItem) {
		if(IsPartOfCollection(hItem, FALSE, hWndTvw)) {
			#ifdef USE_STL
				itemsToRemove.push_back(hItem);
			#else
				itemsToRemove.Add(hItem);
			#endif
		}
		hItem = GetNextItemToProcess(hItem, hWndTvw);
	}
	return RemoveItems(itemsToRemove, hWndTvw);
}