// TreeViewItemContainer.cpp: Manages a collection of TreeViewItem objects

#include "stdafx.h"
#include "TreeViewItemContainer.h"
#include "ClassFactory.h"


DWORD TreeViewItemContainer::nextID = 0;


TreeViewItemContainer::TreeViewItemContainer()
{
	containerID = ++nextID;
}

TreeViewItemContainer::~TreeViewItemContainer()
{
	properties.pOwnerExTvw->DeregisterItemContainer(containerID);
}

void TreeViewItemContainer::Properties::CopyTo(TreeViewItemContainer::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerExTvw = this->pOwnerExTvw;
		if(pTarget->pOwnerExTvw) {
			pTarget->pOwnerExTvw->AddRef();
		}
		pTarget->nextEnumeratedItem = this->nextEnumeratedItem;
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP TreeViewItemContainer::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_ITreeViewItemContainer, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP TreeViewItemContainer::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<TreeViewItemContainer>* pTvwItemContainerObj = NULL;
	CComObject<TreeViewItemContainer>::CreateInstance(&pTvwItemContainerObj);
	pTvwItemContainerObj->AddRef();

	// clone all settings
	properties.CopyTo(&pTvwItemContainerObj->properties);

	#ifdef USE_STL
		pTvwItemContainerObj->properties.items.resize(properties.items.size());
		std::copy(properties.items.begin(), properties.items.end(), pTvwItemContainerObj->properties.items.begin());
	#else
		//pTvwItemContainerObj->properties.items.Copy(properties.items);
		pTvwItemContainerObj->items.Copy(items);
	#endif

	pTvwItemContainerObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pTvwItemContainerObj->Release();

	if(*ppEnumerator) {
		properties.pOwnerExTvw->RegisterItemContainer(static_cast<IItemContainer*>(pTvwItemContainerObj));
	}
	return S_OK;
}

STDMETHODIMP TreeViewItemContainer::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		#ifdef USE_STL
			if(properties.nextEnumeratedItem >= static_cast<int>(properties.items.size())) {
				properties.nextEnumeratedItem = 0;
				// there's nothing more to iterate
				break;
			}
			HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(properties.items[properties.nextEnumeratedItem]);
		#else
			//if(properties.nextEnumeratedItem >= static_cast<int>(properties.items.GetCount())) {
			if(properties.nextEnumeratedItem >= static_cast<int>(items.GetCount())) {
				properties.nextEnumeratedItem = 0;
				// there's nothing more to iterate
				break;
			}
			//HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(properties.items[properties.nextEnumeratedItem]);
			HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(items[properties.nextEnumeratedItem]);
		#endif

		ClassFactory::InitTreeItem(hItem, properties.pOwnerExTvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal));
		pItems[i].vt = VT_DISPATCH;
		++properties.nextEnumeratedItem;
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP TreeViewItemContainer::Reset(void)
{
	properties.nextEnumeratedItem = 0;
	return S_OK;
}

STDMETHODIMP TreeViewItemContainer::Skip(ULONG numberOfItemsToSkip)
{
	properties.nextEnumeratedItem += numberOfItemsToSkip;
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IItemContainer
void TreeViewItemContainer::AddItems(UINT numberOfItems, PLONG pItemIDs)
{
	for(UINT i = 0; i < numberOfItems; i++) {
		#ifdef USE_STL
			properties.items.push_back(pItemIDs[i]);
		#else
			//properties.items.Add(pItemIDs[i]);
			this->items.Add(pItemIDs[i]);
		#endif
	}
}

void TreeViewItemContainer::RemovedItem(LONG itemID)
{
	if(itemID == 0) {
		RemoveAll();
	} else {
		Remove(itemID);
	}
}

DWORD TreeViewItemContainer::GetID(void)
{
	return containerID;
}
// implementation of IItemContainer
//////////////////////////////////////////////////////////////////////


TreeViewItemContainer::Properties::~Properties()
{
	if(pOwnerExTvw) {
		pOwnerExTvw->Release();
	}
}

HWND TreeViewItemContainer::Properties::GetExTvwHWnd(void)
{
	ATLASSUME(pOwnerExTvw);

	OLE_HANDLE handle = NULL;
	pOwnerExTvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void TreeViewItemContainer::SetOwner(ExplorerTreeView* pOwner)
{
	if(properties.pOwnerExTvw) {
		properties.pOwnerExTvw->Release();
	}
	properties.pOwnerExTvw = pOwner;
	if(properties.pOwnerExTvw) {
		properties.pOwnerExTvw->AddRef();
	}
}


STDMETHODIMP TreeViewItemContainer::get_Item(LONG itemID, ITreeViewItem** ppItem)
{
	ATLASSERT_POINTER(ppItem, ITreeViewItem*);
	if(!ppItem) {
		return E_POINTER;
	}

	#ifdef USE_STL
		std::vector<LONG>::iterator iter = std::find(properties.items.begin(), properties.items.end(), itemID);
		if(iter != properties.items.end()) {
			HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(itemID);
			if(hItem) {
				ClassFactory::InitTreeItem(hItem, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppItem));
				return S_OK;
			}
		}
	#else
		//for(size_t i = 0; i < properties.items.GetCount(); ++i) {
		for(size_t i = 0; i < items.GetCount(); ++i) {
			//if(properties.items[i] == itemID) {
			if(items[i] == itemID) {
				HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(itemID);
				if(hItem) {
					ClassFactory::InitTreeItem(hItem, properties.pOwnerExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppItem));
					return S_OK;
				}
				break;
			}
		}
	#endif

	return E_INVALIDARG;
}

STDMETHODIMP TreeViewItemContainer::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP TreeViewItemContainer::Add(VARIANT items)
{
	HRESULT hr = E_FAIL;
	LONG id = 0;
	switch(items.vt) {
		case VT_DISPATCH:
			if(items.pdispVal) {
				CComQIPtr<ITreeViewItem, &IID_ITreeViewItem> pTvwItem(items.pdispVal);
				if(pTvwItem) {
					// add a single TreeViewItem object
					hr = pTvwItem->get_ID(&id);
				} else {
					CComQIPtr<ITreeViewItems, &IID_ITreeViewItems> pTvwItems(items.pdispVal);
					if(pTvwItems) {
						// add a TreeViewItems collection
						CComQIPtr<IEnumVARIANT, &IID_IEnumVARIANT> pEnumerator(pTvwItems);
						if(pEnumerator) {
							hr = S_OK;
							VARIANT item;
							VariantInit(&item);
							ULONG dummy = 0;
							while(pEnumerator->Next(1, &item, &dummy) == S_OK) {
								if((item.vt == VT_DISPATCH) && item.pdispVal) {
									pTvwItem = item.pdispVal;
									if(pTvwItem) {
										id = 0;
										pTvwItem->get_ID(&id);
										#ifdef USE_STL
											std::vector<LONG>::iterator iter = std::find(properties.items.begin(), properties.items.end(), id);
											if(iter == properties.items.end()) {
												properties.items.push_back(id);
											}
										#else
											BOOL alreadyThere = FALSE;
											//for(size_t i = 0; i < properties.items.GetCount(); ++i) {
											for(size_t i = 0; i < this->items.GetCount(); ++i) {
												//if(properties.items[i] == id) {
												if(this->items[i] == id) {
													alreadyThere = TRUE;
													break;
												}
											}
											if(!alreadyThere) {
												//properties.items.Add(id);
												this->items.Add(id);
											}
										#endif
									}
								}
								VariantClear(&item);
							}
							return S_OK;
						}
					}
				}
			}
			break;
		default:
			VARIANT v;
			VariantInit(&v);
			hr = VariantChangeType(&v, &items, 0, VT_UI4);
			id = v.ulVal;
			break;
	}
	if(FAILED(hr)) {
		// invalid arg - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	#ifdef USE_STL
		std::vector<LONG>::iterator iter = std::find(properties.items.begin(), properties.items.end(), id);
		if(iter == properties.items.end()) {
			properties.items.push_back(id);
		}
	#else
		BOOL alreadyThere = FALSE;
		//for(size_t i = 0; i < properties.items.GetCount(); ++i) {
		for(size_t i = 0; i < this->items.GetCount(); ++i) {
			//if(properties.items[i] == id) {
			if(this->items[i] == id) {
				alreadyThere = TRUE;
				break;
			}
		}
		if(!alreadyThere) {
			//properties.items.Add(id);
			this->items.Add(id);
		}
	#endif
	return S_OK;
}

STDMETHODIMP TreeViewItemContainer::Clone(ITreeViewItemContainer** ppClone)
{
	ATLASSERT_POINTER(ppClone, ITreeViewItemContainer*);
	if(!ppClone) {
		return E_POINTER;
	}

	*ppClone = NULL;
	CComObject<TreeViewItemContainer>* pTvwItemContainerObj = NULL;
	CComObject<TreeViewItemContainer>::CreateInstance(&pTvwItemContainerObj);
	pTvwItemContainerObj->AddRef();

	// clone all settings
	properties.CopyTo(&pTvwItemContainerObj->properties);

	#ifdef USE_STL
		pTvwItemContainerObj->properties.items.resize(properties.items.size());
		std::copy(properties.items.begin(), properties.items.end(), pTvwItemContainerObj->properties.items.begin());
	#else
		//pTvwItemContainerObj->properties.items.Copy(properties.items);
		pTvwItemContainerObj->items.Copy(items);
	#endif

	pTvwItemContainerObj->QueryInterface(__uuidof(ITreeViewItemContainer), reinterpret_cast<LPVOID*>(ppClone));
	pTvwItemContainerObj->Release();

	if(*ppClone) {
		properties.pOwnerExTvw->RegisterItemContainer(static_cast<IItemContainer*>(pTvwItemContainerObj));
	}
	return S_OK;
}

STDMETHODIMP TreeViewItemContainer::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef USE_STL
		*pValue = static_cast<LONG>(properties.items.size());
	#else
	//*pValue = static_cast<LONG>(properties.items.GetCount());
		*pValue = static_cast<LONG>(items.GetCount());
	#endif
	return S_OK;
}

STDMETHODIMP TreeViewItemContainer::CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft/* = NULL*/, OLE_YPOS_PIXELS* pYUpperLeft/* = NULL*/, OLE_HANDLE* phImageList/* = NULL*/)
{
	ATLASSERT_POINTER(phImageList, OLE_HANDLE);
	if(!phImageList) {
		return E_POINTER;
	}

	*phImageList = NULL;
	#ifdef USE_STL
		switch(properties.items.size()) {
	#else
		//switch(properties.items.GetCount()) {
		switch(items.GetCount()) {
	#endif
		case 0:
			return S_OK;
			break;
		case 1: {
			ATLASSUME(properties.pOwnerExTvw);
			//HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(properties.items[0]);
			#ifdef USE_STL
				HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(properties.items[0]);
			#else
				HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(items[0]);
			#endif
			if(hItem) {
				POINT upperLeftPoint = {0};
				*phImageList = HandleToLong(properties.pOwnerExTvw->CreateLegacyDragImage(hItem, &upperLeftPoint, NULL));
				if(*phImageList) {
					if(pXUpperLeft) {
						*pXUpperLeft = upperLeftPoint.x;
					}
					if(pYUpperLeft) {
						*pYUpperLeft = upperLeftPoint.y;
					}
					return S_OK;
				}
			}
			break;
		}
		default: {
			// create a large drag image out of small drag images
			ATLASSUME(properties.pOwnerExTvw);

			BOOL use32BPPImage = RunTimeHelper::IsCommCtrl6();

			// calculate the bitmap's required size and collect each item's imagelist
			#ifdef USE_STL
				std::vector<HIMAGELIST> imageLists;
				std::vector<RECT> itemBoundingRects;
			#else
				CAtlArray<HIMAGELIST> imageLists;
				CAtlArray<RECT> itemBoundingRects;
			#endif
			POINT upperLeftPoint = {0};
			WTL::CRect boundingRect;
			#ifdef USE_STL
				for(std::vector<LONG>::iterator iter = properties.items.begin(); iter != properties.items.end(); ++iter) {
					HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(*iter);
			#else
				//for(size_t i = 0; i < properties.items.GetCount(); ++i) {
				for(size_t i = 0; i < items.GetCount(); ++i) {
					//HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(properties.items[i]);
					HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(items[i]);
			#endif
				if(hItem) {
					// NOTE: Windows skips items outside the client area to improve performance. We don't.
					POINT pt = {0};
					RECT itemBoundingRect = {0};
					HIMAGELIST hImageList = properties.pOwnerExTvw->CreateLegacyDragImage(hItem, &pt, &itemBoundingRect);
					boundingRect.UnionRect(&boundingRect, &itemBoundingRect);

					#ifdef USE_STL
						if(imageLists.size() == 0) {
					#else
						if(imageLists.GetCount() == 0) {
					#endif
						upperLeftPoint = pt;
					} else {
						upperLeftPoint.x = min(upperLeftPoint.x, pt.x);
						upperLeftPoint.y = min(upperLeftPoint.y, pt.y);
					}
					#ifdef USE_STL
						imageLists.push_back(hImageList);
						itemBoundingRects.push_back(itemBoundingRect);
					#else
						imageLists.Add(hImageList);
						itemBoundingRects.Add(itemBoundingRect);
					#endif
				}
			}
			WTL::CRect dragImageRect(0, 0, boundingRect.Width(), boundingRect.Height());

			// setup the DCs we'll draw into
			HDC hCompatibleDC = GetDC(HWND_DESKTOP);
			CDC memoryDC;
			memoryDC.CreateCompatibleDC(hCompatibleDC);
			CDC maskMemoryDC;
			maskMemoryDC.CreateCompatibleDC(hCompatibleDC);

			// create the bitmap and its mask
			CBitmap dragImage;
			LPRGBQUAD pDragImageBits = NULL;
			if(use32BPPImage) {
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = dragImageRect.Width();
				bitmapInfo.bmiHeader.biHeight = -dragImageRect.Height();
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;

				dragImage.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
			} else {
				dragImage.CreateCompatibleBitmap(hCompatibleDC, dragImageRect.Width(), dragImageRect.Height());
			}
			HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(dragImage);
			memoryDC.FillRect(&dragImageRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
			CBitmap dragImageMask;
			dragImageMask.CreateBitmap(dragImageRect.Width(), dragImageRect.Height(), 1, 1, NULL);
			HBITMAP hPreviousBitmapMask = maskMemoryDC.SelectBitmap(dragImageMask);
			maskMemoryDC.FillRect(&dragImageRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));

			// draw each single drag image into our bitmap
			BOOL rightToLeft = FALSE;
			HWND hWndTvw = properties.GetExTvwHWnd();
			if(IsWindow(hWndTvw)) {
				rightToLeft = ((CWindow(hWndTvw).GetExStyle() & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
			}
			#ifdef USE_STL
				for(size_t i = 0; i < imageLists.size(); ++i) {
			#else
				for(size_t i = 0; i < imageLists.GetCount(); ++i) {
			#endif
				if(rightToLeft) {
					ImageList_Draw(imageLists[i], 0, memoryDC, boundingRect.right - itemBoundingRects[i].right, itemBoundingRects[i].top - boundingRect.top, ILD_NORMAL);
					ImageList_Draw(imageLists[i], 0, maskMemoryDC, boundingRect.right - itemBoundingRects[i].right, itemBoundingRects[i].top - boundingRect.top, ILD_MASK);
				} else {
					ImageList_Draw(imageLists[i], 0, memoryDC, itemBoundingRects[i].left - boundingRect.left, itemBoundingRects[i].top - boundingRect.top, ILD_NORMAL);
					ImageList_Draw(imageLists[i], 0, maskMemoryDC, itemBoundingRects[i].left - boundingRect.left, itemBoundingRects[i].top - boundingRect.top, ILD_MASK);
				}

				ImageList_Destroy(imageLists[i]);
			}

			// clean up
			#ifdef USE_STL
				imageLists.clear();
				itemBoundingRects.clear();
			#else
				imageLists.RemoveAll();
				itemBoundingRects.RemoveAll();
			#endif
			memoryDC.SelectBitmap(hPreviousBitmap);
			maskMemoryDC.SelectBitmap(hPreviousBitmapMask);
			ReleaseDC(HWND_DESKTOP, hCompatibleDC);

			// create the imagelist
			HIMAGELIST hImageList = ImageList_Create(dragImageRect.Width(), dragImageRect.Height(), (RunTimeHelper::IsCommCtrl6() ? ILC_COLOR32 : ILC_COLOR24) | ILC_MASK, 1, 0);
			ImageList_SetBkColor(hImageList, CLR_NONE);
			ImageList_Add(hImageList, dragImage, dragImageMask);
			*phImageList = HandleToLong(hImageList);

			if(*phImageList) {
				if(pXUpperLeft) {
					*pXUpperLeft = upperLeftPoint.x;
				}
				if(pYUpperLeft) {
					*pYUpperLeft = upperLeftPoint.y;
				}
				return S_OK;
			}
			break;
		}
	}

	return E_FAIL;
}

STDMETHODIMP TreeViewItemContainer::Remove(LONG itemID, VARIANT_BOOL removePhysically/* = VARIANT_FALSE*/)
{
	#ifdef USE_STL
		for(size_t i = 0; i < properties.items.size(); ++i) {
			if(properties.items[i] == itemID) {
				if(removePhysically == VARIANT_FALSE) {
					properties.items.erase(std::find(properties.items.begin(), properties.items.end(), itemID));
					if(i < static_cast<size_t>(properties.nextEnumeratedItem)) {
						--properties.nextEnumeratedItem;
					}
				} else {
					HWND hWndTvw = properties.GetExTvwHWnd();
					if(IsWindow(hWndTvw)) {
						HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(itemID);
						SendMessage(hWndTvw, TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(hItem));
					}

					// our owner will notify us about the deletion, so we don't need to remove the item explicitly
				}

				return S_OK;
			}
		}
	#else
		//for(size_t i = 0; i < properties.items.GetCount(); ++i) {
		for(size_t i = 0; i < items.GetCount(); ++i) {
			//if(properties.items[i] == itemID) {
			if(items[i] == itemID) {
				if(removePhysically == VARIANT_FALSE) {
					//properties.items.RemoveAt(i);
					items.RemoveAt(i);
					if(i < static_cast<size_t>(properties.nextEnumeratedItem)) {
						--properties.nextEnumeratedItem;
					}
				} else {
					HWND hWndTvw = properties.GetExTvwHWnd();
					if(IsWindow(hWndTvw)) {
						HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(itemID);
						SendMessage(hWndTvw, TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(hItem));
					}

					// our owner will notify us about the deletion, so we don't need to remove the item explicitly
				}

				return S_OK;
			}
		}
	#endif
	return E_FAIL;
}

STDMETHODIMP TreeViewItemContainer::RemoveAll(VARIANT_BOOL removePhysically/* = VARIANT_FALSE*/)
{
	if(removePhysically != VARIANT_FALSE) {
		HWND hWndTvw = properties.GetExTvwHWnd();
		if(IsWindow(hWndTvw)) {
			#ifdef USE_STL
				while(properties.items.size() > 0) {
					HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(*properties.items.begin());
					ATLASSERT(hItem != NULL);
					SendMessage(hWndTvw, TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(hItem));
				}
			#else
				//while(properties.items.GetCount() > 0) {
				while(items.GetCount() > 0) {
					//HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(properties.items[0]);
					HTREEITEM hItem = properties.pOwnerExTvw->IDToItemHandle(items[0]);
					ATLASSERT(hItem != NULL);
					SendMessage(hWndTvw, TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(hItem));
				}
			#endif
		}
	} else {
		#ifdef USE_STL
			properties.items.clear();
		#else
			//properties.items.RemoveAll();
			items.RemoveAll();
		#endif
	}
	Reset();
	return S_OK;
}