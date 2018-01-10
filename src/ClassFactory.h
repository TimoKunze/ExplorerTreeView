//////////////////////////////////////////////////////////////////////
/// \class ClassFactory
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A helper class for creating special objects</em>
///
/// This class provides methods to create objects and initialize them with given values.
///
/// \todo Improve documentation.
///
/// \sa ExplorerTreeView
//////////////////////////////////////////////////////////////////////


#pragma once

#include "TargetOLEDataObject.h"
#include "TreeViewItem.h"
#include "TreeViewItems.h"


class ClassFactory
{
public:
	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's \c IOLEDataObject implementation as a smart pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	///
	/// \return A smart pointer to the created object's \c IOLEDataObject implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static CComPtr<IOLEDataObject> InitOLEDataObject(IDataObject* pDataObject)
	{
		CComPtr<IOLEDataObject> pOLEDataObj = NULL;
		InitOLEDataObject(pDataObject, IID_IOLEDataObject, reinterpret_cast<LPUNKNOWN*>(&pOLEDataObj));
		return pOLEDataObj;
	};

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// \overload
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's implementation of a given interface as a raw pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppDataObject Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static void InitOLEDataObject(IDataObject* pDataObject, REFIID requiredInterface, LPUNKNOWN* ppDataObject)
	{
		ATLASSERT_POINTER(ppDataObject, LPUNKNOWN);
		ATLASSUME(ppDataObject);

		*ppDataObject = NULL;

		// create an OLEDataObject object and initialize it with the given parameters
		CComObject<TargetOLEDataObject>* pOLEDataObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObj);
		pOLEDataObj->AddRef();
		pOLEDataObj->Attach(pDataObject);
		pOLEDataObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppDataObject));
		pOLEDataObj->Release();
	};

	/// \brief <em>Creates a \c TreeViewItem object</em>
	///
	/// Creates a \c TreeViewItem object that represents a given treeview item and returns
	/// its \c ITreeViewItem implementation as a smart pointer.
	///
	/// \param[in] hItem The item for which to create the object.
	/// \param[in] pExTvw The \c ExplorerTreeView object the \c TreeViewItem object will work on.
	///
	/// \return A smart pointer to the created object's \c ITreeViewItem implementation.
	static CComPtr<ITreeViewItem> InitTreeItem(HTREEITEM hItem, ExplorerTreeView* pExTvw)
	{
		CComPtr<ITreeViewItem> pTreeItem = NULL;
		InitTreeItem(hItem, pExTvw, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pTreeItem));
		return pTreeItem;
	};

	/// \brief <em>Creates a \c TreeViewItem object</em>
	///
	/// \overload
	///
	/// Creates a \c TreeViewItem object that represents a given treeview item and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] hItem The item for which to create the object.
	/// \param[in] pExTvw The \c ExplorerTreeView object the \c TreeViewItem object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppTreeItem Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static void InitTreeItem(HTREEITEM hItem, ExplorerTreeView* pExTvw, REFIID requiredInterface, LPUNKNOWN* ppTreeItem)
	{
		ATLASSERT_POINTER(ppTreeItem, LPUNKNOWN);
		ATLASSUME(ppTreeItem);

		*ppTreeItem = NULL;
		if(ISPREDEFINEDHTREEITEM(hItem)) {
			// there's nothing useful we could return
			return;
		}

		if(hItem) {
			// create a TreeViewItem object and initialize it with the given parameters
			CComObject<TreeViewItem>* pTvwItemObj = NULL;
			CComObject<TreeViewItem>::CreateInstance(&pTvwItemObj);
			pTvwItemObj->AddRef();
			pTvwItemObj->SetOwner(pExTvw);
			pTvwItemObj->Attach(hItem);
			pTvwItemObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppTreeItem));
			pTvwItemObj->Release();
		}
	};

	/// \brief <em>Creates a \c TreeViewItems object</em>
	///
	/// Creates a \c TreeViewItems object that represents a collection of treeview items and returns
	/// its \c ITreeViewItems implementation as a smart pointer.
	///
	/// \param[in] pExTvw The \c ExplorerTreeView object the \c TreeViewItems object will work on.
	/// \param[in] hParentItem The parent item of all items in the collection. If set to -1, this kind of
	///            filter is ignored.
	/// \param[in] singleNodeAndLevel If \c TRUE, the collection contains immediate sub-items of the item
	///            specified by the \c hParentItem parameter only; otherwise it contains all items.
	///
	/// \return A smart pointer to the created object's \c ITreeViewItems implementation.
	static CComPtr<ITreeViewItems> InitTreeItems(ExplorerTreeView* pExTvw, BOOL singleNodeAndLevel = FALSE, HTREEITEM hParentItem = reinterpret_cast<HTREEITEM>(-1))
	{
		CComPtr<ITreeViewItems> pTreeItems = NULL;
		InitTreeItems(pExTvw, IID_ITreeViewItems, reinterpret_cast<LPUNKNOWN*>(&pTreeItems), singleNodeAndLevel, hParentItem);
		return pTreeItems;
	};

	/// \brief <em>Creates a \c TreeViewItems object</em>
	///
	/// \overload
	///
	/// Creates a \c TreeViewItems object that represents a collection of treeview items and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pExTvw The \c ExplorerTreeView object the \c TreeViewItems object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppTreeItems Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] hParentItem The parent item of all items in the collection. If set to -1, this kind of
	///            filter is ignored.
	/// \param[in] singleNodeAndLevel If \c TRUE, the collection contains immediate sub-items of the item
	///            specified by the \c hParentItem parameter only; otherwise it contains all items.
	static void InitTreeItems(ExplorerTreeView* pExTvw, REFIID requiredInterface, LPUNKNOWN* ppTreeItems, BOOL singleNodeAndLevel = FALSE, HTREEITEM hParentItem = reinterpret_cast<HTREEITEM>(-1))
	{
		ATLASSERT_POINTER(ppTreeItems, LPUNKNOWN);
		ATLASSUME(ppTreeItems);

		*ppTreeItems = NULL;
		// create a TreeViewItems object and initialize it with the given parameters
		CComObject<TreeViewItems>* pTvwItemsObj = NULL;
		CComObject<TreeViewItems>::CreateInstance(&pTvwItemsObj);
		pTvwItemsObj->AddRef();

		pTvwItemsObj->SetOwner(pExTvw);
		pTvwItemsObj->properties.singleNodeAndLevel = singleNodeAndLevel;
		pTvwItemsObj->properties.hSimpleParentItem = hParentItem;

		pTvwItemsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppTreeItems));
		pTvwItemsObj->Release();
	};
};     // ClassFactory