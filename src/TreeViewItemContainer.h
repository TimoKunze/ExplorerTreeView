//////////////////////////////////////////////////////////////////////
/// \class TreeViewItemContainer
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c TreeViewItem objects</em>
///
/// This class provides easy access to collections of \c TreeViewItem objects. While a
/// \c TreeViewItems object is used to group items that have certain properties in common, a
/// \c TreeViewItemContainer object is used to group any items and acts more like a clipboard.
///
/// \if UNICODE
///   \sa ExTVwLibU::ITreeViewItemContainer, TreeViewItem, TreeViewItems, ExplorerTreeView
/// \else
///   \sa ExTVwLibA::ITreeViewItemContainer, TreeViewItem, TreeViewItems, ExplorerTreeView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExTVwU.h"
#else
	#include "ExTVwA.h"
#endif
#include "_ITreeViewItemContainerEvents_CP.h"
#include "IItemContainer.h"
#include "helpers.h"
#include "ExplorerTreeView.h"
#include "TreeViewItem.h"


class ATL_NO_VTABLE TreeViewItemContainer : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TreeViewItemContainer, &CLSID_TreeViewItemContainer>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TreeViewItemContainer>,
    public Proxy_ITreeViewItemContainerEvents<TreeViewItemContainer>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<ITreeViewItemContainer, &IID_ITreeViewItemContainer, &LIBID_ExTVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<ITreeViewItemContainer, &IID_ITreeViewItemContainer, &LIBID_ExTVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IItemContainer
{
	friend class ExplorerTreeView;
	friend class TreeViewItemContainer;

public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used to generate and set this object's ID.
	///
	/// \sa ~TreeViewItemContainer
	TreeViewItemContainer();
	/// \brief <em>The destructor of this class</em>
	///
	/// Used to deregister the container.
	///
	/// \sa TreeViewItemContainer
	~TreeViewItemContainer();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TREEVIEWITEMCONTAINER)

		BEGIN_COM_MAP(TreeViewItemContainer)
			COM_INTERFACE_ENTRY(ITreeViewItemContainer)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TreeViewItemContainer)
			CONNECTION_POINT_ENTRY(__uuidof(_ITreeViewItemContainerEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportErrorInfo
	///
	//@{
	/// \brief <em>Retrieves whether an interface supports the \c IErrorInfo interface</em>
	///
	/// \param[in] interfaceToCheck The IID of the interface to check.
	///
	/// \return \c S_OK if the interface identified by \c interfaceToCheck supports \c IErrorInfo;
	///         otherwise \c S_FALSE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221233.aspx">IErrorInfo</a>
	virtual HRESULT STDMETHODCALLTYPE InterfaceSupportsErrorInfo(REFIID interfaceToCheck);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IEnumVARIANT
	///
	//@{
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the items</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c TreeViewItem objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x items</em>
	///
	/// Retrieves the next \c numberOfMaxItems items from the iterator.
	///
	/// \param[in] numberOfMaxItems The maximum number of items the array identified by \c pItems can
	///            contain.
	/// \param[in,out] pItems An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to an item's \c ITreeViewItem implementation.
	/// \param[out] pNumberOfItemsReturned The number of items that actually were copied to the array
	///             identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, TreeViewItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// item in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x items</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfItemsToSkip items.
	///
	/// \param[in] numberOfItemsToSkip The number of items to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfItemsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ITreeViewItemContainer
	///
	//@{
	/// \brief <em>Retrieves a \c TreeViewItem object from the collection</em>
	///
	/// Retrieves a \c TreeViewItem object from the collection that wraps the treeview item identified
	/// by \c itemID.
	///
	/// \param[in] itemID The unique ID of the treeview item to retrieve.
	/// \param[out] ppItem Receives the item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TreeViewItem::get_ID, Add, Remove
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG itemID, ITreeViewItem** ppItem);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c TreeViewItem objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator Receives the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);

	/// \brief <em>Adds the specified item(s) to the collection</em>
	///
	/// \param[in] items The item(s) to add. May be either an item ID, a \c TreeViewItem object or a
	///            \c TreeViewItems collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa TreeViewItem::get_ID, Count, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Add(VARIANT items);
	/// \brief <em>Clones the collection object</em>
	///
	/// Retrieves an exact copy of the collection.
	///
	/// \param[out] ppClone Receives the clone's \c ITreeViewItemContainer implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ExplorerTreeView::CreateItemContainer
	virtual HRESULT STDMETHODCALLTYPE Clone(ITreeViewItemContainer** ppClone);
	/// \brief <em>Counts the items in the collection</em>
	///
	/// Retrieves the number of \c TreeViewItem objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Retrieves an imagelist containing the items' common drag image</em>
	///
	/// Retrieves the handle to an imagelist containing a bitmap that can be used to visualize
	/// dragging of the items of this collection.
	///
	/// \param[out] pXUpperLeft The x-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] pYUpperLeft The y-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] phImageList The handle to the imagelist.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The caller is responsible for destroying the imagelist.
	///
	/// \sa ExplorerTreeView::CreateLegacyDragImage
	virtual HRESULT STDMETHODCALLTYPE CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft = NULL, OLE_YPOS_PIXELS* pYUpperLeft = NULL, OLE_HANDLE* phImageList = NULL);
	/// \brief <em>Removes the specified item from the collection</em>
	///
	/// \param[in] itemID The unique ID of the treeview item to remove.
	/// \param[in] removePhysically If \c VARIANT_TRUE, the item is removed from the control, too.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa TreeViewItem::get_ID, Add, Count, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG itemID, VARIANT_BOOL removePhysically = VARIANT_FALSE);
	/// \brief <em>Removes all items from the collection</em>
	///
	/// \param[in] removePhysically If \c VARIANT_TRUE, the items get removed from the control, too.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(VARIANT_BOOL removePhysically = VARIANT_FALSE);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Sets the owner of this collection</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerExTvw
	void SetOwner(__in_opt ExplorerTreeView* pOwner);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IItemContainer
	///
	//@{
	/// \brief <em>Adds a large number of items to the container</em>
	///
	/// \param[in] numberOfItems The number of items to add.
	/// \param[in] pItemIDs An array containing the unique IDs of the items to add.
	///
	/// \remarks This method does not check whether any of the specified items already is in the container.
	///
	/// \sa RemovedItem, ExplorerTreeView::OnInternalCreateItemContainer
	void AddItems(UINT numberOfItems, __in_ecount(numberOfItems) PLONG pItemIDs);
	/// \brief <em>An item was removed and the item container should check its content</em>
	///
	/// \param[in] itemID The unique ID of the removed item. If 0, all items were removed.
	///
	/// \sa AddItems, ExplorerTreeView::RemoveItemFromItemContainers
	void RemovedItem(LONG itemID);
	/// \brief <em>Retrieves the container's ID used to identify it</em>
	///
	/// \return The container's ID.
	///
	/// \sa ExplorerTreeView::DeregisterItemContainer, containerID
	DWORD GetID(void);
	/// \brief <em>Holds the container's ID</em>
	DWORD containerID;
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ExplorerTreeView object that owns this collection</em>
		///
		/// \sa SetOwner
		ExplorerTreeView* pOwnerExTvw;
		#ifdef USE_STL
			/// \brief <em>Holds the items that this object contains</em>
			std::vector<LONG> items;
		#else
			/// \brief <em>Holds the items that this object contains</em>
			//CAtlArray<LONG> items;
		#endif
		/// \brief <em>Points to the next enumerated item</em>
		int nextEnumeratedItem;

		Properties()
		{
			pOwnerExTvw = NULL;
			nextEnumeratedItem = 0;
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);

		/// \brief <em>Retrieves the owning treeview's window handle</em>
		///
		/// \return The window handle of the treeview that contains the items in this collection.
		///
		/// \sa pOwnerExTvw
		HWND GetExTvwHWnd(void);
	} properties;

	/* TODO: If we move this one into the Properties struct, the compiler complains that the private
	         = operator of CAtlArray cannot be accessed?! */
	#ifndef USE_STL
		/// \brief <em>Holds the items that this object contains</em>
		CAtlArray<LONG> items;
	#endif

	/// \brief <em>Incremented by the constructor and used as the constructed object's ID</em>
	///
	/// \sa TreeViewItemContainer, containerID
	static DWORD nextID;
};     // TreeViewItemContainer

OBJECT_ENTRY_AUTO(__uuidof(TreeViewItemContainer), TreeViewItemContainer)