//////////////////////////////////////////////////////////////////////
/// \class VirtualTreeViewItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a not existing treeview item</em>
///
/// This class is a wrapper around a treeview item that does not yet or not anymore exist in the
/// control.
///
/// \if UNICODE
///   \sa ExTVwLibU::IVirtualTreeViewItem, TreeViewItem, ExplorerTreeView
/// \else
///   \sa ExTVwLibA::IVirtualTreeViewItem, TreeViewItem, ExplorerTreeView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExTVwU.h"
#else
	#include "ExTVwA.h"
#endif
#include "_IVirtualTreeViewItemEvents_CP.h"
#include "helpers.h"
#include "ExplorerTreeView.h"


class ATL_NO_VTABLE VirtualTreeViewItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualTreeViewItem, &CLSID_VirtualTreeViewItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualTreeViewItem>,
    public Proxy_IVirtualTreeViewItemEvents<VirtualTreeViewItem>,
    #ifdef UNICODE
    	public IDispatchImpl<IVirtualTreeViewItem, &IID_IVirtualTreeViewItem, &LIBID_ExTVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualTreeViewItem, &IID_IVirtualTreeViewItem, &LIBID_ExTVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerTreeView;
	friend class TreeViewItem;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALTREEVIEWITEM)

		BEGIN_COM_MAP(VirtualTreeViewItem)
			COM_INTERFACE_ENTRY(IVirtualTreeViewItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualTreeViewItem)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualTreeViewItemEvents))
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
	/// \name Implementation of IVirtualTreeViewItem
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Bold property</em>
	///
	/// Retrieves whether the item's text will be or was drawn in bold font. If this property is set
	/// to \c VARIANT_TRUE, the item's text will be or was drawn in bold font; otherwise it will be
	/// or was drawn in normal font.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Text
	virtual HRESULT STDMETHODCALLTYPE get_Bold(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c DropHilited property</em>
	///
	/// Retrieves whether the item will be or was drawn as the target of a drag'n'drop operation, i. e.
	/// whether its background will be or was highlighted. If this property is set to \c VARIANT_TRUE,
	/// the item will be or was highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ExplorerTreeView::get_DropHilitedItem, get_Selected
	virtual HRESULT STDMETHODCALLTYPE get_DropHilited(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c ExpandedIconIndex property</em>
	///
	/// Retrieves the zero-based index of the item's expanded icon in the control's \c ilItems imagelist. The
	/// expanded icon is used instead of the normal icon identified by the \c IconIndex property if the item
	/// becomes or was expanded.\n
	/// If set to -1, the control will fire the \c ItemGetDisplayInfo event each time this property's
	/// value is required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExplorerTreeView::get_hImageList, ExplorerTreeView::Raise_ItemGetDisplayInfo, get_IconIndex,
	///       get_SelectedIconIndex, get_OverlayIndex, get_StateImageIndex, get_ExpansionState,
	///       ExTVwLibU::ImageListConstants
	/// \else
	///   \sa ExplorerTreeView::get_hImageList, ExplorerTreeView::Raise_ItemGetDisplayInfo, get_IconIndex,
	///       get_SelectedIconIndex, get_OverlayIndex, get_StateImageIndex, get_ExpansionState,
	///       ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ExpandedIconIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ExpansionState property</em>
	///
	/// Retrieves whether (and how) the item will be or was expanded. Any of the values defined by the
	/// \c ExpansionStateConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_ExpandedIconIndex, ExTVwLibU::ExpansionStateConstants
	/// \else
	///   \sa get_ExpandedIconIndex, ExTVwLibA::ExpansionStateConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ExpansionState(ExpansionStateConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c Ghosted property</em>
	///
	/// Retrieves whether the item's icon will be or was drawn semi-transparent. If this property is
	/// set to \c VARIANT_TRUE, the item's icon will be or was drawn semi-transparent; otherwise it
	/// will be or was drawn normal. Usually you make items ghosted if they're hidden or selected for a
	/// cut-paste-operation.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Grayed, get_Virtual, get_IconIndex
	virtual HRESULT STDMETHODCALLTYPE get_Ghosted(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Grayed property</em>
	///
	/// Retrieves whether the item will be or was drawn as a disabled item, i. e. whether its background
	/// color will be or was changed to gray. If this property is set to \c VARIANT_TRUE, the item's
	/// background will be or was gray; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_Ghosted, ExplorerTreeView::get_Enabled
	virtual HRESULT STDMETHODCALLTYPE get_Grayed(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c HasExpando property</em>
	///
	/// Retrieves whether a "+" or "-" will be or was drawn next to the item indicating that the
	/// item will have or had sub-items. Any of the values defined by the \c HasExpandoConstants
	/// enumeration is valid. If set to \c heCallback, the control will fire the \c ItemGetDisplayInfo
	/// event each time this property's value is required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::HasExpandoConstants, ExplorerTreeView::Raise_ItemGetDisplayInfo,
	///       ExplorerTreeView::get_TreeViewStyle, ExplorerTreeView::get_FadeExpandos
	/// \else
	///   \sa ExTVwLibA::HasExpandoConstants, ExplorerTreeView::Raise_ItemGetDisplayInfo,
	///       ExplorerTreeView::get_TreeViewStyle, ExplorerTreeView::get_FadeExpandos
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_HasExpando(HasExpandoConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c HeightIncrement property</em>
	///
	/// Retrieves the item's height in multiples of control's basic item height. E. g. a value of 2
	/// means that the item will be or was twice as high as an item with \c HeighIncrement set to 1.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ExplorerTreeView::get_ItemHeight
	virtual HRESULT STDMETHODCALLTYPE get_HeightIncrement(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the item's icon in the control's \c ilItems imagelist. If set to
	/// -1, the control will fire the \c ItemGetDisplayInfo event each time this property's value is
	/// required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExplorerTreeView::get_hImageList, ExplorerTreeView::Raise_ItemGetDisplayInfo,
	///       get_SelectedIconIndex, get_ExpandedIconIndex, get_OverlayIndex, get_StateImageIndex,
	///       ExTVwLibU::ImageListConstants
	/// \else
	///   \sa ExplorerTreeView::get_hImageList, ExplorerTreeView::Raise_ItemGetDisplayInfo,
	///       get_SelectedIconIndex, get_ExpandedIconIndex, get_OverlayIndex, get_StateImageIndex,
	///       ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ItemData property</em>
	///
	/// Retrieves the \c LONG value that will be or was associated with the item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ExplorerTreeView::Raise_FreeItemData
	virtual HRESULT STDMETHODCALLTYPE get_ItemData(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Level property</em>
	///
	/// Retrieves the item's zero-based indentation level. An item's indentation level is its parent
	/// item's indentation level incremented by 1. A value of 0 means that the item doesn't have any
	/// parent item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_ParentItem, ExplorerTreeView::get_Indent
	virtual HRESULT STDMETHODCALLTYPE get_Level(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c NextSiblingItem property</em>
	///
	/// Retrieves the item that will be or was this item's proceeding item with the same immediate
	/// parent item.
	///
	/// \param[out] ppNextItem Receives the next item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TreeViewItem, get_PreviousSiblingItem, get_ParentItem
	virtual HRESULT STDMETHODCALLTYPE get_NextSiblingItem(ITreeViewItem** ppNextItem);
	/// \brief <em>Retrieves the current setting of the \c OverlayIndex property</em>
	///
	/// Retrieves the one-based index of the item's overlay icon in the control's \c ilItems imagelist. An
	/// index of 0 means that no overlay will be or was drawn for this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExplorerTreeView::get_hImageList, get_IconIndex, get_SelectedIconIndex, get_ExpandedIconIndex,
	///       get_StateImageIndex, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa ExplorerTreeView::get_hImageList, get_IconIndex, get_SelectedIconIndex, get_ExpandedIconIndex,
	///       get_StateImageIndex, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OverlayIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ParentItem property</em>
	///
	/// Retrieves the item that will be or was this item's immediate parent item.
	///
	/// \param[out] ppParentItem Receives the parent item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TreeViewItem
	virtual HRESULT STDMETHODCALLTYPE get_ParentItem(ITreeViewItem** ppParentItem);
	/// \brief <em>Retrieves the current setting of the \c PreviousSiblingItem property</em>
	///
	/// Retrieves the item that will be or was this item's preceding item with the same immediate parent
	/// item.
	///
	/// \param[out] ppPreviousItem Receives the previous item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TreeViewItem, get_NextSiblingItem, get_ParentItem
	virtual HRESULT STDMETHODCALLTYPE get_PreviousSiblingItem(ITreeViewItem** ppPreviousItem);
	/// \brief <em>Retrieves the current setting of the \c Selected property</em>
	///
	/// Retrieves whether the item will be or was drawn as a selected item, i. e. whether its background
	/// will be or was highlighted. If this property is set to \c VARIANT_TRUE, the item will be or was
	/// highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_DropHilited
	virtual HRESULT STDMETHODCALLTYPE get_Selected(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c SelectedIconIndex property</em>
	///
	/// Retrieves the zero-based index of the item's selected icon in the control's \c ilItems imagelist. The
	/// selected icon is used instead of the normal icon identified by the \c IconIndex property if the item
	/// becomes or was the caret item.\n
	/// If set to -1, the control will fire the \c ItemGetDisplayInfo event each time this property's
	/// value is required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExplorerTreeView::get_hImageList, ExplorerTreeView::Raise_ItemGetDisplayInfo, get_IconIndex,
	///       get_ExpandedIconIndex, get_OverlayIndex, get_StateImageIndex, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa ExplorerTreeView::get_hImageList, ExplorerTreeView::Raise_ItemGetDisplayInfo, get_IconIndex,
	///       get_ExpandedIconIndex, get_OverlayIndex, get_StateImageIndex, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SelectedIconIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c StateImageIndex property</em>
	///
	/// Retrieves the one-based index of the item's state image in the control's \c ilState imagelist. The
	/// state image is drawn next to the item and usually a checkbox.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExplorerTreeView::get_hImageList, get_IconIndex, get_SelectedIconIndex,
	///       get_ExpandedIconIndex, get_OverlayIndex, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa ExplorerTreeView::get_hImageList, get_IconIndex, get_SelectedIconIndex,
	///       get_ExpandedIconIndex, get_OverlayIndex, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_StateImageIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the item's text. The maximum number of characters in this text is specified by
	/// \c MAX_ITEMTEXTLENGTH. If set to \c NULL, the control will fire the \c ItemGetDisplayInfo event
	/// each time this property's value is required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa MAX_ITEMTEXTLENGTH, GetPath, ExplorerTreeView::Raise_ItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c Virtual property</em>
	///
	/// Retrieves whether the item will be or was treated as not existent when drawing the treeview. If this
	/// property is set to \c VARIANT_TRUE, the item won't be or was not drawn. Instead its sub-items will be
	/// or were drawn at this item's position. If this property is set to \c VARIANT_FALSE, the item and its
	/// sub-items will be or were drawn normally.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_Ghosted
	virtual HRESULT STDMETHODCALLTYPE get_Virtual(VARIANT_BOOL* pValue);

	/// \brief <em>Retrieves the item's path</em>
	///
	/// Retrieves a string of the form ...\\a\\b\\c where c is this item's text, b is its parent item's
	/// text, c is b's parent item's text and so on. The separator is definable and defaults to '\\'.
	///
	/// \param[in] pSeparator The separator to use.
	/// \param[out] pPath The item's path.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text
	virtual HRESULT STDMETHODCALLTYPE GetPath(BSTR* pSeparator, BSTR* pPath);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Retrieves this item's parent item</em>
	///
	/// \return This item's parent item that it will have after insertion into or that it had before
	///         deletion from the control.
	HTREEITEM GetParentItem(void);
	/// \brief <em>Retrieves this item's preceding item</em>
	///
	/// \return This item's preceding item that it will have after insertion into or that it had
	///         before deletion from the control.
	HTREEITEM GetPreviousItem(void);
	/// \brief <em>Initializes this object with given data</em>
	///
	/// Initializes this object with the settings from a given source.
	///
	/// \param[in] pSource The data source from which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(__in LPTVINSERTSTRUCT pSource);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(__in LPTVINSERTSTRUCT pTarget);
	/// \brief <em>Sets the owner of this item</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerExTvw
	void SetOwner(__in_opt ExplorerTreeView* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ExplorerTreeView object that owns this item</em>
		///
		/// \sa SetOwner
		ExplorerTreeView* pOwnerExTvw;
		/// \brief <em>A structure holding this item's settings</em>
		TVINSERTSTRUCT settings;

		Properties()
		{
			pOwnerExTvw = NULL;
			ZeroMemory(&settings, sizeof(settings));
		}

		~Properties();

		/// \brief <em>Retrieves the owning treeview's window handle</em>
		///
		/// \return The window handle of the treeview that will contain or has contained this item.
		///
		/// \sa pOwnerExTvw
		HWND GetExTvwHWnd(void);
	} properties;
};     // VirtualTreeViewItem

OBJECT_ENTRY_AUTO(__uuidof(VirtualTreeViewItem), VirtualTreeViewItem)