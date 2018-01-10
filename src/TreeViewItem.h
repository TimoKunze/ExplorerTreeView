//////////////////////////////////////////////////////////////////////
/// \class TreeViewItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing treeview item</em>
///
/// This class is a wrapper around a treeview item that - unlike an item wrapped by
/// \c VirtualTreeViewItem - really exists in the control.
///
/// \todo Write a helper class for moving items.
///
/// \if UNICODE
///   \sa ExTVwLibU::ITreeViewItem, VirtualTreeViewItem, TreeViewItems, TreeViewItemContainer,
///       ExplorerTreeView
/// \else
///   \sa ExTVwLibA::ITreeViewItem, VirtualTreeViewItem, TreeViewItems, TreeViewItemContainer,
///       ExplorerTreeView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExTVwU.h"
#else
	#include "ExTVwA.h"
#endif
#include "CWindowEx.h"
#include "_ITreeViewItemEvents_CP.h"
#include "helpers.h"
#include "ExplorerTreeView.h"


class ATL_NO_VTABLE TreeViewItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TreeViewItem, &CLSID_TreeViewItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TreeViewItem>,
    public Proxy_ITreeViewItemEvents<TreeViewItem>, 
    #ifdef UNICODE
    	public IDispatchImpl<ITreeViewItem, &IID_ITreeViewItem, &LIBID_ExTVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<ITreeViewItem, &IID_ITreeViewItem, &LIBID_ExTVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerTreeView;
	friend class TreeViewItems;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TREEVIEWITEM)

		BEGIN_COM_MAP(TreeViewItem)
			COM_INTERFACE_ENTRY(ITreeViewItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TreeViewItem)
			CONNECTION_POINT_ENTRY(__uuidof(_ITreeViewItemEvents))
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
	/// \name Implementation of ITreeViewItem
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AccessibilityID property</em>
	///
	/// Retrieves the item's accessibility ID which is used to identify the item in an accessibility
	/// explorer.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Handle, get_ID, ExTVwLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa get_Handle, get_ID, ExTVwLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_AccessibilityID(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Bold property</em>
	///
	/// Retrieves whether the item's text is drawn in bold font. If this property is set to
	/// \c VARIANT_TRUE, the item's text is drawn in bold font; otherwise it's drawn in normal font.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Bold, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_Bold(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Bold property</em>
	///
	/// Sets whether the item's text is drawn in bold font. If this property is set to \c VARIANT_TRUE,
	/// the item's text is drawn in bold font; otherwise it's drawn in normal font.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Bold, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_Bold(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Caret property</em>
	///
	/// Retrieves whether the item is the control's caret item, i. e. it has the focus. If it is the
	/// caret item, this property is set to \c VARIANT_TRUE; otherwise it's set to \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Selected, ExplorerTreeView::get_CaretItem
	virtual HRESULT STDMETHODCALLTYPE get_Caret(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c DropHilited property</em>
	///
	/// Retrieves whether the item is drawn as the target of a drag'n'drop operation, i. e. whether its
	/// background is highlighted. If this property is set to \c VARIANT_TRUE, the item is highlighted;
	/// otherwise not.
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
	/// is expanded.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa put_ExpandedIconIndex, ExplorerTreeView::get_hImageList, get_IconIndex, get_SelectedIconIndex,
	///       get_OverlayIndex, get_StateImageIndex, get_ExpansionState, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa put_ExpandedIconIndex, ExplorerTreeView::get_hImageList, get_IconIndex, get_SelectedIconIndex,
	///       get_OverlayIndex, get_StateImageIndex, get_ExpansionState, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ExpandedIconIndex(LONG* pValue);
	/// \brief <em>Sets the \c ExpandedIconIndex property</em>
	///
	/// Sets the zero-based index of the item's expanded icon in the control's \c ilItems imagelist. The
	/// expanded icon is used instead of the normal icon identified by the \c IconIndex property if the item
	/// is expanded. If set to -1, the control will fire the \c ItemGetDisplayInfo event each time this
	/// property's value is needed.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa get_ExpandedIconIndex, ExplorerTreeView::put_hImageList,
	///       ExplorerTreeView::Raise_ItemGetDisplayInfo, put_IconIndex, put_SelectedIconIndex,
	///       put_OverlayIndex, put_StateImageIndex, put_ExpansionState, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa get_ExpandedIconIndex, ExplorerTreeView::put_hImageList,
	///       ExplorerTreeView::Raise_ItemGetDisplayInfo, put_IconIndex, put_SelectedIconIndex,
	///       put_OverlayIndex, put_StateImageIndex, put_ExpansionState, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ExpandedIconIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ExpansionState property</em>
	///
	/// Retrieves whether (and how) the item is expanded. Any of the values defined by the
	/// \c ExpansionStateConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ExpansionState, get_ExpandedIconIndex, ExTVwLibU::ExpansionStateConstants
	/// \else
	///   \sa put_ExpansionState, get_ExpandedIconIndex, ExTVwLibA::ExpansionStateConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ExpansionState(ExpansionStateConstants* pValue);
	/// \brief <em>Sets the \c ExpansionState property</em>
	///
	/// Sets whether (and how) the item is expanded. Any of the values defined by the
	/// \c ExpansionStateConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ExpansionState, put_ExpandedIconIndex, ExTVwLibU::ExpansionStateConstants
	/// \else
	///   \sa get_ExpansionState, put_ExpandedIconIndex, ExTVwLibA::ExpansionStateConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ExpansionState(ExpansionStateConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c FirstSubItem property</em>
	///
	/// Retrieves the item that is this item's first immediate sub-item.
	///
	/// \param[out] ppFirstSubItem Receives the sub-item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_LastSubItem, get_SubItems
	virtual HRESULT STDMETHODCALLTYPE get_FirstSubItem(ITreeViewItem** ppFirstSubItem);
	/// \brief <em>Retrieves the current setting of the \c Ghosted property</em>
	///
	/// Retrieves whether the item's icon is drawn semi-transparent. If this property is set to
	/// \c VARIANT_TRUE, the item's icon is drawn semi-transparent; otherwise it's drawn normal.
	/// Usually you make items ghosted if they're hidden or selected for a cut-paste-operation.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Ghosted, get_Virtual, get_Grayed, get_IconIndex
	virtual HRESULT STDMETHODCALLTYPE get_Ghosted(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Ghosted property</em>
	///
	/// Sets whether the item's icon is drawn semi-transparent. If this property is set to
	/// \c VARIANT_TRUE, the item's icon is drawn semi-transparent; otherwise it's drawn normal.
	/// Usually you make items ghosted if they're hidden or selected for a cut-paste-operation.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Ghosted, get_Virtual, put_Grayed, put_IconIndex
	virtual HRESULT STDMETHODCALLTYPE put_Ghosted(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Grayed property</em>
	///
	/// Retrieves whether the item is drawn as a disabled item, i. e. whether its background color is changed
	/// to gray. If this property is set to \c VARIANT_TRUE, the item's background is gray; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_Grayed, get_Ghosted, ExplorerTreeView::get_Enabled
	virtual HRESULT STDMETHODCALLTYPE get_Grayed(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Grayed property</em>
	///
	/// Sets whether the item is drawn as a disabled item, i. e. whether its background color is changed
	/// to gray. If this property is set to \c VARIANT_TRUE, the item's background is gray; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_Grayed, put_Ghosted, ExplorerTreeView::put_Enabled
	virtual HRESULT STDMETHODCALLTYPE put_Grayed(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Handle property</em>
	///
	/// Retrieves a handle identifying this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Although moving an item changes its handle, the handle is the best (and fastest) option
	///          to identify an item.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_AccessibilityID, get_ID, ExTVwLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa get_AccessibilityID, get_ID, ExTVwLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Handle(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c HasExpando property</em>
	///
	/// Retrieves whether a "+" or "-" is drawn next to the item indicating that the item has sub-items.
	/// Any of the values defined by the \c HasExpandoConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_HasExpando, ExTVwLibU::HasExpandoConstants, get_SubItems,
	///       ExplorerTreeView::get_TreeViewStyle, ExplorerTreeView::get_FadeExpandos
	/// \else
	///   \sa put_HasExpando, ExTVwLibA::HasExpandoConstants, get_SubItems
	///       ExplorerTreeView::get_TreeViewStyle, ExplorerTreeView::get_FadeExpandos
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_HasExpando(HasExpandoConstants* pValue);
	/// \brief <em>Sets the \c HasExpando property</em>
	///
	/// Sets whether a "+" or "-" is drawn next to the item indicating that the item has sub-items.
	/// Any of the values defined by the \c HasExpandoConstants enumeration is valid. If set to
	/// \c heCallback, the control will fire the \c ItemGetDisplayInfo event each time this property's
	/// value is required.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_HasExpando, ExTVwLibU::HasExpandoConstants,
	///       ExplorerTreeView::Raise_ItemGetDisplayInfo, get_SubItems,
	///       ExplorerTreeView::put_TreeViewStyle, ExplorerTreeView::put_FadeExpandos
	/// \else
	///   \sa get_HasExpando, ExTVwLibA::HasExpandoConstants,
	///       ExplorerTreeView::Raise_ItemGetDisplayInfo, get_SubItems,
	///       ExplorerTreeView::put_TreeViewStyle, ExplorerTreeView::put_FadeExpandos
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_HasExpando(HasExpandoConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c HeightIncrement property</em>
	///
	/// Retrieves the item's height in multiples of the control's basic item height. E. g. a
	/// value of 2 means that the item is twice as high as an item with \c HeighIncrement set to 1.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HeightIncrement, ExplorerTreeView::get_ItemHeight
	virtual HRESULT STDMETHODCALLTYPE get_HeightIncrement(LONG* pValue);
	/// \brief <em>Sets the \c HeightIncrement property</em>
	///
	/// Sets the item's height in multiples of the control's basic item height. E. g. a value of
	/// 2 means that the item is twice as high as an item with \c HeighIncrement set to 1.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HeightIncrement, ExplorerTreeView::put_ItemHeight
	virtual HRESULT STDMETHODCALLTYPE put_HeightIncrement(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the item's icon in the control's \c ilItems imagelist.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_IconIndex, ExplorerTreeView::get_hImageList, get_SelectedIconIndex, get_ExpandedIconIndex,
	///       get_OverlayIndex, get_StateImageIndex, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa put_IconIndex, ExplorerTreeView::get_hImageList, get_SelectedIconIndex, get_ExpandedIconIndex,
	///       get_OverlayIndex, get_StateImageIndex, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
	/// \brief <em>Sets the \c IconIndex property</em>
	///
	/// Sets the zero-based index of the item's icon in the control's \c ilItems imagelist. If set to -1, the
	/// control will fire the \c ItemGetDisplayInfo event each time this property's value is required.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IconIndex, ExplorerTreeView::put_hImageList, ExplorerTreeView::Raise_ItemGetDisplayInfo,
	///       put_SelectedIconIndex, put_ExpandedIconIndex, put_OverlayIndex, put_StateImageIndex,
	///       ExTVwLibU::ImageListConstants
	/// \else
	///   \sa get_IconIndex, ExplorerTreeView::put_hImageList, ExplorerTreeView::Raise_ItemGetDisplayInfo,
	///       put_SelectedIconIndex, put_ExpandedIconIndex, put_OverlayIndex, put_StateImageIndex,
	///       ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IconIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ID property</em>
	///
	/// Retrieves an unique ID identifying this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks An item's ID never changes.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_AccessibilityID, get_Handle, ExTVwLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa get_AccessibilityID, get_Handle, ExTVwLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ID(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ItemData property</em>
	///
	/// Retrieves the \c LONG value associated with the item. Use this property to associate any data
	/// with the item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ItemData, ExplorerTreeView::Raise_FreeItemData
	virtual HRESULT STDMETHODCALLTYPE get_ItemData(LONG* pValue);
	/// \brief <em>Sets the \c ItemData property</em>
	///
	/// Sets the \c LONG value associated with the item. Use this property to associate any data
	/// with the item.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ItemData, ExplorerTreeView::Raise_FreeItemData
	virtual HRESULT STDMETHODCALLTYPE put_ItemData(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c LastSubItem property</em>
	///
	/// Retrieves the item that is this item's last immediate sub-item.
	///
	/// \param[out] ppLastSubItem Receives the sub-item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_FirstSubItem, get_SubItems
	virtual HRESULT STDMETHODCALLTYPE get_LastSubItem(ITreeViewItem** ppLastSubItem);
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
	/// Retrieves the item being this item's proceeding item with the same immediate parent item.
	///
	/// \param[out] ppNextItem Receives the next item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_PreviousSiblingItem, get_NextVisibleItem, get_ParentItem
	virtual HRESULT STDMETHODCALLTYPE get_NextSiblingItem(ITreeViewItem** ppNextItem);
	/// \brief <em>Retrieves the current setting of the \c NextVisibleItem property</em>
	///
	/// Retrieves the item being this item's proceeding item. The immediate parent item is not
	/// necessarily the same.
	///
	/// \param[out] ppNextItem Receives the next item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_PreviousVisibleItem, get_NextSiblingItem
	virtual HRESULT STDMETHODCALLTYPE get_NextVisibleItem(ITreeViewItem** ppNextItem);
	/// \brief <em>Retrieves the current setting of the \c OverlayIndex property</em>
	///
	/// Retrieves the one-based index of the item's overlay icon in the control's \c ilItems imagelist. An
	/// index of 0 means that no overlay is drawn for this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_OverlayIndex, ExplorerTreeView::get_hImageList, get_IconIndex, get_SelectedIconIndex,
	///       get_ExpandedIconIndex, get_StateImageIndex, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa put_OverlayIndex, ExplorerTreeView::get_hImageList, get_IconIndex, get_SelectedIconIndex,
	///       get_ExpandedIconIndex, get_StateImageIndex, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OverlayIndex(LONG* pValue);
	/// \brief <em>Sets the \c OverlayIndex property</em>
	///
	/// Sets the one-based index of the item's overlay icon in the control's \c ilItems imagelist. An index
	/// of 0 means that no overlay is drawn for this item.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_OverlayIndex, ExplorerTreeView::put_hImageList, put_IconIndex, put_SelectedIconIndex,
	///       put_ExpandedIconIndex, put_StateImageIndex, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa get_OverlayIndex, ExplorerTreeView::put_hImageList, put_IconIndex, put_SelectedIconIndex,
	///       put_ExpandedIconIndex, put_StateImageIndex, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_OverlayIndex(LONG newValue);
	// \brief <em>Retrieves the current setting of the \c Padding property</em>
	//
	// Retrieves the distance (in pixels) between the item's content and the control's border.
	//
	// \param[out] pValue The property's current setting.
	//
	// \return An \c HRESULT error code.
	//
	// \remarks Requires comctl32.dll version 6.10 or higher.
	//
	// \sa put_Padding, ExplorerTreeView::get_ItemXBorder, ExplorerTreeView::get_ItemYBorder
	//TODO: virtual HRESULT STDMETHODCALLTYPE get_Padding(LONG* pValue);
	// \brief <em>Sets the \c Padding property</em>
	//
	// Sets the distance (in pixels) between the item's content and the control's border.
	//
	// \param[in] newValue The setting to apply.
	//
	// \return An \c HRESULT error code.
	//
	// \remarks Requires comctl32.dll version 6.10 or higher.
	//
	// \sa get_Padding, ExplorerTreeView::put_ItemXBorder, ExplorerTreeView::put_ItemYBorder
	//TODO: virtual HRESULT STDMETHODCALLTYPE put_Padding(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ParentItem property</em>
	///
	/// Retrieves the item being this item's immediate parent item.
	///
	/// \param[out] ppParentItem Receives the parent item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa Move
	virtual HRESULT STDMETHODCALLTYPE get_ParentItem(ITreeViewItem** ppParentItem);
	/// \brief <em>Retrieves the current setting of the \c PreviousSiblingItem property</em>
	///
	/// Retrieves the item being this item's preceding item with the same immediate parent item.
	///
	/// \param[out] ppPreviousItem Receives the previous item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_PreviousSiblingItem, get_NextSiblingItem, get_PreviousVisibleItem, get_ParentItem
	virtual HRESULT STDMETHODCALLTYPE get_PreviousSiblingItem(ITreeViewItem** ppPreviousItem);
	/// \brief <em>Sets the \c PreviousSiblingItem property</em>
	///
	/// Sets the item being this item's preceding item with the same parent item. If set to NULL, the
	/// item becomes the first item in its branch.
	///
	/// \param[in] pNewPreviousItem The new previous item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_PreviousSiblingItem, Move
	virtual HRESULT STDMETHODCALLTYPE putref_PreviousSiblingItem(ITreeViewItem* pNewPreviousItem);
	/// \brief <em>Retrieves the current setting of the \c PreviousVisibleItem property</em>
	///
	/// Retrieves the item being this item's preceding item. The immediate parent item is not
	/// necessarily the same.
	///
	/// \param[out] ppPreviousItem Receives the previous item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_NextVisibleItem, get_PreviousSiblingItem
	virtual HRESULT STDMETHODCALLTYPE get_PreviousVisibleItem(ITreeViewItem** ppPreviousItem);
	/// \brief <em>Retrieves the current setting of the \c Selected property</em>
	///
	/// Retrieves whether the item is drawn as a selected item, i. e. whether its background is
	/// highlighted. If this property is set to \c VARIANT_TRUE, the item is highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Caret, get_DropHilited, ExplorerTreeView::get_CaretItem,
	///     ExplorerTreeView::Raise_ItemSelectionChanging, ExplorerTreeView::Raise_ItemSelectionChanged
	virtual HRESULT STDMETHODCALLTYPE get_Selected(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c SelectedIconIndex property</em>
	///
	/// Retrieves the zero-based index of the item's selected icon in the control's \c ilItems imagelist. The
	/// selected icon is used instead of the normal icon identified by the \c IconIndex property if the item
	/// is the caret item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_SelectedIconIndex, ExplorerTreeView::get_hImageList, get_IconIndex, get_ExpandedIconIndex,
	///       get_OverlayIndex, get_StateImageIndex, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa put_SelectedIconIndex, ExplorerTreeView::get_hImageList, get_IconIndex, get_ExpandedIconIndex,
	///       get_OverlayIndex, get_StateImageIndex, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SelectedIconIndex(LONG* pValue);
	/// \brief <em>Sets the \c SelectedIconIndex property</em>
	///
	/// Sets the zero-based index of the item's selected icon in the control's \c ilItems imagelist. The
	/// selected icon is used instead of the normal icon identified by the \c IconIndex property if the item
	/// is the caret item. If set to -1, the control will fire the \c ItemGetDisplayInfo event each time this
	/// property's value is needed.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_SelectedIconIndex, ExplorerTreeView::put_hImageList,
	///       ExplorerTreeView::Raise_ItemGetDisplayInfo, put_IconIndex, put_ExpandedIconIndex,
	///       put_OverlayIndex, put_StateImageIndex, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa get_SelectedIconIndex, ExplorerTreeView::put_hImageList,
	///       ExplorerTreeView::Raise_ItemGetDisplayInfo, put_IconIndex, put_ExpandedIconIndex,
	///       put_OverlayIndex, put_StateImageIndex, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_SelectedIconIndex(LONG newValue);
	#ifdef INCLUDESHELLBROWSERINTERFACE
		/// \brief <em>Retrieves the current setting of the \c ShellTreeViewItemObject property</em>
		///
		/// Retrieves the \c ShellTreeViewItem object of this treeview item from the attached \c ShellTreeView
		/// control.
		///
		/// \param[out] ppItem Receives the object's \c IDispatch implementation.
		///
		/// \return An \c HRESULT error code.
		///
		/// \remarks This property is read-only.
		virtual HRESULT STDMETHODCALLTYPE get_ShellTreeViewItemObject(IDispatch** ppItem);
	#endif
	/// \brief <em>Retrieves the current setting of the \c StateImageIndex property</em>
	///
	/// Retrieves the one-based index of the item's state image in the control's \c ilState imagelist. The
	/// state image is drawn next to the item and usually a checkbox.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_StateImageIndex, ExplorerTreeView::get_hImageList, get_IconIndex, get_SelectedIconIndex,
	///       get_ExpandedIconIndex, get_OverlayIndex, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa put_StateImageIndex, ExplorerTreeView::get_hImageList, get_IconIndex, get_SelectedIconIndex,
	///       get_ExpandedIconIndex, get_OverlayIndex, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_StateImageIndex(LONG* pValue);
	/// \brief <em>Sets the \c StateImageIndex property</em>
	///
	/// Sets the one-based index of the item's state image in the control's \c ilState imagelist. The state
	/// image is drawn next to the item and usually a checkbox.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_StateImageIndex, ExplorerTreeView::put_hImageList, put_IconIndex, put_SelectedIconIndex,
	///       put_ExpandedIconIndex, put_OverlayIndex, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa get_StateImageIndex, ExplorerTreeView::put_hImageList, put_IconIndex, put_SelectedIconIndex,
	///       put_ExpandedIconIndex, put_OverlayIndex, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_StateImageIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c SubItems property</em>
	///
	/// Retrieves a collection object wrapping all immediate sub-items of this item.
	///
	/// \param[out] ppSubItems Receives the collection's \c ITreeViewItems implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TreeViewItems, get_FirstSubItem
	virtual HRESULT STDMETHODCALLTYPE get_SubItems(ITreeViewItems** ppSubItems = NULL);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the item's text. The maximum number of characters in this text is specified by
	/// \c MAX_ITEMTEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text, MAX_ITEMTEXTLENGTH, GetPath, IsTruncated
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the item's text. The maximum number of characters in this text is specified by
	/// \c MAX_ITEMTEXTLENGTH. If set to \c NULL, the control will fire the \c ItemGetDisplayInfo event
	/// each time this property's value is required.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, MAX_ITEMTEXTLENGTH, IsTruncated, ExplorerTreeView::Raise_ItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c Virtual property</em>
	///
	/// Retrieves whether the item is treated as not existent when drawing the treeview. If this property is
	/// set to \c VARIANT_TRUE, the item is not drawn. Instead its sub-items are drawn at this item's
	/// position. If this property is set to \c VARIANT_FALSE, the item and its sub-items are drawn normally.
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

	/// \brief <em>Collapses the item's immediate sub-items</em>
	///
	/// \param[in] removeSubItems Specifies whether to remove all immediate sub-items. If set to
	///            \c VARIANT_FALSE, the items will be collapsed, but not removed; otherwise the items
	///            will be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Expand, get_ExpansionState, get_SubItems
	virtual HRESULT STDMETHODCALLTYPE Collapse(VARIANT_BOOL removeSubItems = VARIANT_FALSE);
	/// \brief <em>Retrieves whether calling \c Move with the specified parameters would succeed</em>
	///
	/// \param[in] pNewParentItem The new parent item's \c ITreeViewItem implementation. If set to
	///            \c NULL, the item would become a root item.
	/// \param[in] insertAfter A \c VARIANT value specifying the item's new preceding item. This may be
	///            either a \c TreeViewItem object or a handle returned by \c TreeViewItem::get_Handle
	///            or a value defined by the \c InsertAfterConstants enumeration. If set to \c Empty,
	///            the item would become the last item in the new branch.
	/// \param[out] pCouldMove If \c VARIANT_TRUE, calling \c Move would succeed; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method does <strong>not</strong> move the item.
	///
	/// \sa Move, get_ParentItem, get_PreviousSiblingItem
	virtual HRESULT STDMETHODCALLTYPE CouldMoveTo(ITreeViewItem* pNewParentItem, VARIANT insertAfter, VARIANT_BOOL* pCouldMove);
	/// \brief <em>Retrieves an imagelist containing the item's drag image</em>
	///
	/// Retrieves the handle to an imagelist containing a bitmap that can be used to visualize
	/// dragging of this item.
	///
	/// \param[out] pXUpperLeft The x-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] pYUpperLeft The y-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] phImageList The imagelist containing the drag image.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The caller is responsible for destroying the imagelist.
	///
	/// \sa ExplorerTreeView::CreateLegacyDragImage
	virtual HRESULT STDMETHODCALLTYPE CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft = NULL, OLE_YPOS_PIXELS* pYUpperLeft = NULL, OLE_HANDLE* phImageList = NULL);
	/// \brief <em>Displays the item's info tip</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method has no effect if the mouse cursor isn't located over the item.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa ExplorerTreeView::get_ShowToolTips, ExplorerTreeView::Raise_ItemGetInfoTipText
	virtual HRESULT STDMETHODCALLTYPE DisplayInfoTip(void);
	/// \brief <em>Ensures the item is visible</em>
	///
	/// Ensures that the item is visible, expanding the parent item or scrolling the control, if necessary.
	///
	/// \param[out] pExpandedItems If \c VARIANT_TRUE, some items got expanded; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Expand
	virtual HRESULT STDMETHODCALLTYPE EnsureVisible(VARIANT_BOOL* pExpandedItems);
	/// \brief <em>Expands the item's immediate sub-items</em>
	///
	/// \param[in] expandPartial Specifies whether to replace the "+" button next to the item with a
	///            "-" button. If set to \c VARIANT_FALSE, the button is changed to "-"; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Collapse, get_ExpansionState, get_SubItems
	virtual HRESULT STDMETHODCALLTYPE Expand(VARIANT_BOOL expandPartial = VARIANT_FALSE);
	/// \brief <em>Retrieves the item's path</em>
	///
	/// Retrieves a string of the form ...\\a\\b\\c where c is this item's text, b is its parent item's
	/// text, a is b's parent item's text and so on. The separator is customizable and defaults to "\".
	///
	/// \param[in] pSeparator The separator to use.
	/// \param[out] pPath The item's path.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, TreeViewItems::FindByPath
	virtual HRESULT STDMETHODCALLTYPE GetPath(BSTR* pSeparator, BSTR* pPath);
	/// \brief <em>Retrieves the bounding rectangle of either the item or a part of it</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's client area) of either the
	/// item or a part of it.
	///
	/// \param[in] rectangleType The rectangle to retrieve. Any of the values defined by the
	///            \c ItemRectangleTypeConstants enumeration is valid.
	/// \param[out] PXLeft The x-coordinate (in pixels) of the bounding rectangle's left border
	///             relative to the control's upper-left corner.
	/// \param[out] pYTop The y-coordinate (in pixels) of the bounding rectangle's top border
	///             relative to the control's upper-left corner.
	/// \param[out] pXRight The x-coordinate (in pixels) of the bounding rectangle's right border
	///             relative to the control's upper-left corner.
	/// \param[out] pYBottom The y-coordinate (in pixels) of the bounding rectangle's bottom border
	///             relative to the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::ItemRectangleTypeConstants
	/// \else
	///   \sa ExTVwLibA::ItemRectangleTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetRectangle(ItemRectangleTypeConstants rectangleType, OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	/// \brief <em>Retrieves whether the item's text is truncated</em>
	///
	/// \param[out] pTruncated If \c VARIANT_TRUE, the item's text is truncated; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text
	virtual HRESULT STDMETHODCALLTYPE IsTruncated(VARIANT_BOOL* pTruncated);
	/// \brief <em>Sets the \c ParentItem and \c PreviousSiblingItem properties</em>
	///
	/// Moves the item to the branch with the specified immediate parent item. It will be inserted after
	/// the specified item.
	///
	/// \param[in] pNewParentItem The new parent item's \c ITreeViewItem implementation. If set to
	///            \c NULL, the item becomes a root item.
	/// \param[in] insertAfter A \c VARIANT value specifying the item's new preceding item. This may be
	///            either a \c TreeViewItem object or a handle returned by \c TreeViewItem::get_Handle
	///            or a value defined by the \c InsertAfterConstants enumeration. If set to \c Empty,
	///            the item becomes the last item in the new branch.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CouldMoveTo, get_ParentItem, get_PreviousSiblingItem, putref_PreviousSiblingItem
	virtual HRESULT STDMETHODCALLTYPE Move(ITreeViewItem* pNewParentItem, VARIANT insertAfter);
	/// \brief <em>Sorts the item's sub-items</em>
	///
	/// Sorts the item's sub-items by up to 5 criteria.
	///
	/// \param[in] firstCriterion The first criterion by which to sort. Any of the values defined by the
	///            \c SortByConstants enumeration is valid.
	/// \param[in] secondCriterion The second criterion by which to sort. It is used if two items are
	///            equivalent regarding the first criterion. Any of the values defined by the
	///            \c SortByConstants enumeration is valid.
	/// \param[in] thirdCriterion The third criterion by which to sort. It is used if two items are
	///            equivalent regarding the first 2 criteria. Any of the values defined by the
	///            \c SortByConstants enumeration is valid.
	/// \param[in] fourthCriterion The fourth criterion by which to sort. It is used if two items are
	///            equivalent regarding the first 3 criteria. Any of the values defined by the
	///            \c SortByConstants enumeration is valid.
	/// \param[in] fifthCriterion The fifth criterion by which to sort. It is used if two items are
	///            equivalent regarding the first 4 criteria. Any of the values defined by the
	///            \c SortByConstants enumeration is valid.
	/// \param[in] recurse If \c VARIANT_FALSE, only the item's immediate sub-items will be sorted;
	///            otherwise all sub-items will be sorted recursively.
	/// \param[in] caseSensitive If \c VARIANT_TRUE, text comparisons will be case sensitive; otherwise
	///            not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExplorerTreeView::get_SortOrder, ExplorerTreeView::SortItems,
	///       ExplorerTreeView::SortSubItems, ExplorerTreeView::Raise_CompareItems,
	///       ExTVwLibU::SortByConstants
	/// \else
	///   \sa ExplorerTreeView::get_SortOrder, ExplorerTreeView::SortItems,
	///       ExplorerTreeView::SortSubItems, ExplorerTreeView::Raise_CompareItems,
	///       ExTVwLibA::SortByConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE SortSubItems(SortByConstants firstCriterion = sobShell, SortByConstants secondCriterion = sobText, SortByConstants thirdCriterion = sobNone, SortByConstants fourthCriterion = sobNone, SortByConstants fifthCriterion = sobNone, VARIANT_BOOL recurse = VARIANT_FALSE, VARIANT_BOOL caseSensitive = VARIANT_FALSE);
	/// \brief <em>Starts editing the item's text</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ExplorerTreeView::EndLabelEdit, ExplorerTreeView::get_hWndEdit,
	///     ExplorerTreeView::Raise_StartingLabelEditing, ExplorerTreeView::Raise_RenamingItem
	virtual HRESULT STDMETHODCALLTYPE StartLabelEditing(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given item</em>
	///
	/// Attaches this object to a given item, so that the item's properties can be retrieved and set
	/// using this object's methods.
	///
	/// \param[in] hItem The item to attach to.
	///
	/// \sa Detach
	void Attach(HTREEITEM hItem);
	/// \brief <em>Detaches this object from an item</em>
	///
	/// Detaches this object from the item it currently wraps, so that it doesn't wrap any item anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// Applies the settings from a given source to the item wrapped by this object.
	///
	/// \param[in] pSource The data source from which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(LPTVINSERTSTRUCT pSource);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// \overload
	HRESULT LoadState(VirtualTreeViewItem* pSource);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndTvw The treeview window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(LPTVINSERTSTRUCT pTarget, HWND hWndTvw = NULL);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \overload
	HRESULT SaveState(VirtualTreeViewItem* pTarget);
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
		/// \brief <em>The handle to the item wrapped by this object</em>
		HTREEITEM itemHandle;

		Properties()
		{
			pOwnerExTvw = NULL;
			itemHandle = NULL;
		}

		~Properties();

		/// \brief <em>Retrieves the owning treeview's window handle</em>
		///
		/// \return The window handle of the treeview that contains this item.
		///
		/// \sa pOwnerExTvw
		HWND GetExTvwHWnd(void);
	} properties;

	/// \brief <em>Holds the object's flags</em>
	struct Flags
	{
		/// \brief <em>Used to speed up item moving</em>
		///
		/// Specifies whether there're other items of the same branch that are moved together with this
		/// item. The mover of this item (usually a \c TreeViewItems object) sets this to \c TRUE to
		/// disable time consuming processings that would be redundant if more than one item is moved.\n
		/// At the moment this flag isn't used.
		///
		/// \sa Move, TreeViewItems
		BOOL movingWholeBranch;
		/// \brief <em>Holds data used to change the caret item</em>
		struct CaretItemChange {
			/// \brief <em>The item that will become the caret item</em>
			///
			/// \sa hPreviousItem, selected
			HTREEITEM hNewItem;
			/// \brief <em>The item being the current caret item</em>
			///
			/// \sa hNewItem, selected
			HTREEITEM hPreviousItem;
		} caretItemChange;

		Flags()
		{
			movingWholeBranch = FALSE;
		}
	} flags;

	/// \brief <em>Retrieves whether calling \c Move with the specified parameters would succeed</em>
	///
	/// \param[in] hItem The item to move.
	/// \param[in] hNewParentItem The new parent item. If set to \c TVI_ROOT, the item would become a
	///            root item.
	/// \param[in] hNewPreviousItem The new previous sibling item. If set to \c TVI_FIRST, the item
	///            would become the first item in the new branch. If set to \c TVI_LAST, the item
	///            would become the last item in the new branch.
	/// \param[in] hWndTvw The treeview window the method will work on.
	///
	/// \return \c TRUE, if calling \c Move with the same parameters would succeed; otherwise \c FALSE.
	///
	/// \remarks This method does <strong>not</strong> move the item.
	///
	/// \sa Move
	BOOL CouldMoveTo(HTREEITEM hItem, HTREEITEM hNewParentItem, HTREEITEM hNewPreviousItem, HWND hWndTvw = NULL);
	/// \brief <em>Retrieves whether calling \c Move with the specified parameters would succeed</em>
	///
	/// \overload
	BOOL CouldMoveTo(HTREEITEM hNewParentItem, HTREEITEM hNewPreviousItem, HWND hWndTvw = NULL);
	/// \brief <em>Retrieves whether a given item is expanded</em>
	///
	/// \param[in] hItem The item to check.
	/// \param[in] allowPartialExpansion If set to \c TRUE, the method will return \c TRUE, if the item
	///            is either partially (\c TVIS_EXPANDPARTIAL) or entirely (\c TVIS_EXPANDED) expanded;
	///            otherwise it <strong>must</strong> be entirely expanded.
	/// \param[in] hWndTvw The treeview window the method will work on.
	///
	/// \return \c TRUE if the item is entirely or partial expanded; otherwise \c FALSE.
	///
	/// \sa get_ExpansionState
	BOOL IsExpanded(HTREEITEM hItem, BOOL allowPartialExpansion = TRUE, HWND hWndTvw = NULL);
	/// \brief <em>Retrieves whether a given item is selected</em>
	///
	/// \param[in] hItem The item to check.
	/// \param[in] hWndTvw The treeview window the method will work on.
	///
	/// \return \c TRUE if the item is selected; otherwise \c FALSE.
	///
	/// \sa get_Selected
	BOOL IsSelected(HTREEITEM hItem, HWND hWndTvw = NULL);
	/// \brief <em>Retrieves whether the item is selected</em>
	///
	/// \overload
	BOOL IsSelected(HWND hWndTvw = NULL);
	/// \brief <em>Retrieves whether the item's text is truncated</em>
	///
	/// \param[in] hWndTvw The treeview window the method will work on.
	///
	/// \return \c TRUE if the item is truncated; otherwise \c FALSE.
	///
	/// \sa get_Text, IsTreeViewItemTruncated
	BOOL IsTruncated(HWND hWndTvw = NULL);
	/// \brief <em>Moves a given item to a new location within the treeview</em>
	///
	/// \param[in] hItem The item to move.
	/// \param[in] hNewParentItem The new parent item. If set to \c TVI_ROOT, the item becomes a root
	///            item.
	/// \param[in] hNewPreviousItem The new previous sibling item. If set to \c TVI_FIRST, the item
	///            becomes the first item in the new branch. If set to \c TVI_LAST, the item becomes
	///            the last item in the new branch.
	/// \param[in] hWndTvw The treeview window the method will work on.
	/// \param[in] skipErrorChecking If set to \c TRUE, the method won't check the \c hNewParentItem
	///            and \c hNewPreviousItem parameters for validity.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method requires some initialization and can't be used alone.
	///
	/// \sa MoveSubItems
	HRESULT Move(HTREEITEM hItem, HTREEITEM hNewParentItem, HTREEITEM hNewPreviousItem, HWND hWndTvw = NULL, BOOL skipErrorChecking = TRUE);
	/// \brief <em>Moves the item to a new location within the treeview</em>
	///
	/// \overload
	HRESULT Move(HTREEITEM hNewParentItem, HTREEITEM hNewPreviousItem, HWND hWndTvw = NULL, BOOL skipErrorChecking = TRUE);
	/// \brief <em>Recursively moves the sub-items of a given item to a new location within the treeview</em>
	///
	/// \param[in] hParentItem The immediate parent item of the items to move.
	/// \param[in] hNewParentItem The new parent item. If set to \c TVI_ROOT, the immediate sub-items
	///            of \c hParentItem become root items.
	/// \param[in] hWndTvw The treeview window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Move
	HRESULT MoveSubItems(HTREEITEM hParentItem, HTREEITEM hNewParentItem, HWND hWndTvw = NULL);
	/// \brief <em>Writes a given object's settings to a given target</em>
	///
	/// \param[in] hItem The item for which to save the settings.
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndTvw The treeview window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(HTREEITEM hItem, LPTVINSERTSTRUCT pTarget, HWND hWndTvw = NULL);
};     // TreeViewItem

OBJECT_ENTRY_AUTO(__uuidof(TreeViewItem), TreeViewItem)