//////////////////////////////////////////////////////////////////////
/// \class Proxy_IExplorerTreeViewEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IExplorerTreeViewEvents interface</em>
///
/// \if UNICODE
///   \sa ExplorerTreeView, ExTVwLibU::_IExplorerTreeViewEvents
/// \else
///   \sa ExplorerTreeView, ExTVwLibA::_IExplorerTreeViewEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IExplorerTreeViewEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IExplorerTreeViewEvents), CComDynamicUnkArray>
{
public:
	/// \brief <em>Fires the \c AbortedDrag event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::AbortedDrag, ExplorerTreeView::Raise_AbortedDrag,
	///       Fire_Drop
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::AbortedDrag, ExplorerTreeView::Raise_AbortedDrag,
	///       Fire_Drop
	/// \endif
	HRESULT Fire_AbortedDrag(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_ABORTEDDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CaretChanged event</em>
	///
	/// \param[in] pPreviousCaretItem The previous caret item.
	/// \param[in] pNewCaretItem The new caret item.
	/// \param[in] caretChangeReason The reason for the caret change. Any of the values defined by
	///            the \c CaretChangeCausedByConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::CaretChanged, ExplorerTreeView::Raise_CaretChanged,
	///       Fire_CaretChanging, Fire_ItemSelectionChanged, Fire_SingleExpanding
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::CaretChanged, ExplorerTreeView::Raise_CaretChanged,
	///       Fire_CaretChanging, Fire_ItemSelectionChanged, Fire_SingleExpanding
	/// \endif
	HRESULT Fire_CaretChanged(ITreeViewItem* pPreviousCaretItem, ITreeViewItem* pNewCaretItem, CaretChangeCausedByConstants caretChangeReason)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pPreviousCaretItem;
				p[1] = pNewCaretItem;
				p[0].lVal = static_cast<LONG>(caretChangeReason);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_CARETCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CaretChanging event</em>
	///
	/// \param[in] pPreviousCaretItem The previous caret item.
	/// \param[in] pNewCaretItem The new caret item.
	/// \param[in] caretChangeReason The reason for the caret change. Any of the values defined by
	///            the \c CaretChangeCausedByConstants enumeration is valid.
	/// \param[in,out] pCancelCaretChange If \c VARIANT_TRUE, the caller should abort the caret change;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::CaretChanging, ExplorerTreeView::Raise_CaretChanging,
	///       Fire_CaretChanged, Fire_ItemSelectionChanging, Fire_SingleExpanding
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::CaretChanging, ExplorerTreeView::Raise_CaretChanging,
	///       Fire_CaretChanged, Fire_ItemSelectionChanging, Fire_SingleExpanding
	/// \endif
	HRESULT Fire_CaretChanging(ITreeViewItem* pPreviousCaretItem, ITreeViewItem* pNewCaretItem, CaretChangeCausedByConstants caretChangeReason, VARIANT_BOOL* pCancelCaretChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pPreviousCaretItem;
				p[2] = pNewCaretItem;
				p[1].lVal = static_cast<LONG>(caretChangeReason);		p[1].vt = VT_I4;
				p[0].pboolVal = pCancelCaretChange;									p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_CARETCHANGING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChangedSortOrder event</em>
	///
	/// \param[in] previousSortOrder The control's old sort order. Any of the values defined by the
	///            \c SortOrderConstants enumeration is valid.
	/// \param[in] newSortOrder The control's new sort order. Any of the values defined by the
	///            \c SortOrderConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ChangedSortOrder,
	///       ExplorerTreeView::Raise_ChangedSortOrder, Fire_ChangingSortOrder
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ChangedSortOrder,
	///       ExplorerTreeView::Raise_ChangedSortOrder, Fire_ChangingSortOrder
	/// \endif
	HRESULT Fire_ChangedSortOrder(SortOrderConstants previousSortOrder, SortOrderConstants newSortOrder)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].lVal = static_cast<LONG>(previousSortOrder);		p[1].vt = VT_I4;
				p[0].lVal = static_cast<LONG>(newSortOrder);				p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_CHANGEDSORTORDER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChangingSortOrder event</em>
	///
	/// \param[in] previousSortOrder The control's old sort order. Any of the values defined by the
	///            \c SortOrderConstants enumeration is valid.
	/// \param[in] newSortOrder The control's new sort order. Any of the values defined by the
	///            \c SortOrderConstants enumeration is valid.
	/// \param[in,out] pCancelChange If \c VARIANT_TRUE, the caller should abort redefining; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ChangingSortOrder,
	///       ExplorerTreeView::Raise_ChangingSortOrder, Fire_ChangedSortOrder
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ChangingSortOrder,
	///       ExplorerTreeView::Raise_ChangingSortOrder, Fire_ChangedSortOrder
	/// \endif
	HRESULT Fire_ChangingSortOrder(SortOrderConstants previousSortOrder, SortOrderConstants newSortOrder, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2].lVal = static_cast<LONG>(previousSortOrder);		p[2].vt = VT_I4;
				p[1].lVal = static_cast<LONG>(newSortOrder);				p[1].vt = VT_I4;
				p[0].pboolVal = pCancelChange;											p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_CHANGINGSORTORDER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Click event</em>
	///
	/// \param[in] pTreeItem The clicked item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::Click, ExplorerTreeView::Raise_Click, Fire_DblClick,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::Click, ExplorerTreeView::Raise_Click, Fire_DblClick,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_Click(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_CLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CollapsedItem event</em>
	///
	/// \param[in] pTreeItem The immediate parent item of the items that got collapsed.
	/// \param[in] removedAllSubItems If \c VARIANT_TRUE, the item's sub-items were removed; otherwise
	///            they just got collapsed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::CollapsedItem, ExplorerTreeView::Raise_CollapsedItem,
	///       Fire_CollapsingItem, Fire_ExpandedItem
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::CollapsedItem, ExplorerTreeView::Raise_CollapsedItem,
	///       Fire_CollapsingItem, Fire_ExpandedItem
	/// \endif
	HRESULT Fire_CollapsedItem(ITreeViewItem* pTreeItem, VARIANT_BOOL removedAllSubItems)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pTreeItem;
				p[0] = removedAllSubItems;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_COLLAPSEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CollapsingItem event</em>
	///
	/// \param[in] pTreeItem The immediate parent item of the items that get collapsed.
	/// \param[in] removingAllSubItems If \c VARIANT_TRUE, the item's sub-items get removed; otherwise
	///            they just get collapsed.
	/// \param[in,out] pCancelCollapse If \c VARIANT_TRUE, the caller should abort the collapse action;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::CollapsingItem, ExplorerTreeView::Raise_CollapsingItem,
	///       Fire_CollapsedItem, Fire_ExpandingItem
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::CollapsingItem, ExplorerTreeView::Raise_CollapsingItem,
	///       Fire_CollapsedItem, Fire_ExpandingItem
	/// \endif
	HRESULT Fire_CollapsingItem(ITreeViewItem* pTreeItem, VARIANT_BOOL removingAllSubItems, VARIANT_BOOL* pCancelCollapse)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pTreeItem;
				p[1] = removingAllSubItems;
				p[0].pboolVal = pCancelCollapse;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_COLLAPSINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CompareItems event</em>
	///
	/// \param[in] pFirstItem The first item to compare.
	/// \param[in] pSecondItem The second item to compare.
	/// \param[out] pResult Receives one of the values defined by the \c CompareResultConstants
	///             enumeration. If \c SortOrder is set to \c soDescending, the caller should invert
	///             this value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::CompareItems, ExplorerTreeView::Raise_CompareItems
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::CompareItems, ExplorerTreeView::Raise_CompareItems
	/// \endif
	HRESULT Fire_CompareItems(ITreeViewItem* pFirstItem, ITreeViewItem* pSecondItem, CompareResultConstants* pResult)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pFirstItem;
				p[1] = pSecondItem;
				p[0].plVal = reinterpret_cast<PLONG>(pResult);		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_COMPAREITEMS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ContextMenu event</em>
	///
	/// \param[in] pTreeItem The item that the context menu refers to. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that the menu's proposed position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pShowDefaultMenu If \c VARIANT_FALSE, the caller should prevent the \c ShellTreeView
	///                control from showing the default context menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ContextMenu, ExplorerTreeView::Raise_ContextMenu,
	///       Fire_RClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ContextMenu, ExplorerTreeView::Raise_ContextMenu,
	///       Fire_RClick
	/// \endif
	HRESULT Fire_ContextMenu(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pShowDefaultMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pTreeItem;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].pboolVal = pShowDefaultMenu;									p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_CONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CreatedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The label-edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::CreatedEditControlWindow,
	///       ExplorerTreeView::Raise_CreatedEditControlWindow, Fire_DestroyedEditControlWindow,
	///       Fire_StartingLabelEditing
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::CreatedEditControlWindow,
	///       ExplorerTreeView::Raise_CreatedEditControlWindow, Fire_DestroyedEditControlWindow,
	///       Fire_StartingLabelEditing
	/// \endif
	HRESULT Fire_CreatedEditControlWindow(LONG hWndEdit)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndEdit;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_CREATEDEDITCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CustomDraw event</em>
	///
	/// \param[in] pTreeItem The item that the notification refers to. May be \c NULL.
	/// \param[in,out] pItemLevel The item's indentation level. The client may change this value.
	/// \param[in,out] pTextColor An \c OLE_COLOR value specifying the color to draw the item's text in.
	///                The client may change this value.
	/// \param[in,out] pTextBackColor An \c OLE_COLOR value specifying the color to fill the item's
	///                text background with. The client may change this value.
	/// \param[in] drawStage The stage of custom drawing this event is raised for. Any of the values
	///            defined by the \c CustomDrawStageConstants enumeration is valid.
	/// \param[in] itemState The item's current state (focused, selected etc.). Most of the values
	///            defined by the \c CustomDrawItemStateConstants enumeration are valid.
	/// \param[in] hDC The handle of the device context in which all drawing shall take place.
	/// \param[in] pDrawingRectangle A \c RECTANGLE structure specifying the bounding rectangle of the
	///            area that needs to be drawn.
	/// \param[in,out] pFurtherProcessing A value controlling further drawing. Most of the values defined
	///                by the \c CustomDrawReturnValuesConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::CustomDraw, ExplorerTreeView::Raise_CustomDraw
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::CustomDraw, ExplorerTreeView::Raise_CustomDraw
	/// \endif
	HRESULT Fire_CustomDraw(ITreeViewItem* pTreeItem, LONG* pItemLevel, OLE_COLOR* pTextColor, OLE_COLOR* pTextBackColor, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants itemState, LONG hDC, RECTANGLE* pDrawingRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pTreeItem;
				p[7].plVal = pItemLevel;																		p[7].vt = VT_I4 | VT_BYREF;
				p[6].plVal = reinterpret_cast<PLONG>(pTextColor);						p[6].vt = VT_I4 | VT_BYREF;
				p[5].plVal = reinterpret_cast<PLONG>(pTextBackColor);				p[5].vt = VT_I4 | VT_BYREF;
				p[4].lVal = static_cast<LONG>(drawStage);										p[4].vt = VT_I4;
				p[3].lVal = static_cast<LONG>(itemState);										p[3].vt = VT_I4;
				p[2] = hDC;
				p[0].plVal = reinterpret_cast<PLONG>(pFurtherProcessing);		p[0].vt = VT_I4 | VT_BYREF;

				// pack the pDrawingRectangle parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo = NULL;
				CLSID clsidRECTANGLE = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{E46C3FEF-0F13-4d0d-88BE-D73F9B46B49E}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExTVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#else
					LPOLESTR clsid = OLESTR("{43DF2793-E1D0-42b4-9284-D1D3EB707840}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExTVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#endif
				VariantInit(&p[1]);
				p[1].vt = VT_RECORD | VT_BYREF;
				p[1].pRecInfo = pRecordInfo;
				p[1].pvRecord = pRecordInfo->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Bottom = pDrawingRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Left = pDrawingRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Right = pDrawingRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Top = pDrawingRectangle->Top;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_CUSTOMDRAW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[1].pvRecord);
				}
			}
		}

		/* Although pTextColor and pTextBackColor are of type OLE_COLOR, they may contain RGB colors only.
		   So convert OLE_COLOR colors into RGB colors. */
		if(*pTextColor & 0x80000000) {
			COLORREF color = RGB(0x00, 0x00, 0x00);
			OleTranslateColor(*pTextColor, NULL, &color);
			*pTextColor = static_cast<OLE_COLOR>(color);
		}
		if(*pTextBackColor & 0x80000000) {
			COLORREF color = RGB(0x00, 0x00, 0x00);
			OleTranslateColor(*pTextBackColor, NULL, &color);
			*pTextBackColor = static_cast<OLE_COLOR>(color);
		}

		return hr;
	}

	/// \brief <em>Fires the \c DblClick event</em>
	///
	/// \param[in] pTreeItem The double-clicked item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::DblClick, ExplorerTreeView::Raise_DblClick, Fire_Click,
	///       Fire_MDblClick, Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::DblClick, ExplorerTreeView::Raise_DblClick, Fire_Click,
	///       Fire_MDblClick, Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_DblClick(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_DBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::DestroyedControlWindow,
	///       ExplorerTreeView::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::DestroyedControlWindow,
	///       ExplorerTreeView::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
	/// \endif
	HRESULT Fire_DestroyedControlWindow(LONG hWnd)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWnd;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_DESTROYEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The label-edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::DestroyedEditControlWindow,
	///       ExplorerTreeView::Raise_DestroyedEditControlWindow, Fire_CreatedEditControlWindow,
	///       Fire_StartingLabelEditing
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::DestroyedEditControlWindow,
	///       ExplorerTreeView::Raise_DestroyedEditControlWindow, Fire_CreatedEditControlWindow,
	///       Fire_StartingLabelEditing
	/// \endif
	HRESULT Fire_DestroyedEditControlWindow(LONG hWndEdit)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndEdit;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_DESTROYEDEDITCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DragMouseMove event</em>
	///
	/// \param[in,out] ppDropTarget The item that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c ppDropTarget item's upper border.
	/// \param[in] hitTestDetails The part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoExpandItem If set to \c VARIANT_TRUE, the caller should auto-expand the item
	///                specified by \c ppDropTarget; otherwise it should cancel auto-expansion. See the
	///                following <strong>remarks</strong> section for details.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the control to the left; if set to a value greater than 0, it should
	///                scroll the control to the right. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the control upwardly; if set to a value greater than 0, it should
	///                scroll the control downwards. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Auto-expansion is timered, i. e. the caller should start the timer after this event, if
	///          the \c ppDropTarget parameter specifies another item than the last time this event was
	///          fired. If the timer isn't canceled, the caller should expand the item when the timer
	///          expires.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::DragMouseMove, ExplorerTreeView::Raise_DragMouseMove,
	///       Fire_MouseMove, Fire_OLEDragMouseMove
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::DragMouseMove, ExplorerTreeView::Raise_DragMouseMove,
	///       Fire_MouseMove, Fire_OLEDragMouseMove
	/// \endif
	HRESULT Fire_DragMouseMove(ITreeViewItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoExpandItem, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[10];
				p[9].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[9].vt = VT_DISPATCH | VT_BYREF;
				p[8] = button;																									p[8].vt = VT_I2;
				p[7] = shift;																										p[7].vt = VT_I2;
				p[6] = x;																												p[6].vt = VT_XPOS_PIXELS;
				p[5] = y;																												p[5].vt = VT_YPOS_PIXELS;
				p[4] = yToItemTop;																							p[4].vt = VT_I4;
				p[3].lVal = static_cast<LONG>(hitTestDetails);									p[3].vt = VT_I4;
				p[2].pboolVal = pAutoExpandItem;																p[2].vt = VT_BOOL | VT_BYREF;
				p[1].plVal = pAutoHScrollVelocity;															p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = pAutoVScrollVelocity;															p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 10, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_DRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Drop event</em>
	///
	/// \param[in] pDropTarget The item object that is the nearest one from the mouse cursor's position.
	///            If the mouse cursor's position lies outside the control's client area, this parameter
	///            should be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c pDropTarget item's upper border.
	/// \param[in] hitTestDetails The part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::Drop, ExplorerTreeView::Raise_Drop, Fire_AbortedDrag
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::Drop, ExplorerTreeView::Raise_Drop, Fire_AbortedDrag
	/// \endif
	HRESULT Fire_Drop(ITreeViewItem* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pDropTarget;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1] = yToItemTop;																p[1].vt = VT_I4;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_DROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditChange event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditChange, ExplorerTreeView::Raise_EditChange,
	///       Fire_EditKeyPress
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditChange, ExplorerTreeView::Raise_EditChange,
	///       Fire_EditKeyPress
	/// \endif
	HRESULT Fire_EditChange(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITCHANGE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the label-edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditClick, ExplorerTreeView::Raise_EditClick,
	///       Fire_EditDblClick, Fire_EditMClick, Fire_EditRClick, Fire_EditXClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditClick, ExplorerTreeView::Raise_EditClick,
	///       Fire_EditDblClick, Fire_EditMClick, Fire_EditRClick, Fire_EditXClick
	/// \endif
	HRESULT Fire_EditClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditContextMenu event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in,out] pShowDefaultMenu If \c VARIANT_FALSE, the caller should prevent the label-edit
	///                control from showing the default context menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditContextMenu,
	///       ExplorerTreeView::Raise_EditContextMenu, Fire_EditRClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditContextMenu,
	///       ExplorerTreeView::Raise_EditContextMenu, Fire_EditRClick
	/// \endif
	HRESULT Fire_EditContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;											p[4].vt = VT_I2;
				p[3] = shift;												p[3].vt = VT_I2;
				p[2] = x;														p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;														p[1].vt = VT_YPOS_PIXELS;
				p[0].pboolVal = pShowDefaultMenu;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITCONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the label-edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditDblClick, ExplorerTreeView::Raise_EditDblClick,
	///       Fire_EditClick, Fire_EditMDblClick, Fire_EditRDblClick, Fire_EditXDblClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditDblClick, ExplorerTreeView::Raise_EditDblClick,
	///       Fire_EditClick, Fire_EditMDblClick, Fire_EditRDblClick, Fire_EditXDblClick
	/// \endif
	HRESULT Fire_EditDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditGotFocus event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditGotFocus, ExplorerTreeView::Raise_EditGotFocus,
	///       Fire_EditLostFocus
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditGotFocus, ExplorerTreeView::Raise_EditGotFocus,
	///       Fire_EditLostFocus
	/// \endif
	HRESULT Fire_EditGotFocus(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITGOTFOCUS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditKeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditKeyDown, ExplorerTreeView::Raise_EditKeyDown,
	///       Fire_EditKeyUp, Fire_EditKeyPress
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditKeyDown, ExplorerTreeView::Raise_EditKeyDown,
	///       Fire_EditKeyUp, Fire_EditKeyPress
	/// \endif
	HRESULT Fire_EditKeyDown(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITKEYDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditKeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditKeyPress, ExplorerTreeView::Raise_EditKeyPress,
	///       Fire_EditKeyDown, Fire_EditKeyUp
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditKeyPress, ExplorerTreeView::Raise_EditKeyPress,
	///       Fire_EditKeyDown, Fire_EditKeyUp
	/// \endif
	HRESULT Fire_EditKeyPress(SHORT* pKeyAscii)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].piVal = pKeyAscii;		p[0].vt = VT_I2 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITKEYPRESS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditKeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditKeyUp, ExplorerTreeView::Raise_EditKeyUp,
	///       Fire_EditKeyDown, Fire_EditKeyPress
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditKeyUp, ExplorerTreeView::Raise_EditKeyUp,
	///       Fire_EditKeyDown, Fire_EditKeyPress
	/// \endif
	HRESULT Fire_EditKeyUp(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITKEYUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditLostFocus event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditLostFocus, ExplorerTreeView::Raise_EditLostFocus,
	///       Fire_EditGotFocus
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditLostFocus, ExplorerTreeView::Raise_EditLostFocus,
	///       Fire_EditGotFocus
	/// \endif
	HRESULT Fire_EditLostFocus(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITLOSTFOCUS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the label-edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditMClick, ExplorerTreeView::Raise_EditMClick,
	///       Fire_EditMDblClick, Fire_EditClick, Fire_EditRClick, Fire_EditXClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditMClick, ExplorerTreeView::Raise_EditMClick,
	///       Fire_EditMDblClick, Fire_EditClick, Fire_EditRClick, Fire_EditXClick
	/// \endif
	HRESULT Fire_EditMClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITMCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the label-edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditMDblClick, ExplorerTreeView::Raise_EditMDblClick,
	///       Fire_EditMClick, Fire_EditDblClick, Fire_EditRDblClick, Fire_EditXDblClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditMDblClick, ExplorerTreeView::Raise_EditMDblClick,
	///       Fire_EditMClick, Fire_EditDblClick, Fire_EditRDblClick, Fire_EditXDblClick
	/// \endif
	HRESULT Fire_EditMDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITMDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseDown event</em>
	///
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the label-edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditMouseDown, ExplorerTreeView::Raise_EditMouseDown,
	///       Fire_EditMouseUp, Fire_EditClick, Fire_EditMClick, Fire_EditRClick, Fire_EditXClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditMouseDown, ExplorerTreeView::Raise_EditMouseDown,
	///       Fire_EditMouseUp, Fire_EditClick, Fire_EditMClick, Fire_EditRClick, Fire_EditXClick
	/// \endif
	HRESULT Fire_EditMouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITMOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseEnter event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the label-edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditMouseEnter, ExplorerTreeView::Raise_EditMouseEnter,
	///       Fire_EditMouseLeave, Fire_EditMouseHover, Fire_EditMouseMove
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditMouseEnter, ExplorerTreeView::Raise_EditMouseEnter,
	///       Fire_EditMouseLeave, Fire_EditMouseHover, Fire_EditMouseMove
	/// \endif
	HRESULT Fire_EditMouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseHover event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the label-edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditMouseHover, ExplorerTreeView::Raise_EditMouseHover,
	///       Fire_EditMouseEnter, Fire_EditMouseLeave, Fire_EditMouseMove
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditMouseHover, ExplorerTreeView::Raise_EditMouseHover,
	///       Fire_EditMouseEnter, Fire_EditMouseLeave, Fire_EditMouseMove
	/// \endif
	HRESULT Fire_EditMouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITMOUSEHOVER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseLeave event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the label-edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditMouseLeave, ExplorerTreeView::Raise_EditMouseLeave,
	///       Fire_EditMouseEnter, Fire_EditMouseHover, Fire_EditMouseMove
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditMouseLeave, ExplorerTreeView::Raise_EditMouseLeave,
	///       Fire_EditMouseEnter, Fire_EditMouseHover, Fire_EditMouseMove
	/// \endif
	HRESULT Fire_EditMouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the label-edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditMouseMove, ExplorerTreeView::Raise_EditMouseMove,
	///       Fire_EditMouseEnter, Fire_EditMouseLeave, Fire_EditMouseDown, Fire_EditMouseUp,
	///       Fire_EditMouseWheel
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditMouseMove, ExplorerTreeView::Raise_EditMouseMove,
	///       Fire_EditMouseEnter, Fire_EditMouseLeave, Fire_EditMouseDown, Fire_EditMouseUp,
	///       Fire_EditMouseWheel
	/// \endif
	HRESULT Fire_EditMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseUp event</em>
	///
	/// \param[in] button The released mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the label-edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditMouseUp, ExplorerTreeView::Raise_EditMouseUp,
	///       Fire_EditMouseDown, Fire_EditClick, Fire_EditMClick, Fire_EditRClick, Fire_EditXClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditMouseUp, ExplorerTreeView::Raise_EditMouseUp,
	///       Fire_EditMouseDown, Fire_EditClick, Fire_EditMClick, Fire_EditRClick, Fire_EditXClick
	/// \endif
	HRESULT Fire_EditMouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITMOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseWheel event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditMouseWheel, ExplorerTreeView::Raise_EditMouseWheel,
	///       Fire_EditMouseMove, Fire_MouseWheel
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditMouseWheel, ExplorerTreeView::Raise_EditMouseWheel,
	///       Fire_EditMouseMove, Fire_MouseWheel
	/// \endif
	HRESULT Fire_EditMouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = button;																p[5].vt = VT_I2;
				p[4] = shift;																	p[4].vt = VT_I2;
				p[3] = x;																			p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																			p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(scrollAxis);		p[1].vt = VT_I4;
				p[0] = wheelDelta;														p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITMOUSEWHEEL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditRClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the label-edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditRClick, ExplorerTreeView::Raise_EditRClick,
	///       Fire_EditContextMenu, Fire_EditRDblClick, Fire_EditClick, Fire_EditMClick, Fire_EditXClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditRClick, ExplorerTreeView::Raise_EditRClick,
	///       Fire_EditContextMenu, Fire_EditRDblClick, Fire_EditClick, Fire_EditMClick, Fire_EditXClick
	/// \endif
	HRESULT Fire_EditRClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITRCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditRDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the label-edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the label-edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditRDblClick, ExplorerTreeView::Raise_EditRDblClick,
	///       Fire_EditRClick, Fire_EditDblClick, Fire_EditMDblClick, Fire_EditXDblClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditRDblClick, ExplorerTreeView::Raise_EditRDblClick,
	///       Fire_EditRClick, Fire_EditDblClick, Fire_EditMDblClick, Fire_EditXDblClick
	/// \endif
	HRESULT Fire_EditRDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITRDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditXClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be a
	///            constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the contained edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the contained edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditXClick, ExplorerTreeView::Raise_EditXClick,
	///       Fire_EditXDblClick, Fire_EditClick, Fire_EditMClick, Fire_EditRClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditXClick, ExplorerTreeView::Raise_EditXClick,
	///       Fire_EditXDblClick, Fire_EditClick, Fire_EditMClick, Fire_EditRClick
	/// \endif
	HRESULT Fire_EditXClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITXCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditXDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the contained
	///            edit control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::EditXDblClick, ExplorerTreeView::Raise_EditXDblClick,
	///       Fire_EditXClick, Fire_EditDblClick, Fire_EditMDblClick, Fire_EditRDblClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::EditXDblClick, ExplorerTreeView::Raise_EditXDblClick,
	///       Fire_EditXClick, Fire_EditDblClick, Fire_EditMDblClick, Fire_EditRDblClick
	/// \endif
	HRESULT Fire_EditXDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EDITXDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ExpandedItem event</em>
	///
	/// \param[in] pTreeItem The expanded item.
	/// \param[in] expandedPartially If \c VARIANT_TRUE, the "+" button next to the item was replaced
	///            with a "-" button; otherwise the button remained a "+".
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ExpandedItem, ExplorerTreeView::Raise_ExpandedItem,
	///       Fire_ExpandingItem, Fire_CollapsedItem
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ExpandedItem, ExplorerTreeView::Raise_ExpandedItem,
	///       Fire_ExpandingItem, Fire_CollapsedItem
	/// \endif
	HRESULT Fire_ExpandedItem(ITreeViewItem* pTreeItem, VARIANT_BOOL expandedPartially)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pTreeItem;
				p[0] = expandedPartially;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EXPANDEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ExpandingItem event</em>
	///
	/// \param[in] pTreeItem The expanding item.
	/// \param[in] expandingPartially If \c VARIANT_TRUE, the "+" button next to the item is replaced
	///            with a "-" button; otherwise the button remains a "+".
	/// \param[in,out] pCancelExpansion If \c VARIANT_TRUE, the caller should abort the expand action;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ExpandingItem, ExplorerTreeView::Raise_ExpandingItem,
	///       Fire_ExpandedItem, Fire_CollapsingItem
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ExpandingItem, ExplorerTreeView::Raise_ExpandingItem,
	///       Fire_ExpandedItem, Fire_CollapsingItem
	/// \endif
	HRESULT Fire_ExpandingItem(ITreeViewItem* pTreeItem, VARIANT_BOOL expandingPartially, VARIANT_BOOL* pCancelExpansion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pTreeItem;
				p[1] = expandingPartially;
				p[0].pboolVal = pCancelExpansion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_EXPANDINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c FreeItemData event</em>
	///
	/// \param[in] pTreeItem The item whose associated data shall be freed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::FreeItemData, ExplorerTreeView::Raise_FreeItemData,
	///       Fire_RemovingItem, Fire_RemovedItem
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::FreeItemData, ExplorerTreeView::Raise_FreeItemData,
	///       Fire_RemovingItem, Fire_RemovedItem
	/// \endif
	HRESULT Fire_FreeItemData(ITreeViewItem* pTreeItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pTreeItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_FREEITEMDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c IncrementalSearchStringChanging event</em>
	///
	/// \param[in] currentSearchString The control's current incremental search-string.
	/// \param[in] keyCodeOfCharToBeAdded The key code of the character to be added to the search-string.
	///            Most of the values defined by VB's \c KeyCodeConstants enumeration are valid.
	/// \param[in,out] pCancelChange If \c VARIANT_TRUE, the caller should discard the character and
	///                leave the search-string unchanged; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::IncrementalSearchStringChanging,
	///       ExplorerTreeView::Raise_IncrementalSearchStringChanging, Fire_KeyPress
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::IncrementalSearchStringChanging,
	///       ExplorerTreeView::Raise_IncrementalSearchStringChanging, Fire_KeyPress
	/// \endif
	HRESULT Fire_IncrementalSearchStringChanging(BSTR currentSearchString, SHORT keyCodeOfCharToBeAdded, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = currentSearchString;
				p[1] = keyCodeOfCharToBeAdded;
				p[0].pboolVal = pCancelChange;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_INCREMENTALSEARCHSTRINGCHANGING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertedItem event</em>
	///
	/// \param[in] pTreeItem The inserted item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::InsertedItem, ExplorerTreeView::Raise_InsertedItem,
	///       Fire_InsertingItem, Fire_RemovedItem
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::InsertedItem, ExplorerTreeView::Raise_InsertedItem,
	///       Fire_InsertingItem, Fire_RemovedItem
	/// \endif
	HRESULT Fire_InsertedItem(ITreeViewItem* pTreeItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pTreeItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_INSERTEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertingItem event</em>
	///
	/// \param[in] pTreeItem The item being inserted.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::InsertingItem, ExplorerTreeView::Raise_InsertingItem,
	///       Fire_InsertedItem, Fire_RemovingItem
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::InsertingItem, ExplorerTreeView::Raise_InsertingItem,
	///       Fire_InsertedItem, Fire_RemovingItem
	/// \endif
	HRESULT Fire_InsertingItem(IVirtualTreeViewItem* pTreeItem, VARIANT_BOOL* pCancelInsertion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pTreeItem;
				p[0].pboolVal = pCancelInsertion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_INSERTINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemAsynchronousDrawFailed event</em>
	///
	/// \param[in] pTreeItem The item whose image failed to be drawn.
	/// \param[in,out] pNotificationDetails A \c NMTVASYNCDRAW struct holding the notification details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ItemAsynchronousDrawFailed,
	///       ExplorerTreeView::Raise_ItemAsynchronousDrawFailed, ExTVwLibU::FAILEDIMAGEDETAILS,
	///       <a href="https://msdn.microsoft.com/en-us/library/aa361676.aspx">NMTVASYNCDRAW</a>
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ItemAsynchronousDrawFailed,
	///       ExplorerTreeView::Raise_ItemAsynchronousDrawFailed, ExTVwLibA::FAILEDIMAGEDETAILS,
	///       <a href="https://msdn.microsoft.com/en-us/library/aa361676.aspx">NMTVASYNCDRAW</a>
	/// \endif
	HRESULT Fire_ItemAsynchronousDrawFailed(ITreeViewItem* pTreeItem, NMTVASYNCDRAW* pNotificationDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = pTreeItem;
				p[2].lVal = static_cast<LONG>(pNotificationDetails->hr);												p[2].vt = VT_I4;
				p[1].plVal = reinterpret_cast<PLONG>(&pNotificationDetails->dwRetFlags);				p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = reinterpret_cast<PLONG>(&pNotificationDetails->iRetImageIndex);		p[0].vt = VT_I4 | VT_BYREF;

				// pack the pNotificationDetails->pimldp parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo = NULL;
				CLSID clsidFAILEDIMAGEDETAILS = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{95ECE649-48D0-4133-9D53-31304AA41EE1}");
					CLSIDFromString(clsid, &clsidFAILEDIMAGEDETAILS);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExTVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidFAILEDIMAGEDETAILS), &pRecordInfo)));
				#else
					LPOLESTR clsid = OLESTR("{C5F09B61-A863-4e6e-80F5-0E6949CCD241}");
					CLSIDFromString(clsid, &clsidFAILEDIMAGEDETAILS);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExTVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidFAILEDIMAGEDETAILS), &pRecordInfo)));
				#endif
				VariantInit(&p[3]);
				p[3].vt = VT_RECORD | VT_BYREF;
				p[3].pRecInfo = pRecordInfo;
				p[3].pvRecord = pRecordInfo->RecordCreate();
				// transfer data
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->hImageList = HandleToLong(pNotificationDetails->pimldp->himl);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->IconIndex = pNotificationDetails->pimldp->i;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->hDC = HandleToLong(pNotificationDetails->pimldp->hdcDst);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->TargetPositionX = pNotificationDetails->pimldp->x;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->TargetPositionY = pNotificationDetails->pimldp->y;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawLeft = pNotificationDetails->pimldp->xBitmap;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawTop = pNotificationDetails->pimldp->yBitmap;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawWidth = pNotificationDetails->pimldp->cx;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawHeight = pNotificationDetails->pimldp->cy;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->BackColor = pNotificationDetails->pimldp->rgbBk;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->ForeColor = pNotificationDetails->pimldp->rgbFg;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->DrawingStyle = static_cast<ImageDrawingStyleConstants>(pNotificationDetails->pimldp->fStyle & ~ILD_OVERLAYMASK);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->OverlayIndex = (pNotificationDetails->pimldp->fStyle & ILD_OVERLAYMASK) >> 8;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->RasterOperationCode = pNotificationDetails->pimldp->dwRop;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->DrawingEffects = static_cast<ImageDrawingEffectsConstants>(pNotificationDetails->pimldp->fState);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->AlphaChannel = static_cast<BYTE>(pNotificationDetails->pimldp->Frame);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->EffectColor = pNotificationDetails->pimldp->crEffect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_ITEMASYNCHRONOUSDRAWFAILED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				pNotificationDetails->pimldp->himl = static_cast<HIMAGELIST>(LongToHandle(reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->hImageList));
				pNotificationDetails->pimldp->i = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->IconIndex;
				pNotificationDetails->pimldp->hdcDst = static_cast<HDC>(LongToHandle(reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->hDC));
				pNotificationDetails->pimldp->x = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->TargetPositionX;
				pNotificationDetails->pimldp->y = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->TargetPositionY;
				pNotificationDetails->pimldp->xBitmap = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawLeft;
				pNotificationDetails->pimldp->yBitmap = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawTop;
				pNotificationDetails->pimldp->cx = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawWidth;
				pNotificationDetails->pimldp->cy = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawHeight;
				pNotificationDetails->pimldp->rgbBk = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->BackColor;
				pNotificationDetails->pimldp->rgbFg = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->ForeColor;
				pNotificationDetails->pimldp->fStyle = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->DrawingStyle | INDEXTOOVERLAYMASK(reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->OverlayIndex);
				pNotificationDetails->pimldp->dwRop = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->RasterOperationCode;
				pNotificationDetails->pimldp->fState = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->DrawingEffects;
				pNotificationDetails->pimldp->Frame = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->AlphaChannel;
				pNotificationDetails->pimldp->crEffect = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->EffectColor;

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[3].pvRecord);
				}
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemBeginDrag event</em>
	///
	/// \param[in] pTreeItem The item that the user wants to drag.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbLeftButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ItemBeginDrag, ExplorerTreeView::Raise_ItemBeginDrag,
	///       Fire_ItemBeginRDrag
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ItemBeginDrag, ExplorerTreeView::Raise_ItemBeginDrag,
	///       Fire_ItemBeginRDrag
	/// \endif
	HRESULT Fire_ItemBeginDrag(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_ITEMBEGINDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemBeginRDrag event</em>
	///
	/// \param[in] pTreeItem The item that the user wants to drag.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbRightButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ItemBeginRDrag, ExplorerTreeView::Raise_ItemBeginRDrag,
	///       Fire_ItemBeginDrag
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ItemBeginRDrag, ExplorerTreeView::Raise_ItemBeginRDrag,
	///       Fire_ItemBeginDrag
	/// \endif
	HRESULT Fire_ItemBeginRDrag(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_ITEMBEGINRDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemGetDisplayInfo event</em>
	///
	/// \param[in] pTreeItem The item whose properties are requested.
	/// \param[in] requestedInfo Specifies which properties' values are required. Any combination of
	///            the values defined by the \c RequestedInfoConstants enumeration is valid.
	/// \param[out] pIconIndex The zero-based index of the requested icon. If the \c requestedInfo
	///             parameter doesn't include \c riIconIndex, the caller should ignore this value.
	/// \param[out] pSelectedIconIndex The zero-based index of the requested selected icon. If the
	///             \c requestedInfo parameter doesn't include \c riSelectedIconIndex, the caller should
	///             ignore this value.
	/// \param[out] pExpandedIconIndex The zero-based index of the requested expanded icon. If the
	///             \c requestedInfo parameter doesn't include \c riExpandedIconIndex, the caller should
	///             ignore this value.
	/// \param[out] pHasButton If set to \c VARIANT_TRUE, the caller should display an item button next
	///             to the item; otherwise not. If the \c requestedInfo parameter doesn't include
	///             \c riHasExpando, the caller should ignore this value.
	/// \param[in] maxItemTextLength The maximum number of characters the item's text may consist of.
	///            If the \c requestedInfo parameter doesn't include \c riItemText, the client should
	///            ignore this value.
	/// \param[out] pItemText The item's text. If the \c requestedInfo parameter doesn't include
	///             \c riItemText, the caller should ignore this value.
	/// \param[in,out] pDontAskAgain If \c VARIANT_TRUE, the caller should always use the same settings
	///                and never fire this event again for these properties of this item; otherwise it
	///                shouldn't persist the values.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ItemGetDisplayInfo,
	///       ExplorerTreeView::Raise_ItemGetDisplayInfo, Fire_ItemSetText
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ItemGetDisplayInfo,
	///       ExplorerTreeView::Raise_ItemGetDisplayInfo, Fire_ItemSetText
	/// \endif
	HRESULT Fire_ItemGetDisplayInfo(ITreeViewItem* pTreeItem, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG* pSelectedIconIndex, LONG* pExpandedIconIndex, VARIANT_BOOL* pHasButton, LONG maxItemTextLength, BSTR* pItemText, VARIANT_BOOL* pDontAskAgain)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pTreeItem;
				p[7] = static_cast<LONG>(requestedInfo);				p[7].vt = VT_I4;
				p[6].plVal = pIconIndex;												p[6].vt = VT_I4 | VT_BYREF;
				p[5].plVal = pSelectedIconIndex;								p[5].vt = VT_I4 | VT_BYREF;
				p[4].plVal = pExpandedIconIndex;								p[4].vt = VT_I4 | VT_BYREF;
				p[3].pboolVal = pHasButton;											p[3].vt = VT_BOOL | VT_BYREF;
				p[2] = maxItemTextLength;
				p[1].pbstrVal = pItemText;											p[1].vt = VT_BSTR | VT_BYREF;
				p[0].pboolVal = pDontAskAgain;									p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_ITEMGETDISPLAYINFO, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemGetInfoTipText event</em>
	///
	/// \param[in] pTreeItem The item whose info tip text is requested.
	/// \param[in] maxInfoTipLength The maximum number of characters the info tip text may consist of.
	/// \param[out] pInfoTipText The item's info tip text.
	/// \param[in,out] pAbortToolTip If \c VARIANT_TRUE, the caller should NOT show the tooltip;
	///                otherwise it should.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ItemGetInfoTipText,
	///       ExplorerTreeView::Raise_ItemGetInfoTipText
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ItemGetInfoTipText,
	///       ExplorerTreeView::Raise_ItemGetInfoTipText
	/// \endif
	HRESULT Fire_ItemGetInfoTipText(ITreeViewItem* pTreeItem, LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pTreeItem;
				p[2] = maxInfoTipLength;
				p[1].pbstrVal = pInfoTipText;			p[1].vt = VT_BSTR | VT_BYREF;
				p[0].pboolVal = pAbortToolTip;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_ITEMGETINFOTIPTEXT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemMouseEnter event</em>
	///
	/// \param[in] pTreeItem The item that was entered.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ItemMouseEnter, ExplorerTreeView::Raise_ItemMouseEnter,
	///       Fire_ItemMouseLeave, Fire_MouseMove
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ItemMouseEnter, ExplorerTreeView::Raise_ItemMouseEnter,
	///       Fire_ItemMouseLeave, Fire_MouseMove
	/// \endif
	HRESULT Fire_ItemMouseEnter(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_ITEMMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemMouseLeave event</em>
	///
	/// \param[in] pTreeItem The item that was left.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ItemMouseLeave, ExplorerTreeView::Raise_ItemMouseLeave,
	///       Fire_ItemMouseEnter, Fire_MouseMove
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ItemMouseLeave, ExplorerTreeView::Raise_ItemMouseLeave,
	///       Fire_ItemMouseEnter, Fire_MouseMove
	/// \endif
	HRESULT Fire_ItemMouseLeave(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_ITEMMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemSelectionChanged event</em>
	///
	/// \param[in] pTreeItem The item that was selected/deselected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ItemSelectionChanged,
	///       ExplorerTreeView::Raise_ItemSelectionChanged, Fire_ItemSelectionChanging, Fire_CaretChanged
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ItemSelectionChanged,
	///       ExplorerTreeView::Raise_ItemSelectionChanged, Fire_ItemSelectionChanging, Fire_CaretChanged
	/// \endif
	HRESULT Fire_ItemSelectionChanged(ITreeViewItem* pTreeItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); i++) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pTreeItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_ITEMSELECTIONCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemSelectionChanging event</em>
	///
	/// \param[in] pTreeItem The item that will be selected/deselected.
	/// \param[in,out] pCancelChange If \c VARIANT_TRUE, the caller should prevent the selection state from
	///                changing; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ItemSelectionChanging,
	///       ExplorerTreeView::Raise_ItemSelectionChanging, Fire_ItemSelectionChanged,
	///       Fire_CaretChanging
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ItemSelectionChanging,
	///       ExplorerTreeView::Raise_ItemSelectionChanging, Fire_ItemSelectionChanged,
	///       Fire_CaretChanging
	/// \endif
	HRESULT Fire_ItemSelectionChanging(ITreeViewItem* pTreeItem, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); i++) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pTreeItem;
				p[0].pboolVal = pCancelChange;			p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_ITEMSELECTIONCHANGING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemSetText event</em>
	///
	/// \param[in] pTreeItem The item whose text was changed.
	/// \param[in] itemText The item's new text.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ItemSetText, ExplorerTreeView::Raise_ItemSetText,
	///       Fire_ItemGetDisplayInfo
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ItemSetText, ExplorerTreeView::Raise_ItemSetText,
	///       Fire_ItemGetDisplayInfo
	/// \endif
	HRESULT Fire_ItemSetText(ITreeViewItem* pTreeItem, BSTR itemText)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pTreeItem;
				p[0] = itemText;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_ITEMSETTEXT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemStateImageChanged event</em>
	///
	/// \param[in] pTreeItem The item whose state image was changed.
	/// \param[in] previousStateImageIndex The one-based index of the item's previous state image.
	/// \param[in] newStateImageIndex The one-based index of the item's new state image.
	/// \param[in] causedBy The reason for the state image change. Any of the values defined by the
	///            \c StateImageChangeCausedByConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ItemStateImageChanged,
	///       ExplorerTreeView::Raise_ItemStateImageChanged, Fire_ItemStateImageChanging
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ItemStateImageChanged,
	///       ExplorerTreeView::Raise_ItemStateImageChanged, Fire_ItemStateImageChanging
	/// \endif
	HRESULT Fire_ItemStateImageChanged(ITreeViewItem* pTreeItem, LONG previousStateImageIndex, LONG newStateImageIndex, StateImageChangeCausedByConstants causedBy)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pTreeItem;
				p[2] = previousStateImageIndex;
				p[1] = newStateImageIndex;
				p[0] = static_cast<LONG>(causedBy);			p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_ITEMSTATEIMAGECHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemStateImageChanging event</em>
	///
	/// \param[in] pTreeItem The item whose state image shall be changed.
	/// \param[in] previousStateImageIndex The one-based index of the item's previous state image.
	/// \param[in,out] pNewStateImageIndex The one-based index of the item's new state image. The client
	///                may change this value.
	/// \param[in] causedBy The reason for the state image change. Any of the values defined by the
	///            \c StateImageChangeCausedByConstants enumeration is valid.
	/// \param[in,out] pCancelChange If \c VARIANT_TRUE, the caller should prevent the state image from
	///                changing; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ItemStateImageChanging,
	///       ExplorerTreeView::Raise_ItemStateImageChanging, Fire_ItemStateImageChanged
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ItemStateImageChanging,
	///       ExplorerTreeView::Raise_ItemStateImageChanging, Fire_ItemStateImageChanged
	/// \endif
	HRESULT Fire_ItemStateImageChanging(ITreeViewItem* pTreeItem, LONG previousStateImageIndex, LONG* pNewStateImageIndex, StateImageChangeCausedByConstants causedBy, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = pTreeItem;
				p[3] = previousStateImageIndex;
				p[2].plVal = pNewStateImageIndex;				p[2].vt = VT_I4 | VT_BYREF;
				p[1] = static_cast<LONG>(causedBy);			p[1].vt = VT_I4;
				p[0].pboolVal = pCancelChange;					p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_ITEMSTATEIMAGECHANGING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::KeyDown, ExplorerTreeView::Raise_KeyDown, Fire_KeyUp,
	///       Fire_KeyPress
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::KeyDown, ExplorerTreeView::Raise_KeyDown, Fire_KeyUp,
	///       Fire_KeyPress
	/// \endif
	HRESULT Fire_KeyDown(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_KEYDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::KeyPress, ExplorerTreeView::Raise_KeyPress,
	///       Fire_IncrementalSearchStringChanging, Fire_KeyDown, Fire_KeyUp
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::KeyPress, ExplorerTreeView::Raise_KeyPress,
	///       Fire_IncrementalSearchStringChanging, Fire_KeyDown, Fire_KeyUp
	/// \endif
	HRESULT Fire_KeyPress(SHORT* pKeyAscii)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].piVal = pKeyAscii;		p[0].vt = VT_I2 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_KEYPRESS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::KeyUp, ExplorerTreeView::Raise_KeyUp, Fire_KeyDown,
	///       Fire_KeyPress
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::KeyUp, ExplorerTreeView::Raise_KeyUp, Fire_KeyDown,
	///       Fire_KeyPress
	/// \endif
	HRESULT Fire_KeyUp(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_KEYUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MClick event</em>
	///
	/// \param[in] pTreeItem The clicked item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::MClick, ExplorerTreeView::Raise_MClick, Fire_MDblClick,
	///       Fire_Click, Fire_RClick, Fire_XClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::MClick, ExplorerTreeView::Raise_MClick, Fire_MDblClick,
	///       Fire_Click, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MClick(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_MCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MDblClick event</em>
	///
	/// \param[in] pTreeItem The double-clicked item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::MDblClick, ExplorerTreeView::Raise_MDblClick,
	///       Fire_MClick, Fire_DblClick, Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::MDblClick, ExplorerTreeView::Raise_MDblClick,
	///       Fire_MClick, Fire_DblClick, Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_MDblClick(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_MDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseDown event</em>
	///
	/// \param[in] pTreeItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::MouseDown, ExplorerTreeView::Raise_MouseDown,
	///       Fire_MouseUp, Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::MouseDown, ExplorerTreeView::Raise_MouseDown,
	///       Fire_MouseUp, Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseDown(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_MOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseEnter event</em>
	///
	/// \param[in] pTreeItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::MouseEnter, ExplorerTreeView::Raise_MouseEnter,
	///       Fire_MouseLeave, Fire_ItemMouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::MouseEnter, ExplorerTreeView::Raise_MouseEnter,
	///       Fire_MouseLeave, Fire_ItemMouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseEnter(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_MOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseHover event</em>
	///
	/// \param[in] pTreeItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::MouseHover, ExplorerTreeView::Raise_MouseHover,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseMove
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::MouseHover, ExplorerTreeView::Raise_MouseHover,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseHover(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_MOUSEHOVER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseLeave event</em>
	///
	/// \param[in] pTreeItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::MouseLeave, ExplorerTreeView::Raise_MouseLeave,
	///       Fire_MouseEnter, Fire_ItemMouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::MouseLeave, ExplorerTreeView::Raise_MouseLeave,
	///       Fire_MouseEnter, Fire_ItemMouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseLeave(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_MOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseMove event</em>
	///
	/// \param[in] pTreeItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::MouseMove, ExplorerTreeView::Raise_MouseMove,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp, Fire_MouseWheel
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::MouseMove, ExplorerTreeView::Raise_MouseMove,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp, Fire_MouseWheel
	/// \endif
	HRESULT Fire_MouseMove(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_MOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseUp event</em>
	///
	/// \param[in] pTreeItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::MouseUp, ExplorerTreeView::Raise_MouseUp,
	///       Fire_MouseDown, Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::MouseUp, ExplorerTreeView::Raise_MouseUp,
	///       Fire_MouseDown, Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseUp(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_MOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseWheel event</em>
	///
	/// \param[in] pTreeItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::MouseWheel, ExplorerTreeView::Raise_MouseWheel,
	///       Fire_MouseMove, Fire_EditMouseWheel
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::MouseWheel, ExplorerTreeView::Raise_MouseWheel,
	///       Fire_MouseMove, Fire_EditMouseWheel
	/// \endif
	HRESULT Fire_MouseWheel(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pTreeItem;
				p[6] = button;																		p[6].vt = VT_I2;
				p[5] = shift;																			p[5].vt = VT_I2;
				p[4] = x;																					p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																					p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(hitTestDetails);		p[2].vt = VT_I4;
				p[1].lVal = static_cast<LONG>(scrollAxis);				p[1].vt = VT_I4;
				p[0] = wheelDelta;																p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_MOUSEWHEEL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLECompleteDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] performedEffect The performed drop effect. Any of the values (except \c odeScroll)
	///            defined by the \c OLEDropEffectConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::OLECompleteDrag, ExplorerTreeView::Raise_OLECompleteDrag,
	///       Fire_OLEStartDrag
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::OLECompleteDrag, ExplorerTreeView::Raise_OLECompleteDrag,
	///       Fire_OLEStartDrag
	/// \endif
	HRESULT Fire_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pData;
				p[0].lVal = static_cast<LONG>(performedEffect);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_OLECOMPLETEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client finally executed.
	/// \param[in,out] ppDropTarget The item that is the target of the drag'n'drop operation. The client
	///                may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c pDropTarget item's upper border.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::OLEDragDrop, ExplorerTreeView::Raise_OLEDragDrop,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::OLEDragDrop, ExplorerTreeView::Raise_OLEDragDrop,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp
	/// \endif
	HRESULT Fire_OLEDragDrop(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, ITreeViewItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pData;
				p[7].plVal = reinterpret_cast<PLONG>(pEffect);									p[7].vt = VT_I4 | VT_BYREF;
				p[6].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[6].vt = VT_DISPATCH | VT_BYREF;
				p[5] = button;																									p[5].vt = VT_I2;
				p[4] = shift;																										p[4].vt = VT_I2;
				p[3] = x;																												p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																												p[2].vt = VT_YPOS_PIXELS;
				p[1] = yToItemTop;																							p[1].vt = VT_I4;
				p[0].lVal = static_cast<LONG>(hitTestDetails);									p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_OLEDRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The item that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c ppDropTarget item's upper border.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoExpandItem If set to \c VARIANT_TRUE, the caller should auto-expand the item
	///                specified by \c ppDropTarget; otherwise it should cancel auto-expansion. See the
	///                following <strong>remarks</strong> section for details.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the control to the left; if set to a value greater than 0, it should
	///                scroll the control to the right. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the control upwardly; if set to a value greater than 0, it should
	///                scroll the control downwards. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Auto-expansion is timered, i. e. the caller should start the timer after this event, if
	///          the \c ppDropTarget parameter specifies another item than the last time an \c OLEDrag*
	///          event was fired. If the timer isn't canceled, the caller should expand the item when the
	///          timer expires.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::OLEDragEnter, ExplorerTreeView::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::OLEDragEnter, ExplorerTreeView::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter
	/// \endif
	HRESULT Fire_OLEDragEnter(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, ITreeViewItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoExpandItem, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[12];
				p[11] = pData;
				p[10].plVal = reinterpret_cast<PLONG>(pEffect);									p[10].vt = VT_I4 | VT_BYREF;
				p[9].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[9].vt = VT_DISPATCH | VT_BYREF;
				p[8] = button;																									p[8].vt = VT_I2;
				p[7] = shift;																										p[7].vt = VT_I2;
				p[6] = x;																												p[6].vt = VT_XPOS_PIXELS;
				p[5] = y;																												p[5].vt = VT_YPOS_PIXELS;
				p[4] = yToItemTop;																							p[4].vt = VT_I4;
				p[3].lVal = static_cast<LONG>(hitTestDetails);									p[3].vt = VT_I4;
				p[2].pboolVal = pAutoExpandItem;																p[2].vt = VT_BOOL | VT_BYREF;
				p[1].plVal = pAutoHScrollVelocity;															p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = pAutoVScrollVelocity;															p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 12, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_OLEDRAGENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragEnterPotentialTarget event</em>
	///
	/// \param[in] hWndPotentialTarget The handle of the potential drag'n'drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::OLEDragEnterPotentialTarget,
	///       ExplorerTreeView::Raise_OLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::OLEDragEnterPotentialTarget,
	///       ExplorerTreeView::Raise_OLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget
	/// \endif
	HRESULT Fire_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndPotentialTarget;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_OLEDRAGENTERPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragLeave event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in] pDropTarget The item that is the current target of the drag'n'drop operation.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c ppDropTarget item's upper border.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::OLEDragLeave, ExplorerTreeView::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::OLEDragLeave, ExplorerTreeView::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave
	/// \endif
	HRESULT Fire_OLEDragLeave(IOLEDataObject* pData, ITreeViewItem* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pData;
				p[6] = pDropTarget;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1] = yToItemTop;																p[1].vt = VT_I4;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_OLEDRAGLEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragLeavePotentialTarget event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::OLEDragLeavePotentialTarget,
	///       ExplorerTreeView::Raise_OLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::OLEDragLeavePotentialTarget,
	///       ExplorerTreeView::Raise_OLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget
	/// \endif
	HRESULT Fire_OLEDragLeavePotentialTarget(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_OLEDRAGLEAVEPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragMouseMove event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The item that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c ppDropTarget item's upper border.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoExpandItem If set to \c VARIANT_TRUE, the caller should auto-expand the item
	///                specified by \c ppDropTarget; otherwise it should cancel auto-expansion. See the
	///                following <strong>remarks</strong> section for details.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the control to the left; if set to a value greater than 0, it should
	///                scroll the control to the right. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the control upwardly; if set to a value greater than 0, it should
	///                scroll the control downwards. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Auto-expansion is timered, i. e. the caller should start the timer after this event, if
	///          the \c ppDropTarget parameter specifies another item than the last time an \c OLEDrag*
	///          event was fired. If the timer isn't canceled, the caller should expand the item when the
	///          timer expires.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::OLEDragMouseMove,
	///       ExplorerTreeView::Raise_OLEDragMouseMove, Fire_OLEDragEnter, Fire_OLEDragLeave,
	///       Fire_OLEDragDrop, Fire_MouseMove
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::OLEDragMouseMove,
	///       ExplorerTreeView::Raise_OLEDragMouseMove, Fire_OLEDragEnter, Fire_OLEDragLeave,
	///       Fire_OLEDragDrop, Fire_MouseMove
	/// \endif
	HRESULT Fire_OLEDragMouseMove(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, ITreeViewItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoExpandItem, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[12];
				p[11] = pData;
				p[10].plVal = reinterpret_cast<PLONG>(pEffect);									p[10].vt = VT_I4 | VT_BYREF;
				p[9].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[9].vt = VT_DISPATCH | VT_BYREF;
				p[8] = button;																									p[8].vt = VT_I2;
				p[7] = shift;																										p[7].vt = VT_I2;
				p[6] = x;																												p[6].vt = VT_XPOS_PIXELS;
				p[5] = y;																												p[5].vt = VT_YPOS_PIXELS;
				p[4] = yToItemTop;																							p[4].vt = VT_I4;
				p[3].lVal = static_cast<LONG>(hitTestDetails);									p[3].vt = VT_I4;
				p[2].pboolVal = pAutoExpandItem;																p[2].vt = VT_BOOL | VT_BYREF;
				p[1].plVal = pAutoHScrollVelocity;															p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = pAutoVScrollVelocity;															p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 12, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_OLEDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEGiveFeedback event</em>
	///
	/// \param[in] effect The current drop effect. It is chosen by the potential drop target. Any of
	///            the values defined by the \c OLEDropEffectConstants enumeration is valid.
	/// \param[in,out] pUseDefaultCursors If set to \c VARIANT_TRUE, the system's default mouse cursors
	///                shall be used to visualize the various drop effects. If set to \c VARIANT_FALSE,
	///                the client has set a custom mouse cursor.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::OLEGiveFeedback,
	///       ExplorerTreeView::Raise_OLEGiveFeedback, Fire_OLEQueryContinueDrag
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::OLEGiveFeedback,
	///       ExplorerTreeView::Raise_OLEGiveFeedback, Fire_OLEQueryContinueDrag
	/// \endif
	HRESULT Fire_OLEGiveFeedback(OLEDropEffectConstants effect, VARIANT_BOOL* pUseDefaultCursors)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = static_cast<LONG>(effect);			p[1].vt = VT_I4;
				p[0].pboolVal = pUseDefaultCursors;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_OLEGIVEFEEDBACK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEQueryContinueDrag event</em>
	///
	/// \param[in] pressedEscape If \c VARIANT_TRUE, the user has pressed the \c ESC key since the last
	///            time this event was fired; otherwise not.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in,out] pActionToContinueWith Indicates whether to continue, cancel or complete the
	///                drag'n'drop operation. Any of the values defined by the
	///                \c OLEActionToContinueWithConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::OLEQueryContinueDrag,
	///       ExplorerTreeView::Raise_OLEQueryContinueDrag, Fire_OLEGiveFeedback
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::OLEQueryContinueDrag,
	///       ExplorerTreeView::Raise_OLEQueryContinueDrag, Fire_OLEGiveFeedback
	/// \endif
	HRESULT Fire_OLEQueryContinueDrag(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, OLEActionToContinueWithConstants* pActionToContinueWith)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pressedEscape;
				p[2] = button;																									p[2].vt = VT_I2;
				p[1] = shift;																										p[1].vt = VT_I2;
				p[0].plVal = reinterpret_cast<PLONG>(pActionToContinueWith);		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_OLEQUERYCONTINUEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEReceivedNewData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the data object has received data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::OLEReceivedNewData,
	///       ExplorerTreeView::Raise_OLEReceivedNewData, Fire_OLESetData
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::OLEReceivedNewData,
	///       ExplorerTreeView::Raise_OLEReceivedNewData, Fire_OLESetData
	/// \endif
	HRESULT Fire_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pData;
				p[2] = formatID;
				p[1] = index;
				p[0] = dataOrViewAspect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_OLERECEIVEDNEWDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLESetData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the drop target is requesting data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::OLESetData,
	///       ExplorerTreeView::Raise_OLESetData, Fire_OLEStartDrag
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::OLESetData,
	///       ExplorerTreeView::Raise_OLESetData, Fire_OLEStartDrag
	/// \endif
	HRESULT Fire_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pData;
				p[2] = formatID;
				p[1] = index;
				p[0] = dataOrViewAspect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_OLESETDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEStartDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::OLEStartDrag, ExplorerTreeView::Raise_OLEStartDrag,
	///       Fire_OLESetData, Fire_OLECompleteDrag
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::OLEStartDrag, ExplorerTreeView::Raise_OLEStartDrag,
	///       Fire_OLESetData, Fire_OLECompleteDrag
	/// \endif
	HRESULT Fire_OLEStartDrag(IOLEDataObject* pData)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pData;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_OLESTARTDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RClick event</em>
	///
	/// \param[in] pTreeItem The clicked item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::RClick, ExplorerTreeView::Raise_RClick,
	///       Fire_ContextMenu, Fire_RDblClick, Fire_Click, Fire_MClick, Fire_XClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::RClick, ExplorerTreeView::Raise_RClick,
	///       Fire_ContextMenu, Fire_RDblClick, Fire_Click, Fire_MClick, Fire_XClick
	/// \endif
	HRESULT Fire_RClick(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_RCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RDblClick event</em>
	///
	/// \param[in] pTreeItem The double-clicked item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::RDblClick, ExplorerTreeView::Raise_RDblClick,
	///       Fire_RClick, Fire_DblClick, Fire_MDblClick, Fire_XDblClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::RDblClick, ExplorerTreeView::Raise_RDblClick,
	///       Fire_RClick, Fire_DblClick, Fire_MDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_RDblClick(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_RDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::RecreatedControlWindow,
	///       ExplorerTreeView::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::RecreatedControlWindow,
	///       ExplorerTreeView::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
	/// \endif
	HRESULT Fire_RecreatedControlWindow(LONG hWnd)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWnd;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_RECREATEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovedItem event</em>
	///
	/// \param[in] pTreeItem The removed item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::RemovedItem, ExplorerTreeView::Raise_RemovedItem,
	///       Fire_RemovingItem, Fire_InsertedItem
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::RemovedItem, ExplorerTreeView::Raise_RemovedItem,
	///       Fire_RemovingItem, Fire_InsertedItem
	/// \endif
	HRESULT Fire_RemovedItem(IVirtualTreeViewItem* pTreeItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pTreeItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_REMOVEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovingItem event</em>
	///
	/// \param[in] pTreeItem The item being removed.
	/// \param[in,out] pCancelDeletion If \c VARIANT_TRUE, the caller should abort deletion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::RemovingItem, ExplorerTreeView::Raise_RemovingItem,
	///       Fire_RemovedItem, Fire_InsertingItem
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::RemovingItem, ExplorerTreeView::Raise_RemovingItem,
	///       Fire_RemovedItem, Fire_InsertingItem
	/// \endif
	HRESULT Fire_RemovingItem(ITreeViewItem* pTreeItem, VARIANT_BOOL* pCancelDeletion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pTreeItem;
				p[0].pboolVal = pCancelDeletion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_REMOVINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RenamedItem event</em>
	///
	/// \param[in] pTreeItem The renamed item.
	/// \param[in] previousItemText The item's previous text.
	/// \param[in] newItemText The item's new text.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::RenamedItem, ExplorerTreeView::Raise_RenamedItem,
	///       Fire_RenamingItem, Fire_StartingLabelEditing, Fire_DestroyedEditControlWindow
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::RenamedItem, ExplorerTreeView::Raise_RenamedItem,
	///       Fire_RenamingItem, Fire_StartingLabelEditing, Fire_DestroyedEditControlWindow
	/// \endif
	HRESULT Fire_RenamedItem(ITreeViewItem* pTreeItem, BSTR previousItemText, BSTR newItemText)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pTreeItem;
				p[1] = previousItemText;
				p[0] = newItemText;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_RENAMEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RenamingItem event</em>
	///
	/// \param[in] pTreeItem The item being renamed.
	/// \param[in] previousItemText The item's previous text.
	/// \param[in] newItemText The item's new text.
	/// \param[in,out] pCancelRenaming If \c VARIANT_TRUE, the caller should abort renaming; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::RenamingItem, ExplorerTreeView::Raise_RenamingItem,
	///       Fire_RenamedItem, Fire_StartingLabelEditing, Fire_DestroyedEditControlWindow
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::RenamingItem, ExplorerTreeView::Raise_RenamingItem,
	///       Fire_RenamedItem, Fire_StartingLabelEditing, Fire_DestroyedEditControlWindow
	/// \endif
	HRESULT Fire_RenamingItem(ITreeViewItem* pTreeItem, BSTR previousItemText, BSTR newItemText, VARIANT_BOOL* pCancelRenaming)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pTreeItem;
				p[2] = previousItemText;
				p[1] = newItemText;
				p[0].pboolVal = pCancelRenaming;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_RENAMINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::ResizedControlWindow,
	///       ExplorerTreeView::Raise_ResizedControlWindow
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::ResizedControlWindow,
	///       ExplorerTreeView::Raise_ResizedControlWindow
	/// \endif
	HRESULT Fire_ResizedControlWindow(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_RESIZEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SingleExpanding event</em>
	///
	/// \param[in] pPreviousCaretItem The previous caret item.
	/// \param[in] pNewCaretItem The new caret item.
	/// \param[in,out] pDontChangePreviousItem If \c VARIANT_TRUE, the caller shouldn't change the
	///                previous caret item's expansion state; otherwise it should.
	/// \param[in,out] pDontChangeNewItem If \c VARIANT_TRUE, the caller shouldn't change the new caret
	///                item's expansion state; otherwise it should.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::SingleExpanding,
	///       ExplorerTreeView::Raise_SingleExpanding, Fire_CaretChanging
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::SingleExpanding,
	///       ExplorerTreeView::Raise_SingleExpanding, Fire_CaretChanging
	/// \endif
	HRESULT Fire_SingleExpanding(ITreeViewItem* pPreviousCaretItem, ITreeViewItem* pNewCaretItem, VARIANT_BOOL* pDontChangePreviousItem, VARIANT_BOOL* pDontChangeNewItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pPreviousCaretItem;
				p[2] = pNewCaretItem;
				p[1].pboolVal = pDontChangePreviousItem;		p[1].vt = VT_BOOL | VT_BYREF;
				p[0].pboolVal = pDontChangeNewItem;					p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_SINGLEEXPANDING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c StartingLabelEditing event</em>
	///
	/// \param[in] pTreeItem The item being edited.
	/// \param[in,out] pCancelEditing If \c VARIANT_TRUE, the caller should prevent label-editing;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::StartingLabelEditing,
	///       ExplorerTreeView::Raise_StartingLabelEditing, Fire_RenamingItem, Fire_CreatedEditControlWindow
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::StartingLabelEditing,
	///       ExplorerTreeView::Raise_StartingLabelEditing, Fire_RenamingItem, Fire_CreatedEditControlWindow
	/// \endif
	HRESULT Fire_StartingLabelEditing(ITreeViewItem* pTreeItem, VARIANT_BOOL* pCancelEditing)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pTreeItem;
				p[0].pboolVal = pCancelEditing;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_STARTINGLABELEDITING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XClick event</em>
	///
	/// \param[in] pTreeItem The clicked item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be a
	///            constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::XClick, ExplorerTreeView::Raise_XClick, Fire_XDblClick,
	///       Fire_Click, Fire_MClick, Fire_RClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::XClick, ExplorerTreeView::Raise_XClick, Fire_XDblClick,
	///       Fire_Click, Fire_MClick, Fire_RClick
	/// \endif
	HRESULT Fire_XClick(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_XCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XDblClick event</em>
	///
	/// \param[in] pTreeItem The double-clicked item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::_IExplorerTreeViewEvents::XDblClick, ExplorerTreeView::Raise_XDblClick,
	///       Fire_XClick, Fire_DblClick, Fire_MDblClick, Fire_RDblClick
	/// \else
	///   \sa ExTVwLibA::_IExplorerTreeViewEvents::XDblClick, ExplorerTreeView::Raise_XDblClick,
	///       Fire_XClick, Fire_DblClick, Fire_MDblClick, Fire_RDblClick
	/// \endif
	HRESULT Fire_XDblClick(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pTreeItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXTVWE_XDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}
};     // Proxy_IExplorerTreeViewEvents
