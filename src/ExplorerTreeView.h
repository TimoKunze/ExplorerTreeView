//////////////////////////////////////////////////////////////////////
/// \class ExplorerTreeView
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Superclasses \c SysTreeView32</em>
///
/// This class superclasses \c SysTreeView32 and makes it accessible by COM.
///
/// \todo Move the OLE drag'n'drop flags into their own struct?
/// \todo Separate \c MouseStatus for the treeview and the label-edit control.
/// \todo Support \c BlendSelectedItemsIcons together with \c ShowStateImages.
/// \todo Support \c BlendSelectedItemsIcons together with \c ItemYBorder.
/// \todo \c IMEFlags is the name of a struct as well as a variable.
/// \todo Verify documentation of \c PreTranslateAccelerator.
///
/// \if UNICODE
///   \sa ExTVwLibU::IExplorerTreeView, TreeViewItem, TreeViewItems, TreeViewItemContainer,
///       VirtualTreeViewItem
/// \else
///   \sa ExTVwLibA::IExplorerTreeView, TreeViewItem, TreeViewItems, TreeViewItemContainer,
///       VirtualTreeViewItem
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "CompilerFlags.h"
#include "res/resource.h"
#ifdef UNICODE
	#include "ExTVwU.h"
#else
	#include "ExTVwA.h"
#endif
#include "CWindowEx.h"
#include "_IExplorerTreeViewEvents_CP.h"
#include "ICategorizeProperties.h"
#include "ICreditsProvider.h"
#include "IItemComparator.h"
#include "UndocComctl32.h"
#include "helpers.h"
#include "EnumOLEVERB.h"
#include "PropertyNotifySinkImpl.h"
#include "CaretChangeTasks.h"
#include "CCEnsureVisibleTask.h"
#include "AboutDlg.h"
#include "CommonProperties.h"
#include "TreeViewItem.h"
#include "TreeViewItems.h"
#include "TreeViewItemContainer.h"
#include "VirtualTreeViewItem.h"
#include "TargetOLEDataObject.h"
#include "SourceOLEDataObject.h"

#ifdef INCLUDESHELLBROWSERINTERFACE
	#include "shbrowser/definitions.h"
	#include "shbrowser/IMessageListener.h"
	#include "shbrowser/IInternalMessageListener.h"
#endif


class ATL_NO_VTABLE ExplorerTreeView : 
    public CComObjectRootEx<CComSingleThreadModel>,
    #ifdef UNICODE
    	public IDispatchImpl<IExplorerTreeView, &IID_IExplorerTreeView, &LIBID_ExTVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IExplorerTreeView, &IID_IExplorerTreeView, &LIBID_ExTVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<ExplorerTreeView>,
    public IOleControlImpl<ExplorerTreeView>,
    public IOleObjectImpl<ExplorerTreeView>,
    public IOleInPlaceActiveObjectImpl<ExplorerTreeView>,
    public IViewObjectExImpl<ExplorerTreeView>,
    public IOleInPlaceObjectWindowlessImpl<ExplorerTreeView>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ExplorerTreeView>,
    public Proxy_IExplorerTreeViewEvents<ExplorerTreeView>,
    public IPersistStorageImpl<ExplorerTreeView>,
    public IPersistPropertyBagImpl<ExplorerTreeView>,
    public ISpecifyPropertyPages,
    public IQuickActivateImpl<ExplorerTreeView>,
    #ifdef UNICODE
    	public IProvideClassInfo2Impl<&CLSID_ExplorerTreeView, &__uuidof(_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IProvideClassInfo2Impl<&CLSID_ExplorerTreeView, &__uuidof(_IExplorerTreeViewEvents), &LIBID_ExTVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPropertyNotifySinkCP<ExplorerTreeView>,
    public CComCoClass<ExplorerTreeView, &CLSID_ExplorerTreeView>,
    public CComControl<ExplorerTreeView>,
    public IPerPropertyBrowsingImpl<ExplorerTreeView>,
    public IDropTarget,
    public IDropSource,
    public IDropSourceNotify,
    #ifdef INCLUDESHELLBROWSERINTERFACE
    	public IInternalMessageListener,
    #endif
    public ICategorizeProperties,
    public ICreditsProvider,
    public IItemComparator
{
	friend class TreeViewItem;
	friend class TreeViewItems;
	friend class TreeViewItemContainer;
	friend class SourceOLEDataObject;

public:
	/// \brief <em>The contained label-edit control</em>
	CContainedWindow containedEdit;

	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	ExplorerTreeView();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_OLEMISC_STATUS(OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_ALIGNABLE | OLEMISC_INSIDEOUT | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST)
		DECLARE_REGISTRY_RESOURCEID(IDR_EXPLORERTREEVIEW)

		#ifdef UNICODE
			DECLARE_WND_SUPERCLASS(TEXT("ExplorerTreeViewU"), WC_TREEVIEWW)
		#else
			DECLARE_WND_SUPERCLASS(TEXT("ExplorerTreeViewA"), WC_TREEVIEWA)
		#endif

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

		BEGIN_COM_MAP(ExplorerTreeView)
			COM_INTERFACE_ENTRY(IExplorerTreeView)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(IViewObjectEx)
			COM_INTERFACE_ENTRY(IViewObject2)
			COM_INTERFACE_ENTRY(IViewObject)
			COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceObject)
			COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
			COM_INTERFACE_ENTRY(IOleControl)
			COM_INTERFACE_ENTRY(IOleObject)
			COM_INTERFACE_ENTRY(IPersistStreamInit)
			COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IPersistPropertyBag)
			COM_INTERFACE_ENTRY(IQuickActivate)
			COM_INTERFACE_ENTRY(IPersistStorage)
			COM_INTERFACE_ENTRY(IProvideClassInfo)
			COM_INTERFACE_ENTRY(IProvideClassInfo2)
			COM_INTERFACE_ENTRY_IID(IID_ICategorizeProperties, ICategorizeProperties)
			COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
			COM_INTERFACE_ENTRY(IPerPropertyBrowsing)
			COM_INTERFACE_ENTRY(IDropTarget)
			COM_INTERFACE_ENTRY(IDropSource)
			COM_INTERFACE_ENTRY(IDropSourceNotify)
		END_COM_MAP()

		BEGIN_PROP_MAP(ExplorerTreeView)
			// NOTE: Don't forget to update Load and Save! This is for property bags only, not for streams!
			PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
			PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
			PROP_ENTRY_TYPE("AllowDragDrop", DISPID_EXTVW_ALLOWDRAGDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AllowLabelEditing", DISPID_EXTVW_ALLOWLABELEDITING, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AlwaysShowSelection", DISPID_EXTVW_ALWAYSSHOWSELECTION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Appearance", DISPID_EXTVW_APPEARANCE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("AutoHScroll", DISPID_EXTVW_AUTOHSCROLL, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AutoHScrollPixelsPerSecond", DISPID_EXTVW_AUTOHSCROLLPIXELSPERSECOND, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("AutoHScrollRedrawInterval", DISPID_EXTVW_AUTOHSCROLLREDRAWINTERVAL, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("BackColor", DISPID_EXTVW_BACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("BlendSelectedItemsIcons", DISPID_EXTVW_BLENDSELECTEDITEMSICONS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("BorderStyle", DISPID_EXTVW_BORDERSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("BuiltInStateImages", DISPID_EXTVW_BUILTINSTATEIMAGES, CLSID_CommonProperties, VT_I4)
			PROP_ENTRY_TYPE("CaretChangedDelayTime", DISPID_EXTVW_CARETCHANGEDDELAYTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DisabledEvents", DISPID_EXTVW_DISABLEDEVENTS, CLSID_CommonProperties, VT_I4)
			PROP_ENTRY_TYPE("DontRedraw", DISPID_EXTVW_DONTREDRAW, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DragExpandTime", DISPID_EXTVW_DRAGEXPANDTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DragScrollTimeBase", DISPID_EXTVW_DRAGSCROLLTIMEBASE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DrawImagesAsynchronously", DISPID_EXTVW_DRAWIMAGESASYNCHRONOUSLY, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("EditBackColor", DISPID_EXTVW_EDITBACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("EditForeColor", DISPID_EXTVW_EDITFORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("EditHoverTime", DISPID_EXTVW_EDITHOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("EditIMEMode", DISPID_EXTVW_EDITIMEMODE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Enabled", DISPID_EXTVW_ENABLED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("FadeExpandos", DISPID_EXTVW_FADEEXPANDOS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("FavoritesStyle", DISPID_EXTVW_FAVORITESSTYLE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Font", DISPID_EXTVW_FONT, CLSID_StockFontPage, VT_DISPATCH)
			PROP_ENTRY_TYPE("ForeColor", DISPID_EXTVW_FORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("FullRowSelect", DISPID_EXTVW_FULLROWSELECT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("GroupBoxColor", DISPID_EXTVW_GROUPBOXCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("HotTracking", DISPID_EXTVW_HOTTRACKING, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("HoverTime", DISPID_EXTVW_HOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IMEMode", DISPID_EXTVW_IMEMODE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Indent", DISPID_EXTVW_INDENT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IndentStateImages", DISPID_EXTVW_INDENTSTATEIMAGES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("InsertMarkColor", DISPID_EXTVW_INSERTMARKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("ItemBoundingBoxDefinition", DISPID_EXTVW_ITEMBOUNDINGBOXDEFINITION, CLSID_CommonProperties, VT_I4)
			PROP_ENTRY_TYPE("ItemHeight", DISPID_EXTVW_ITEMHEIGHT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ItemXBorder", DISPID_EXTVW_ITEMXBORDER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ItemYBorder", DISPID_EXTVW_ITEMYBORDER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("LineColor", DISPID_EXTVW_LINECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("LineStyle", DISPID_EXTVW_LINESTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MaxScrollTime", DISPID_EXTVW_MAXSCROLLTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MouseIcon", DISPID_EXTVW_MOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("MousePointer", DISPID_EXTVW_MOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("OLEDragImageStyle", DISPID_EXTVW_OLEDRAGIMAGESTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ProcessContextMenuKeys", DISPID_EXTVW_PROCESSCONTEXTMENUKEYS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RegisterForOLEDragDrop", DISPID_EXTVW_REGISTERFOROLEDRAGDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RichToolTips", DISPID_EXTVW_RICHTOOLTIPS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RightToLeft", DISPID_EXTVW_RIGHTTOLEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ScrollBars", DISPID_EXTVW_SCROLLBARS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ShowStateImages", DISPID_EXTVW_SHOWSTATEIMAGES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ShowToolTips", DISPID_EXTVW_SHOWTOOLTIPS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("SingleExpand", DISPID_EXTVW_SINGLEEXPAND, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("SortOrder", DISPID_EXTVW_SORTORDER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("SupportOLEDragImages", DISPID_EXTVW_SUPPORTOLEDRAGIMAGES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("TreeViewStyle", DISPID_EXTVW_TREEVIEWSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("UseSystemFont", DISPID_EXTVW_USESYSTEMFONT, CLSID_NULL, VT_BOOL)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(ExplorerTreeView)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_IExplorerTreeViewEvents))
		END_CONNECTION_POINT_MAP()

		#ifdef INCLUDESHELLBROWSERINTERFACE
			BEGIN_FILTERED_MSG_MAP(ExplorerTreeView, shellBrowserInterface.pMessageListener)
				FILTERED_MESSAGE_HANDLER(WM_CHAR, OnChar)
				FILTERED_MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
				FILTERED_MESSAGE_HANDLER(WM_CREATE, OnCreate)
				FILTERED_MESSAGE_HANDLER(WM_CTLCOLOREDIT, OnCtlColorEdit)
				FILTERED_MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
				FILTERED_MESSAGE_HANDLER(WM_INPUTLANGCHANGE, OnInputLangChange)
				FILTERED_MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
				FILTERED_MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
				FILTERED_MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
				FILTERED_MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
				FILTERED_MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
				FILTERED_MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
				FILTERED_MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
				FILTERED_MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
				FILTERED_MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
				FILTERED_MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
				FILTERED_MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnMouseWheel)
				FILTERED_MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
				FILTERED_MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
				FILTERED_MESSAGE_HANDLER(WM_MOUSEWHEEL, OnMouseWheel)
				FILTERED_MESSAGE_HANDLER(WM_NOTIFY, OnNotify)
				FILTERED_MESSAGE_HANDLER(WM_PAINT, OnPaint)
				FILTERED_MESSAGE_HANDLER(WM_PARENTNOTIFY, OnParentNotify)
				FILTERED_MESSAGE_HANDLER(WM_PRINTCLIENT, OnPaint)
				FILTERED_MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
				FILTERED_MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
				FILTERED_MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
				FILTERED_MESSAGE_HANDLER(WM_SETFONT, OnSetFont)
				FILTERED_MESSAGE_HANDLER(WM_SETREDRAW, OnSetRedraw)
				FILTERED_MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
				FILTERED_MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
				FILTERED_MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
				FILTERED_MESSAGE_HANDLER(WM_THEMECHANGED, OnThemeChanged)
				FILTERED_MESSAGE_HANDLER(WM_TIMER, OnTimer)
				FILTERED_MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
				FILTERED_MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
				FILTERED_MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
				FILTERED_MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)

				FILTERED_MESSAGE_HANDLER(OCM_NOTIFY, OnReflectedNotify)
				FILTERED_MESSAGE_HANDLER(OCM__BASE + WM_NOTIFYFORMAT, OnReflectedNotifyFormat)

				FILTERED_MESSAGE_HANDLER(SHBM_HANDSHAKE, OnHandshake)

				FILTERED_MESSAGE_HANDLER(GetDragImageMessage(), OnGetDragImage)

				FILTERED_MESSAGE_HANDLER(TVM_CREATEDRAGIMAGE, OnCreateDragImage)
				FILTERED_MESSAGE_HANDLER(TVM_DELETEITEM, OnDeleteItem)
				FILTERED_MESSAGE_HANDLER(TVM_EDITLABELA, OnEditLabel)
				FILTERED_MESSAGE_HANDLER(TVM_EDITLABELW, OnEditLabel)
				FILTERED_MESSAGE_HANDLER(TVM_EXPAND, OnExpandItem)
				FILTERED_MESSAGE_HANDLER(TVM_GETITEMA, OnGetItem)
				FILTERED_MESSAGE_HANDLER(TVM_GETITEMW, OnGetItem)
				FILTERED_MESSAGE_HANDLER(TVM_INSERTITEMA, OnInsertItem)
				FILTERED_MESSAGE_HANDLER(TVM_INSERTITEMW, OnInsertItem)
				FILTERED_MESSAGE_HANDLER(TVM_SELECTITEM, OnSelectItem)
				FILTERED_MESSAGE_HANDLER(TVM_SETAUTOSCROLLINFO, OnSetAutoScrollInfo)
				FILTERED_MESSAGE_HANDLER(TVM_SETBORDER, OnSetBorder)
				FILTERED_MESSAGE_HANDLER(TVM_SETIMAGELIST, OnSetImageList)
				FILTERED_MESSAGE_HANDLER(TVM_SETINSERTMARK, OnSetInsertMark)
				FILTERED_MESSAGE_HANDLER(TVM_SETITEMA, OnSetItem)
				FILTERED_MESSAGE_HANDLER(TVM_SETITEMW, OnSetItem)
				FILTERED_MESSAGE_HANDLER(TVM_SETITEMHEIGHT, OnSetItemHeight)

				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(NM_CLICK, OnClickNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(NM_DBLCLK, OnDblClkNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(NM_RCLICK, OnRClickNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(NM_RDBLCLK, OnRDblClkNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(NM_SETCURSOR, OnSetCursorNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(NM_SETFOCUS, OnSetFocusNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(NM_TVSTATEIMAGECHANGING, OnTVStateImageChangingNotification)

				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_ASYNCDRAW, OnAsyncDrawNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_BEGINDRAGA, OnBeginDragNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_BEGINDRAGW, OnBeginDragNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_BEGINLABELEDITA, OnBeginLabelEditNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_BEGINLABELEDITW, OnBeginLabelEditNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_BEGINRDRAGA, OnBeginRDragNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_BEGINRDRAGW, OnBeginRDragNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_DELETEITEMA, OnDeleteItemNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_DELETEITEMW, OnDeleteItemNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_ENDLABELEDITA, OnEndLabelEditNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_ENDLABELEDITW, OnEndLabelEditNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_GETDISPINFOA, OnGetDispInfoNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_GETDISPINFOW, OnGetDispInfoNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_GETINFOTIPA, OnGetInfoTipNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_GETINFOTIPW, OnGetInfoTipNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMCHANGEDA, OnItemChangedNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMCHANGEDW, OnItemChangedNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMCHANGINGA, OnItemChangingNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMCHANGINGW, OnItemChangingNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMEXPANDEDA, OnItemExpandedNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMEXPANDEDW, OnItemExpandedNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMEXPANDINGA, OnItemExpandingNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMEXPANDINGW, OnItemExpandingNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_KEYDOWN, OnKeyDownNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_SELCHANGEDA, OnCaretChangedNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_SELCHANGEDW, OnCaretChangedNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_SELCHANGINGA, OnCaretChangingNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_SELCHANGINGW, OnCaretChangingNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_SETDISPINFOA, OnSetDispInfoNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_SETDISPINFOW, OnSetDispInfoNotification)
				FILTERED_REFLECTED_NOTIFY_CODE_HANDLER(TVN_SINGLEEXPAND, OnSingleExpandNotification)

				FILTERED_COMMAND_CODE_HANDLER(EN_CHANGE, OnEditChange)

				FILTERED_CHAIN_MSG_MAP(CComControl<ExplorerTreeView>)
				FILTERED_ALT_MSG_MAP(1)
				FILTERED_MESSAGE_HANDLER(WM_CHAR, OnEditChar)
				FILTERED_MESSAGE_HANDLER(WM_CONTEXTMENU, OnEditContextMenu)
				FILTERED_MESSAGE_HANDLER(WM_DESTROY, OnEditDestroy)
				FILTERED_MESSAGE_HANDLER(WM_KEYDOWN, OnEditKeyDown)
				FILTERED_MESSAGE_HANDLER(WM_KEYUP, OnEditKeyUp)
				FILTERED_MESSAGE_HANDLER(WM_KILLFOCUS, OnEditKillFocus)
				FILTERED_MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnEditLButtonDblClk)
				FILTERED_MESSAGE_HANDLER(WM_LBUTTONDOWN, OnEditLButtonDown)
				FILTERED_MESSAGE_HANDLER(WM_LBUTTONUP, OnEditLButtonUp)
				FILTERED_MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnEditMButtonDblClk)
				FILTERED_MESSAGE_HANDLER(WM_MBUTTONDOWN, OnEditMButtonDown)
				FILTERED_MESSAGE_HANDLER(WM_MBUTTONUP, OnEditMButtonUp)
				FILTERED_MESSAGE_HANDLER(WM_MOUSEHOVER, OnEditMouseHover)
				FILTERED_MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnEditMouseWheel)
				FILTERED_MESSAGE_HANDLER(WM_MOUSELEAVE, OnEditMouseLeave)
				FILTERED_MESSAGE_HANDLER(WM_MOUSEMOVE, OnEditMouseMove)
				FILTERED_MESSAGE_HANDLER(WM_MOUSEWHEEL, OnEditMouseWheel)
				FILTERED_MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnEditRButtonDblClk)
				FILTERED_MESSAGE_HANDLER(WM_RBUTTONDOWN, OnEditRButtonDown)
				FILTERED_MESSAGE_HANDLER(WM_RBUTTONUP, OnEditRButtonUp)
				FILTERED_MESSAGE_HANDLER(WM_SETFOCUS, OnEditSetFocus)
				FILTERED_MESSAGE_HANDLER(WM_SYSKEYDOWN, OnEditKeyDown)
				FILTERED_MESSAGE_HANDLER(WM_SYSKEYUP, OnEditKeyUp)
				FILTERED_MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnEditXButtonDblClk)
				FILTERED_MESSAGE_HANDLER(WM_XBUTTONDOWN, OnEditXButtonDown)
				FILTERED_MESSAGE_HANDLER(WM_XBUTTONUP, OnEditXButtonUp)
			END_FILTERED_MSG_MAP(shellBrowserInterface.pMessageListener)
		#else
			BEGIN_MSG_MAP(ExplorerTreeView)
				MESSAGE_HANDLER(WM_CHAR, OnChar)
				MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
				MESSAGE_HANDLER(WM_CREATE, OnCreate)
				MESSAGE_HANDLER(WM_CTLCOLOREDIT, OnCtlColorEdit)
				MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
				MESSAGE_HANDLER(WM_INPUTLANGCHANGE, OnInputLangChange)
				MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
				MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
				MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
				MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
				MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
				MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
				MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
				MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
				MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
				MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
				MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnMouseWheel)
				MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
				MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
				MESSAGE_HANDLER(WM_MOUSEWHEEL, OnMouseWheel)
				MESSAGE_HANDLER(WM_NOTIFY, OnNotify)
				MESSAGE_HANDLER(WM_PAINT, OnPaint)
				MESSAGE_HANDLER(WM_PARENTNOTIFY, OnParentNotify)
				MESSAGE_HANDLER(WM_PRINTCLIENT, OnPaint)
				MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
				MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
				MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
				MESSAGE_HANDLER(WM_SETFONT, OnSetFont)
				MESSAGE_HANDLER(WM_SETREDRAW, OnSetRedraw)
				MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
				MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
				MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
				MESSAGE_HANDLER(WM_THEMECHANGED, OnThemeChanged)
				MESSAGE_HANDLER(WM_TIMER, OnTimer)
				MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
				MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
				MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
				MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)

				MESSAGE_HANDLER(OCM_NOTIFY, OnReflectedNotify)
				MESSAGE_HANDLER(OCM__BASE + WM_NOTIFYFORMAT, OnReflectedNotifyFormat)

				MESSAGE_HANDLER(GetDragImageMessage(), OnGetDragImage)

				MESSAGE_HANDLER(TVM_CREATEDRAGIMAGE, OnCreateDragImage)
				MESSAGE_HANDLER(TVM_DELETEITEM, OnDeleteItem)
				MESSAGE_HANDLER(TVM_EDITLABELA, OnEditLabel)
				MESSAGE_HANDLER(TVM_EDITLABELW, OnEditLabel)
				MESSAGE_HANDLER(TVM_EXPAND, OnExpandItem)
				MESSAGE_HANDLER(TVM_GETITEMA, OnGetItem)
				MESSAGE_HANDLER(TVM_GETITEMW, OnGetItem)
				MESSAGE_HANDLER(TVM_INSERTITEMA, OnInsertItem)
				MESSAGE_HANDLER(TVM_INSERTITEMW, OnInsertItem)
				MESSAGE_HANDLER(TVM_SELECTITEM, OnSelectItem)
				MESSAGE_HANDLER(TVM_SETAUTOSCROLLINFO, OnSetAutoScrollInfo)
				MESSAGE_HANDLER(TVM_SETBORDER, OnSetBorder)
				MESSAGE_HANDLER(TVM_SETIMAGELIST, OnSetImageList)
				MESSAGE_HANDLER(TVM_SETINSERTMARK, OnSetInsertMark)
				MESSAGE_HANDLER(TVM_SETITEMA, OnSetItem)
				MESSAGE_HANDLER(TVM_SETITEMW, OnSetItem)
				MESSAGE_HANDLER(TVM_SETITEMHEIGHT, OnSetItemHeight)

				REFLECTED_NOTIFY_CODE_HANDLER(NM_CLICK, OnClickNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(NM_DBLCLK, OnDblClkNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(NM_RCLICK, OnRClickNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(NM_RDBLCLK, OnRDblClkNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(NM_SETCURSOR, OnSetCursorNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(NM_SETFOCUS, OnSetFocusNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(NM_TVSTATEIMAGECHANGING, OnTVStateImageChangingNotification)

				REFLECTED_NOTIFY_CODE_HANDLER(TVN_ASYNCDRAW, OnAsyncDrawNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_BEGINDRAGA, OnBeginDragNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_BEGINDRAGW, OnBeginDragNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_BEGINLABELEDITA, OnBeginLabelEditNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_BEGINLABELEDITW, OnBeginLabelEditNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_BEGINRDRAGA, OnBeginRDragNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_BEGINRDRAGW, OnBeginRDragNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_DELETEITEMA, OnDeleteItemNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_DELETEITEMW, OnDeleteItemNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_ENDLABELEDITA, OnEndLabelEditNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_ENDLABELEDITW, OnEndLabelEditNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_GETDISPINFOA, OnGetDispInfoNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_GETDISPINFOW, OnGetDispInfoNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_GETINFOTIPA, OnGetInfoTipNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_GETINFOTIPW, OnGetInfoTipNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMCHANGEDA, OnItemChangedNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMCHANGEDW, OnItemChangedNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMCHANGINGA, OnItemChangingNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMCHANGINGW, OnItemChangingNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMEXPANDEDA, OnItemExpandedNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMEXPANDEDW, OnItemExpandedNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMEXPANDINGA, OnItemExpandingNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_ITEMEXPANDINGW, OnItemExpandingNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_KEYDOWN, OnKeyDownNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_SELCHANGEDA, OnCaretChangedNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_SELCHANGEDW, OnCaretChangedNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_SELCHANGINGA, OnCaretChangingNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_SELCHANGINGW, OnCaretChangingNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_SETDISPINFOA, OnSetDispInfoNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_SETDISPINFOW, OnSetDispInfoNotification)
				REFLECTED_NOTIFY_CODE_HANDLER(TVN_SINGLEEXPAND, OnSingleExpandNotification)

				COMMAND_CODE_HANDLER(EN_CHANGE, OnEditChange)

				CHAIN_MSG_MAP(CComControl<ExplorerTreeView>)
				ALT_MSG_MAP(1)
				MESSAGE_HANDLER(WM_CHAR, OnEditChar)
				MESSAGE_HANDLER(WM_CONTEXTMENU, OnEditContextMenu)
				MESSAGE_HANDLER(WM_DESTROY, OnEditDestroy)
				MESSAGE_HANDLER(WM_KEYDOWN, OnEditKeyDown)
				MESSAGE_HANDLER(WM_KEYUP, OnEditKeyUp)
				MESSAGE_HANDLER(WM_KILLFOCUS, OnEditKillFocus)
				MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnEditLButtonDblClk)
				MESSAGE_HANDLER(WM_LBUTTONDOWN, OnEditLButtonDown)
				MESSAGE_HANDLER(WM_LBUTTONUP, OnEditLButtonUp)
				MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnEditMButtonDblClk)
				MESSAGE_HANDLER(WM_MBUTTONDOWN, OnEditMButtonDown)
				MESSAGE_HANDLER(WM_MBUTTONUP, OnEditMButtonUp)
				MESSAGE_HANDLER(WM_MOUSEHOVER, OnEditMouseHover)
				MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnEditMouseWheel)
				MESSAGE_HANDLER(WM_MOUSELEAVE, OnEditMouseLeave)
				MESSAGE_HANDLER(WM_MOUSEMOVE, OnEditMouseMove)
				MESSAGE_HANDLER(WM_MOUSEWHEEL, OnEditMouseWheel)
				MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnEditRButtonDblClk)
				MESSAGE_HANDLER(WM_RBUTTONDOWN, OnEditRButtonDown)
				MESSAGE_HANDLER(WM_RBUTTONUP, OnEditRButtonUp)
				MESSAGE_HANDLER(WM_SETFOCUS, OnEditSetFocus)
				MESSAGE_HANDLER(WM_SYSKEYDOWN, OnEditKeyDown)
				MESSAGE_HANDLER(WM_SYSKEYUP, OnEditKeyUp)
				MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnEditXButtonDblClk)
				MESSAGE_HANDLER(WM_XBUTTONDOWN, OnEditXButtonDown)
				MESSAGE_HANDLER(WM_XBUTTONUP, OnEditXButtonUp)
			END_MSG_MAP()
		#endif
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
	/// \name Implementation of persistance
	///
	//@{
	/// \brief <em>Overrides \c IPersistStreamInitImpl::GetSizeMax to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we communicate directly with the stream. This requires we override
	/// \c IPersistStreamInitImpl::GetSizeMax.
	///
	/// \param[in] pSize The maximum number of bytes that persistence of the control's properties will
	///            consume.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687287.aspx">IPersistStreamInit::GetSizeMax</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	virtual HRESULT STDMETHODCALLTYPE GetSizeMax(ULARGE_INTEGER* pSize);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Load to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Load and read directly from
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680730.aspx">IPersistStreamInit::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPSTREAM pStream);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Save to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Save and write directly into
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694439.aspx">IPersistStreamInit::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPSTREAM pStream, BOOL clearDirtyFlag);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IExplorerTreeView
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AllowDragDrop property</em>
	///
	/// Retrieves whether drag'n'drop mode can be entered. If set to \c VARIANT_TRUE, drag'n'drop mode
	/// can be entered by pressing the left or right mouse button over an item and then moving the
	/// mouse with the button still pressed. If set to \c VARIANT_FALSE, drag'n'drop mode is not
	/// available - this also means the \c TVN_BEGINDRAG and \c TVN_BEGINRDRAG notifications are not
	/// sent.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AllowDragDrop, get_RegisterForOLEDragDrop, get_DragExpandTime, get_DragScrollTimeBase,
	///     SetInsertMarkPosition, Raise_ItemBeginDrag, Raise_ItemBeginRDrag
	virtual HRESULT STDMETHODCALLTYPE get_AllowDragDrop(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowDragDrop property</em>
	///
	/// Sets whether drag'n'drop mode can be entered. If set to \c VARIANT_TRUE, drag'n'drop mode
	/// can be entered by pressing the left or right mouse button over an item and then moving the
	/// mouse with the button still pressed. If set to \c VARIANT_FALSE, drag'n'drop mode is not
	/// available - this also means the \c TVN_BEGINDRAG and \c TVN_BEGINRDRAG notifications are not
	/// sent.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AllowDragDrop, put_RegisterForOLEDragDrop, put_DragExpandTime, put_DragScrollTimeBase,
	///     SetInsertMarkPosition, Raise_ItemBeginDrag, Raise_ItemBeginRDrag
	virtual HRESULT STDMETHODCALLTYPE put_AllowDragDrop(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AllowLabelEditing property</em>
	///
	/// Retrieves whether label-editing mode can be entered. If set to \c VARIANT_TRUE, label-edit mode
	/// can be entered by either calling \c TreeViewItem::StartLabelEditing or by single-clicking the
	/// control's caret item. If set to \c VARIANT_FALSE, label-edit mode is not available.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AllowLabelEditing, get_hWndEdit, TreeViewItem::StartLabelEditing,
	///     Raise_StartingLabelEditing
	virtual HRESULT STDMETHODCALLTYPE get_AllowLabelEditing(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowLabelEditing property</em>
	///
	/// Sets whether label-editing mode can be entered. If set to \c VARIANT_TRUE, label-edit mode
	/// can be entered by either calling \c TreeViewItem::StartLabelEditing or by single-clicking the
	/// control's caret item. If set to \c VARIANT_FALSE, label-edit mode is not available.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AllowLabelEditing, get_hWndEdit, TreeViewItem::StartLabelEditing,
	///     Raise_StartingLabelEditing
	virtual HRESULT STDMETHODCALLTYPE put_AllowLabelEditing(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AlwaysShowSelection property</em>
	///
	/// Retrieves whether the selected items will be highlighted even if the control doesn't have the
	/// focus. If set to \c VARIANT_TRUE, selected items are drawn as selected if the control does not
	/// have the focus; otherwise they're drawn as normal items.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AlwaysShowSelection, get_CaretItem, TreeViewItem::get_Selected
	virtual HRESULT STDMETHODCALLTYPE get_AlwaysShowSelection(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AlwaysShowSelection property</em>
	///
	/// Sets whether the selected items will be highlighted even if the control doesn't have the focus.
	/// If set to \c VARIANT_TRUE, selected items are drawn as selected if the control does not have the
	/// focus; otherwise they're drawn as normal items.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AlwaysShowSelection, putref_CaretItem, TreeViewItem::get_Selected
	virtual HRESULT STDMETHODCALLTYPE put_AlwaysShowSelection(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Appearance property</em>
	///
	/// Retrieves the kind of border that is drawn around the control. Any of the values defined by the
	/// \c AppearanceConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Appearance, get_BorderStyle, ExTVwLibU::AppearanceConstants
	/// \else
	///   \sa put_Appearance, get_BorderStyle, ExTVwLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Appearance(AppearanceConstants* pValue);
	/// \brief <em>Sets the \c Appearance property</em>
	///
	/// Sets the kind of border that is drawn around the control. Any of the values defined by the
	/// \c AppearanceConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Appearance, put_BorderStyle, ExTVwLibU::AppearanceConstants
	/// \else
	///   \sa get_Appearance, put_BorderStyle, ExTVwLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Appearance(AppearanceConstants newValue);
	/// \brief <em>Retrieves the control's application ID</em>
	///
	/// Retrieves the control's application ID. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppID(SHORT* pValue);
	/// \brief <em>Retrieves the control's application name</em>
	///
	/// Retrieves the control's application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppName(BSTR* pValue);
	/// \brief <em>Retrieves the control's short application name</em>
	///
	/// Retrieves the control's short application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The short application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_Build, get_CharSet, get_IsRelease, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppShortName(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c AutoHScroll property</em>
	///
	/// Retrieves whether the control does automatic horizontal scrolling to scroll the selected item into
	/// view. If set to \c VARIANT_TRUE, the control does automatic scrolling; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_AutoHScroll, get_AutoHScrollPixelsPerSecond, get_AutoHScrollRedrawInterval, get_ScrollBars
	virtual HRESULT STDMETHODCALLTYPE get_AutoHScroll(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AutoHScroll property</em>
	///
	/// Sets whether the control does automatic horizontal scrolling to scroll the selected item into
	/// view. If set to \c VARIANT_TRUE, the control does automatic scrolling; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_AutoHScroll, put_AutoHScrollPixelsPerSecond, put_AutoHScrollRedrawInterval, put_ScrollBars
	virtual HRESULT STDMETHODCALLTYPE put_AutoHScroll(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AutoHScrollPixelsPerSecond property</em>
	///
	/// Retrieves the number of pixels to scroll per second when doing automatic horizontal scrolling. This
	/// value influences the scrolling speed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_AutoHScrollPixelsPerSecond, get_AutoHScroll, get_AutoHScrollRedrawInterval
	virtual HRESULT STDMETHODCALLTYPE get_AutoHScrollPixelsPerSecond(LONG* pValue);
	/// \brief <em>Sets the \c AutoHScrollPixelsPerSecond property</em>
	///
	/// Sets the number of pixels to scroll per second when doing automatic horizontal scrolling. This
	/// value influences the scrolling speed.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_AutoHScrollPixelsPerSecond, put_AutoHScroll, put_AutoHScrollRedrawInterval
	virtual HRESULT STDMETHODCALLTYPE put_AutoHScrollPixelsPerSecond(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c AutoHScrollRedrawInterval property</em>
	///
	/// Retrieves the time (in milliseconds) between two redraws of the control when doing automatic
	/// horizontal scrolling. This value influences the scrolling smoothness. The lower the value, the
	/// smoother the scrolling.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_AutoHScrollRedrawInterval, get_AutoHScroll, get_AutoHScrollPixelsPerSecond
	virtual HRESULT STDMETHODCALLTYPE get_AutoHScrollRedrawInterval(LONG* pValue);
	/// \brief <em>Sets the \c AutoHScrollRedrawInterval property</em>
	///
	/// Sets the time (in milliseconds) between two redraws of the control when doing automatic
	/// horizontal scrolling. This value influences the scrolling smoothness. The lower the value, the
	/// smoother the scrolling.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_AutoHScrollRedrawInterval, put_AutoHScroll, put_AutoHScrollPixelsPerSecond
	virtual HRESULT STDMETHODCALLTYPE put_AutoHScrollRedrawInterval(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c BackColor property</em>
	///
	/// Retrieves the control's background color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_BackColor, get_ForeColor, get_EditBackColor, get_InsertMarkColor, get_LineColor
	virtual HRESULT STDMETHODCALLTYPE get_BackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c BackColor property</em>
	///
	/// Sets the control's background color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_BackColor, put_ForeColor, put_EditBackColor, put_InsertMarkColor, put_LineColor
	virtual HRESULT STDMETHODCALLTYPE put_BackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c BlendSelectedItemsIcons property</em>
	///
	/// Retrieves whether the selected items' icons are blended with the system highlight color.
	/// If set to \c VARIANT_TRUE, the icons are blended; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c VARIANT_TRUE, the \c ShowStateImages property should be set
	///          to \c VARIANT_FALSE and the \c ItemXBorder and \c ItemYBorder properties to 0 to avoid
	///          drawing bugs.\n
	///          If this property is set to \c VARIANT_TRUE, the \c CustomDraw event will be fired for each
	///          single item also if the \c cdrvNotifyItemDraw flag isn't set.
	///
	/// \sa put_BlendSelectedItemsIcons, get_hImageList, get_ItemXBorder, get_ItemYBorder,
	///     get_ShowStateImages, Raise_CustomDraw
	virtual HRESULT STDMETHODCALLTYPE get_BlendSelectedItemsIcons(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c BlendSelectedItemsIcons property</em>
	///
	/// Sets whether the selected items' icons are blended with the system highlight color.
	/// If set to \c VARIANT_TRUE, the icons are blended; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c VARIANT_TRUE, the \c ShowStateImages property should be set
	///          to \c VARIANT_FALSE and the \c ItemXBorder and \c ItemYBorder properties to 0 to avoid
	///          drawing bugs.\n
	///          If this property is set to \c VARIANT_TRUE, the \c CustomDraw event will be fired for each
	///          single item also if the \c cdrvNotifyItemDraw flag isn't set.
	///
	/// \sa get_BlendSelectedItemsIcons, put_hImageList, put_ItemXBorder, put_ItemYBorder,
	///     put_ShowStateImages, Raise_CustomDraw
	virtual HRESULT STDMETHODCALLTYPE put_BlendSelectedItemsIcons(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c BorderStyle property</em>
	///
	/// Retrieves the kind of inner border that is drawn around the control. Any of the values
	/// defined by the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_BorderStyle, get_Appearance, ExTVwLibU::BorderStyleConstants
	/// \else
	///   \sa put_BorderStyle, get_Appearance, ExTVwLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_BorderStyle(BorderStyleConstants* pValue);
	/// \brief <em>Sets the \c BorderStyle property</em>
	///
	/// Sets the kind of inner border that is drawn around the control. Any of the values
	/// defined by the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_BorderStyle, put_Appearance, ExTVwLibU::BorderStyleConstants
	/// \else
	///   \sa get_BorderStyle, put_Appearance, ExTVwLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_BorderStyle(BorderStyleConstants newValue);
	/// \brief <em>Retrieves the control's build number</em>
	///
	/// Retrieves the control's build number. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The build number.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Build(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c BuiltInStateImages property</em>
	///
	/// Retrieves which images the built-in state imagelist does consist of. The built-in state imagelist
	/// is used, if the \c ShowStateImages property is set to \c VARIANT_TRUE and the \c ilState imagelist
	/// is NOT set to a custom imagelist.\n
	/// Any combination of the values defined by the \c BuiltInStateImagesConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property may destroy and recreate the control window.
	///
	/// \if UNICODE
	///   \sa put_BuiltInStateImages, get_ShowStateImages, get_hImageList, get_IndentStateImages,
	///       TreeViewItem::get_StateImageIndex, ExTVwLibU::BuiltInStateImagesConstants,
	///       ExTVwLibU::ImageListConstants
	/// \else
	///   \sa put_BuiltInStateImages, get_ShowStateImages, get_hImageList, get_IndentStateImages,
	///       TreeViewItem::get_StateImageIndex, ExTVwLibA::BuiltInStateImagesConstants,
	///       ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_BuiltInStateImages(BuiltInStateImagesConstants* pValue);
	/// \brief <em>Sets the \c BuiltInStateImages property</em>
	///
	/// Sets which images the built-in state imagelist does consist of. The built-in state imagelist
	/// is used, if the \c ShowStateImages property is set to \c VARIANT_TRUE and the \c ilState imagelist
	/// is NOT set to a custom imagelist.\n
	/// Any combination of the values defined by the \c BuiltInStateImagesConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property may destroy and recreate the control window.
	///
	/// \if UNICODE
	///   \sa get_BuiltInStateImages, put_ShowStateImages, put_hImageList, put_IndentStateImages,
	///       TreeViewItem::put_StateImageIndex, ExTVwLibU::BuiltInStateImagesConstants,
	///       ExTVwLibU::ImageListConstants
	/// \else
	///   \sa get_BuiltInStateImages, put_ShowStateImages, put_hImageList, put_IndentStateImages,
	///       TreeViewItem::put_StateImageIndex, ExTVwLibA::BuiltInStateImagesConstants,
	///       ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_BuiltInStateImages(BuiltInStateImagesConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c CaretChangedDelayTime property</em>
	///
	/// Retrieves the number of milliseconds the control waits before raising the \c CaretChanged event
	/// if the caret item was changed by keyboard input.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CaretChangedDelayTime, get_CaretItem, Raise_CaretChanged
	virtual HRESULT STDMETHODCALLTYPE get_CaretChangedDelayTime(LONG* pValue);
	/// \brief <em>Sets the \c CaretChangedDelayTime property</em>
	///
	/// Sets the number of milliseconds the control waits before raising the \c CaretChanged event
	/// if the caret item was changed by keyboard input.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CaretChangedDelayTime, putref_CaretItem, Raise_CaretChanged
	virtual HRESULT STDMETHODCALLTYPE put_CaretChangedDelayTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c CaretItem property</em>
	///
	/// Retrieves the control's caret item. The caret item is the item that has the focus.
	///
	/// \param[in] noSingleExpand Ignored.
	/// \param[out] ppCaretItem Receives the caret item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_CaretItem, TreeViewItem::get_Caret
	virtual HRESULT STDMETHODCALLTYPE get_CaretItem(VARIANT_BOOL noSingleExpand = VARIANT_TRUE, ITreeViewItem** ppCaretItem = NULL);
	/// \brief <em>Sets the \c CaretItem property</em>
	///
	/// Sets the control's caret item. The caret item is the item that has the focus.
	///
	/// \param[in] noSingleExpand If \c VARIANT_TRUE, the control behaves like it does if the
	///            \c SingleExpand property is set to \c seNone; otherwise the control behaves like
	///            it does if the caret change is done by the user using the mouse.
	/// \param[in] pNewCaretItem The new caret item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CaretItem, put_SingleExpand
	virtual HRESULT STDMETHODCALLTYPE putref_CaretItem(VARIANT_BOOL noSingleExpand = VARIANT_TRUE, ITreeViewItem* pNewCaretItem = NULL);
	/// \brief <em>Retrieves the control's character set</em>
	///
	/// Retrieves the control's character set (Unicode/ANSI). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The character set.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_CharSet(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c DisabledEvents property</em>
	///
	/// Retrieves the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DisabledEvents, ExTVwLibU::DisabledEventsConstants
	/// \else
	///   \sa put_DisabledEvents, ExTVwLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisabledEvents(DisabledEventsConstants* pValue);
	/// \brief <em>Sets the \c DisabledEvents property</em>
	///
	/// Sets the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DisabledEvents, ExTVwLibU::DisabledEventsConstants
	/// \else
	///   \sa get_DisabledEvents, ExTVwLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisabledEvents(DisabledEventsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DontRedraw property</em>
	///
	/// Retrieves whether automatic redrawing of the control is enabled or disabled. If set to
	/// \c VARIANT_FALSE, the control will redraw itself automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DontRedraw
	virtual HRESULT STDMETHODCALLTYPE get_DontRedraw(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DontRedraw property</em>
	///
	/// Enables or disables automatic redrawing of the control. Disabling redraw while doing large changes
	/// on the control (like adding many items) may increase performance. If set to \c VARIANT_FALSE, the
	/// control will redraw itself automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DontRedraw
	virtual HRESULT STDMETHODCALLTYPE put_DontRedraw(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DragExpandTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be placed over an item during a
	/// drag'n'drop operation before this item will be expanded automatically. If set to 0,
	/// auto-expansion is disabled. If set to -1, the system's double-click time, multiplied with
	/// 4, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DragExpandTime, get_AllowDragDrop, get_RegisterForOLEDragDrop, get_DragScrollTimeBase,
	///     Raise_DragMouseMove, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_DragExpandTime(LONG* pValue);
	/// \brief <em>Sets the \c DragExpandTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be placed over an item during a
	/// drag'n'drop operation before this item will be expanded automatically. If set to 0,
	/// auto-expansion is disabled. If set to -1, the system's double-click time multiplied with
	/// 4 is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DragExpandTime, put_AllowDragDrop, put_RegisterForOLEDragDrop, put_DragScrollTimeBase,
	///     Raise_DragMouseMove, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_DragExpandTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c DraggedItems property</em>
	///
	/// Retrieves a collection object wrapping the control's items that are currently dragged. These
	/// are the same items that were passed to the \c BeginDrag or \c OLEDrag method.
	///
	/// \param[out] ppItems Receives the collection object's \c ITreeViewItemContainer implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa BeginDrag, OLEDrag, TreeViewItemContainer, get_TreeItems
	virtual HRESULT STDMETHODCALLTYPE get_DraggedItems(ITreeViewItemContainer** ppItems);
	/// \brief <em>Retrieves the current setting of the \c DragScrollTimeBase property</em>
	///
	/// Retrieves the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time, divided by 4, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DragScrollTimeBase, get_AllowDragDrop, get_RegisterForOLEDragDrop, get_DragExpandTime,
	///     Raise_DragMouseMove, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_DragScrollTimeBase(LONG* pValue);
	/// \brief <em>Sets the \c DragScrollTimeBase property</em>
	///
	/// Sets the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time divided by 4 is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DragScrollTimeBase, put_AllowDragDrop, put_RegisterForOLEDragDrop, put_DragExpandTime,
	///     Raise_DragMouseMove, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_DragScrollTimeBase(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c DrawImagesAsynchronously property</em>
	///
	/// Retrieves whether the control draws the items' associated images asynchronously. Asynchronous
	/// image drawing may improve performance. If set to \c VARIANT_TRUE, item images are drawn
	/// asynchronously; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_DrawImagesAsynchronously, get_hImageList, Raise_ItemAsynchronousDrawFailed
	virtual HRESULT STDMETHODCALLTYPE get_DrawImagesAsynchronously(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DrawImagesAsynchronously property</em>
	///
	/// Sets whether the control draws the items' associated images asynchronously. Asynchronous
	/// image drawing may improve performance. If set to \c VARIANT_TRUE, item images are drawn
	/// asynchronously; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_DrawImagesAsynchronously, put_hImageList, Raise_ItemAsynchronousDrawFailed
	virtual HRESULT STDMETHODCALLTYPE put_DrawImagesAsynchronously(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DropHilitedItem property</em>
	///
	/// Retrieves the item, that is the target of a drag'n'drop operation. Its background is drawn
	/// highlighted. If set to \c NULL, no item is a drop target.
	///
	/// \param[out] ppDropHilitedItem Receives the drop target's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_DropHilitedItem, TreeViewItem::get_DropHilited, get_AllowDragDrop,
	///     get_RegisterForOLEDragDrop, get_CaretItem
	virtual HRESULT STDMETHODCALLTYPE get_DropHilitedItem(ITreeViewItem** ppDropHilitedItem);
	/// \brief <em>Sets the \c DropHilitedItem property</em>
	///
	/// Sets the item, that is the target of a drag'n'drop operation. It's background is drawn
	/// highlighted. If set to \c NULL, no item is a drop target.
	///
	/// \param[in] pNewDropHilitedItem The new drop target's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DropHilitedItem, put_AllowDragDrop, put_RegisterForOLEDragDrop, putref_CaretItem,
	///     Raise_DragMouseMove, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE putref_DropHilitedItem(ITreeViewItem* pNewDropHilitedItem);
	/// \brief <em>Retrieves the current setting of the \c EditBackColor property</em>
	///
	/// Retrieves the label-edit control's background color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_EditBackColor, get_EditForeColor, get_BackColor
	virtual HRESULT STDMETHODCALLTYPE get_EditBackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c EditBackColor property</em>
	///
	/// Sets the label-edit control's background color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_EditBackColor, put_EditForeColor, put_BackColor
	virtual HRESULT STDMETHODCALLTYPE put_EditBackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c EditForeColor property</em>
	///
	/// Retrieves the label-edit control's text color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_EditForeColor, get_EditBackColor, get_ForeColor
	virtual HRESULT STDMETHODCALLTYPE get_EditForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c EditForeColor property</em>
	///
	/// Sets the label-edit control's text color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_EditForeColor, put_EditBackColor, put_ForeColor
	virtual HRESULT STDMETHODCALLTYPE put_EditForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c EditHoverTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be located over the label-edit
	/// control's client area before the \c EditMouseHover event is fired. If set to -1, the system
	/// hover time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_EditHoverTime, Raise_EditMouseHover, get_HoverTime
	virtual HRESULT STDMETHODCALLTYPE get_EditHoverTime(LONG* pValue);
	/// \brief <em>Sets the \c EditHoverTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be located over the label-edit control's
	/// client area before the \c EditMouseHover event is fired. If set to -1, the system hover time
	/// is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_EditHoverTime, Raise_EditMouseHover, put_HoverTime
	virtual HRESULT STDMETHODCALLTYPE put_EditHoverTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c EditIMEMode property</em>
	///
	/// Retrieves the label-edit control's IME mode. IME is a Windows feature making it easy to enter
	/// Asian characters. Any of the values defined by the \c IMEModeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_EditIMEMode, get_IMEMode, ExTVwLibU::IMEModeConstants
	/// \else
	///   \sa put_EditIMEMode, get_IMEMode, ExTVwLibA::IMEModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_EditIMEMode(IMEModeConstants* pValue);
	/// \brief <em>Sets the \c EditIMEMode property</em>
	///
	/// Sets the label-edit control's IME mode. IME is a Windows feature making it easy to enter
	/// Asian characters. Any of the values defined by the \c IMEModeConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_EditIMEMode, put_IMEMode, ExTVwLibU::IMEModeConstants
	/// \else
	///   \sa get_EditIMEMode, put_IMEMode, ExTVwLibA::IMEModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_EditIMEMode(IMEModeConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c EditText property</em>
	///
	/// Retrieves the text displayed by the control's label-edit control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_EditText, TreeViewItem::get_Text, Raise_EditChange
	virtual HRESULT STDMETHODCALLTYPE get_EditText(BSTR* pValue);
	/// \brief <em>Sets the \c EditText property</em>
	///
	/// Sets the text displayed by the control's label-edit control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_EditText, TreeViewItem::put_Text, Raise_EditChange
	virtual HRESULT STDMETHODCALLTYPE put_EditText(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c Enabled property</em>
	///
	/// Retrieves whether the control is enabled or disabled for user input. If set to \c VARIANT_TRUE,
	/// it reacts to user input; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Enabled, TreeViewItem::get_Grayed
	virtual HRESULT STDMETHODCALLTYPE get_Enabled(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Enabled property</em>
	///
	/// Enables or disables the control for user input. If set to \c VARIANT_TRUE, the control reacts to
	/// user input; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Enabled, TreeViewItem::put_Grayed
	virtual HRESULT STDMETHODCALLTYPE put_Enabled(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c FadeExpandos property</em>
	///
	/// Retrieves whether the expando images are faded in, if the user moves the mouse cursor into the
	/// control's client area, and out, if the user moves the mouse cursor out of the control's client
	/// area. If set to \c VARIANT_TRUE, fading takes place; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          If this property is set to \c VARIANT_TRUE, the \c HotTracking property should be set to a
	///          value other than \c htrNone. Otherwise expando fading won't work correctly.
	///
	/// \sa put_FadeExpandos, get_HotTracking, TreeViewItem::get_HasExpando, Raise_MouseEnter,
	///     Raise_MouseLeave
	virtual HRESULT STDMETHODCALLTYPE get_FadeExpandos(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c FadeExpandos property</em>
	///
	/// Sets whether the expando images are faded in, if the user moves the mouse cursor into the
	/// control's client area, and out, if the user moves the mouse cursor out of the control's client
	/// area. If set to \c VARIANT_TRUE, fading takes place; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          If this property is set to \c VARIANT_TRUE, the \c HotTracking property should be set to a
	///          value other than \c htrNone. Otherwise expando fading won't work correctly.
	///
	/// \sa get_FadeExpandos, put_HotTracking, TreeViewItem::put_HasExpando, Raise_MouseEnter,
	///     Raise_MouseLeave
	virtual HRESULT STDMETHODCALLTYPE put_FadeExpandos(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c FavoritesStyle property</em>
	///
	/// Retrieves whether a rectangle is drawn around the caret item and its sub-items to group them
	/// (like in Windows Favorites management window). If set to \c VARIANT_TRUE, the items are grouped by a
	/// rectangle; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This feature should be used with \c SingleExpand set to \c seNormal.
	///
	/// \sa put_FavoritesStyle, get_GroupBoxColor, get_SingleExpand, get_FullRowSelect, get_HotTracking
	virtual HRESULT STDMETHODCALLTYPE get_FavoritesStyle(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c FavoritesStyle property</em>
	///
	/// Sets whether a rectangle is drawn around the caret item and its subitems to group them
	/// (like in Windows Favorites management window). If set to \c VARIANT_TRUE, the items are grouped by a
	/// rectangle; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This feature should be used with \c SingleExpand set to \c seNormal.
	///
	/// \sa get_FavoritesStyle, put_GroupBoxColor, put_SingleExpand, put_FullRowSelect, put_HotTracking
	virtual HRESULT STDMETHODCALLTYPE put_FavoritesStyle(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c FirstRootItem property</em>
	///
	/// Retrieves the control's very first item.
	///
	/// \param[out] ppFirstItem Receives the first item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_LastRootItem, get_FirstVisibleItem, get_TreeItems
	virtual HRESULT STDMETHODCALLTYPE get_FirstRootItem(ITreeViewItem** ppFirstItem);
	/// \brief <em>Retrieves the current setting of the \c FirstVisibleItem property</em>
	///
	/// Retrieves the control's first item, that is entirely located in the control's client area and
	/// therefore visible to the user.
	///
	/// \param[out] ppFirstItem Receives the first visible item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_FirstVisibleItem, get_LastVisibleItem, get_FirstRootItem, get_TreeItems
	virtual HRESULT STDMETHODCALLTYPE get_FirstVisibleItem(ITreeViewItem** ppFirstItem);
	/// \brief <em>Sets the \c FirstVisibleItem property</em>
	///
	/// Sets the control's first item, that is entirely located in the control's client area and
	/// therefore visible to the user.
	///
	/// \param[in] pNewFirstItem The new first visible item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FirstVisibleItem
	virtual HRESULT STDMETHODCALLTYPE putref_FirstVisibleItem(ITreeViewItem* pNewFirstItem);
	/// \brief <em>Retrieves the current setting of the \c Font property</em>
	///
	/// Retrieves the control's font. It's used to draw the items' text.
	///
	/// \param[out] ppFont Receives the font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Font, putref_Font, get_UseSystemFont, get_ForeColor, TreeViewItem::get_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE get_Font(IFontDisp** ppFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the items' text.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewFont is cloned.
	///
	/// \sa get_Font, putref_Font, put_UseSystemFont, put_ForeColor, TreeViewItem::put_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE put_Font(IFontDisp* pNewFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the items' text.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Font, put_Font, put_UseSystemFont, put_ForeColor, TreeViewItem::put_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE putref_Font(IFontDisp* pNewFont);
	/// \brief <em>Retrieves the current setting of the \c ForeColor property</em>
	///
	/// Retrieves the control's text color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ForeColor, get_BackColor, get_EditForeColor, get_InsertMarkColor, get_LineColor
	virtual HRESULT STDMETHODCALLTYPE get_ForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ForeColor property</em>
	///
	/// Sets the control's text color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ForeColor, put_BackColor, put_EditForeColor, put_InsertMarkColor, put_LineColor
	virtual HRESULT STDMETHODCALLTYPE put_ForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c FullRowSelect property</em>
	///
	/// Retrieves whether an item's entire row is highlighted if it is selected. If set to
	/// \c VARIANT_TRUE, the entire row, otherwise only the item's text is highlighted.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_FullRowSelect, get_SingleExpand, get_FavoritesStyle
	virtual HRESULT STDMETHODCALLTYPE get_FullRowSelect(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c FullRowSelect property</em>
	///
	/// Sets whether an item's entire row is highlighted if it is selected. If set to
	/// \c VARIANT_TRUE, the entire row, otherwise only the item's text is highlighted.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c VARIANT_TRUE while the \c TreeViewStyle property is set to
	///          include \c tvsLines, the \c tvsLines flag will be removed from the \c TreeViewStyle
	///          property.
	///
	/// \sa get_FullRowSelect, put_SingleExpand, put_FavoritesStyle, put_TreeViewStyle
	virtual HRESULT STDMETHODCALLTYPE put_FullRowSelect(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c GroupBoxColor property</em>
	///
	/// Retrieves the color of the rectangle that is drawn around the caret item and its sub-items
	/// if the \c FavoritesStyle property is set to \c VARIANT_TRUE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_GroupBoxColor, get_FavoritesStyle
	virtual HRESULT STDMETHODCALLTYPE get_GroupBoxColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c GroupBoxColor property</em>
	///
	/// Sets the color of the rectangle that is drawn around the caret item and its sub-items if the
	/// \c FavoritesStyle property is set to \c VARIANT_TRUE.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_GroupBoxColor, put_FavoritesStyle
	virtual HRESULT STDMETHODCALLTYPE put_GroupBoxColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c hDragImageList property</em>
	///
	/// Retrieves the handle to the imagelist containing the drag image that is used during a
	/// drag'n'drop operation to visualize the dragged items.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowDragImage, BeginDrag, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_hDragImageList(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hImageList property</em>
	///
	/// Retrieves the handle of the specified imagelist.
	///
	/// \param[in] imageList The imageList to retrieve. Any of the values defined by the
	///            \c ImageListConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \if UNICODE
	///   \sa put_hImageList, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa put_hImageList, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_hImageList(ImageListConstants imageList, OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hImageList property</em>
	///
	/// Sets the handle of the specified imagelist.
	///
	/// \param[in] imageList The imageList to set. Any of the values defined by the \c ImageListConstants
	///            enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \if UNICODE
	///   \sa get_hImageList, ExTVwLibU::ImageListConstants
	/// \else
	///   \sa get_hImageList, ExTVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_hImageList(ImageListConstants imageList, OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c HotTracking property</em>
	///
	/// Retrieves the hot tracking mode. If enabled, the item underneath the mouse cursor is underlined.
	/// Any of the values defined by the \c HotTrackingConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c htrNone, expando fading (see \c FadeExpandos property)
	///          won't work correctly.
	///
	/// \if UNICODE
	///   \sa put_HotTracking, get_FadeExpandos, get_SingleExpand, ExTVwLibU::HotTrackingConstants
	/// \else
	///   \sa put_HotTracking, get_FadeExpandos, get_SingleExpand, ExTVwLibA::HotTrackingConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_HotTracking(HotTrackingConstants* pValue);
	/// \brief <em>Sets the \c HotTracking property</em>
	///
	/// Sets the hot tracking mode. If enabled, the item underneath the mouse cursor is underlined.
	/// Any of the values defined by the \c HotTrackingConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c htrNone, expando fading (see \c FadeExpandos property)
	///          won't work correctly.
	///
	/// \if UNICODE
	///   \sa get_HotTracking, put_FadeExpandos, put_SingleExpand, ExTVwLibU::HotTrackingConstants
	/// \else
	///   \sa get_HotTracking, put_FadeExpandos, put_SingleExpand, ExTVwLibA::HotTrackingConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_HotTracking(HotTrackingConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c HoverTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HoverTime, Raise_MouseHover, get_EditHoverTime
	virtual HRESULT STDMETHODCALLTYPE get_HoverTime(LONG* pValue);
	/// \brief <em>Sets the \c HoverTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HoverTime, Raise_MouseHover, put_EditHoverTime
	virtual HRESULT STDMETHODCALLTYPE put_HoverTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c hWnd property</em>
	///
	/// Retrieves the control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hWndEdit, Raise_RecreatedControlWindow, Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndEdit property</em>
	///
	/// Retrieves the label-editing control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hWnd, Raise_CreatedEditControlWindow, Raise_DestroyedEditControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWndEdit(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndToolTip property</em>
	///
	/// Retrieves the tooltip control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_hWndToolTip, get_ShowToolTips, get_RichToolTips
	virtual HRESULT STDMETHODCALLTYPE get_hWndToolTip(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hWndToolTip property</em>
	///
	/// Sets the tooltip control to use.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_hWndToolTip, put_ShowToolTips, put_RichToolTips
	virtual HRESULT STDMETHODCALLTYPE put_hWndToolTip(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c IMEMode property</em>
	///
	/// Retrieves the control's IME mode. IME is a Windows feature making it easy to enter Asian
	/// characters. Any of the values defined by the \c IMEModeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_IMEMode, get_EditIMEMode, ExTVwLibU::IMEModeConstants
	/// \else
	///   \sa put_IMEMode, get_EditIMEMode, ExTVwLibA::IMEModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IMEMode(IMEModeConstants* pValue);
	/// \brief <em>Sets the \c IMEMode property</em>
	///
	/// Sets the control's IME mode. IME is a Windows feature making it easy to enter Asian
	/// characters. Any of the values defined by the \c IMEModeConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IMEMode, put_EditIMEMode, ExTVwLibU::IMEModeConstants
	/// \else
	///   \sa get_IMEMode, put_EditIMEMode, ExTVwLibA::IMEModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IMEMode(IMEModeConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c IncrementalSearchString property</em>
	///
	/// Retrieves the control's current incremental search-string. This string is used to select
	/// an item based on characters entered by the user.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa Raise_IncrementalSearchStringChanging
	virtual HRESULT STDMETHODCALLTYPE get_IncrementalSearchString(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c Indent property</em>
	///
	/// Retrieves the items' indentation in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Indent, TreeViewItem::get_Level
	virtual HRESULT STDMETHODCALLTYPE get_Indent(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c Indent property</em>
	///
	/// Sets the items' indentation in pixels.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Indent, TreeViewItem::get_Level
	virtual HRESULT STDMETHODCALLTYPE put_Indent(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c IndentStateImages property</em>
	///
	/// Retrieves whether the state images are left-aligned or placed between the items' expandos and the
	/// items' captions. If set to \c VARIANT_TRUE, state images follow the respective item's indentation and
	/// are placed between its expando and its caption. If set to \c VARIANT_FALSE, state images are aligned
	/// to the control's edge.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks On current versions of Windows state-images can't be toggled using the mouse if this
	///          property is set to \c False.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_IndentStateImages, get_hImageList, get_ShowStateImages, TreeViewItem::get_StateImageIndex
	virtual HRESULT STDMETHODCALLTYPE get_IndentStateImages(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c IndentStateImages property</em>
	///
	/// Sets whether the state images are left-aligned or placed between the items' expandos and the items'
	/// captions. If set to \c VARIANT_TRUE, state images follow the respective item's indentation and are
	/// placed between its expando and its caption. If set to \c VARIANT_FALSE, state images are aligned to
	/// the control's edge.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks On current versions of Windows state-images can't be toggled using the mouse if this
	///          property is set to \c False.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_IndentStateImages, put_hImageList, put_ShowStateImages, TreeViewItem::put_StateImageIndex
	virtual HRESULT STDMETHODCALLTYPE put_IndentStateImages(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c InsertMarkColor property</em>
	///
	/// Retrieves the color that is the control's insertion mark is drawn in.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_InsertMarkColor, GetInsertMarkPosition, get_BackColor, get_ForeColor, get_LineColor
	virtual HRESULT STDMETHODCALLTYPE get_InsertMarkColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c InsertMarkColor property</em>
	///
	/// Sets the color that is the control's insertion mark is drawn in.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_InsertMarkColor, SetInsertMarkPosition, put_BackColor, put_ForeColor, put_LineColor
	virtual HRESULT STDMETHODCALLTYPE put_InsertMarkColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the control's release type</em>
	///
	/// Retrieves the control's release type. This property is part of the fingerprint that uniquely
	/// identifies each software written by Timo "TimoSoft" Kunze. If set to \c VARIANT_TRUE, the
	/// control was compiled for release; otherwise it was compiled for debugging.
	///
	/// \param[out] pValue The release type.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_IsRelease(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c ItemBoundingBoxDefinition property</em>
	///
	/// Retrieves the parts of an item that are treated as such if firing any kind of mouse event. Any
	/// combination of the values defined by the \c ItemBoundingBoxDefinitionConstants enumeration is
	/// valid. E. g. if set to \c ibbdItemIcon combined with \c ibbdItemIndent, the \c treeItem
	/// parameter of the \c MouseMove event will identify the item only if the mouse cursor is located
	/// over the item's icon or indentation space; otherwise (e. g. if the cursor is located over the
	/// item's text) it will be \c NULL.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ItemBoundingBoxDefinition, HitTest, ExTVwLibU::ItemBoundingBoxDefinitionConstants
	/// \else
	///   \sa put_ItemBoundingBoxDefinition, HitTest, ExTVwLibA::ItemBoundingBoxDefinitionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ItemBoundingBoxDefinition(ItemBoundingBoxDefinitionConstants* pValue);
	/// \brief <em>Sets the \c ItemBoundingBoxDefinition property</em>
	///
	/// Sets the parts of an item that are treated as such if firing any kind of mouse event. Any
	/// combination of the values defined by the \c ItemBoundingBoxDefinitionConstants enumeration is
	/// valid. E. g. if set to \c ibbdItemIcon combined with \c ibbdItemIndent, the \c treeItem
	/// parameter of the \c MouseMove event will identify the item only if the mouse cursor is located
	/// over the item's icon or indentation space; otherwise (e. g. if the cursor is located over the
	/// item's text) it will be \c NULL.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ItemBoundingBoxDefinition, HitTest, ExTVwLibU::ItemBoundingBoxDefinitionConstants
	/// \else
	///   \sa get_ItemBoundingBoxDefinition, HitTest, ExTVwLibA::ItemBoundingBoxDefinitionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ItemBoundingBoxDefinition(ItemBoundingBoxDefinitionConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ItemHeight property</em>
	///
	/// Retrieves the items' basic height in pixels. The height of an item may be any multiple of this
	/// height, depending on the item's \c HeightIncrement setting.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ItemHeight, TreeViewItem::get_HeightIncrement
	virtual HRESULT STDMETHODCALLTYPE get_ItemHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c ItemHeight property</em>
	///
	/// Sets the items' basic height in pixels. The height of an item may be any multiple of this
	/// height, depending on the item's \c HeightIncrement setting.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ItemHeight, TreeViewItem::put_HeightIncrement
	virtual HRESULT STDMETHODCALLTYPE put_ItemHeight(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c ItemXBorder property</em>
	///
	/// Retrieves the thickness (in pixels) of the left and right borders of items.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is an undocumented feature of SysTreeView32.\n
	///          You should set this property to 0 if you want to set the \c BlendSelectedItemsIcons
	///          property to \c VARIANT_TRUE.
	///
	/// \sa put_ItemXBorder, get_ItemYBorder
	virtual HRESULT STDMETHODCALLTYPE get_ItemXBorder(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c ItemXBorder property</em>
	///
	/// Sets the thickness (in pixels) of the left and right borders of items.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is an undocumented feature of SysTreeView32.\n
	///          You should set this property to 0 if you want to set the \c BlendSelectedItemsIcons
	///          property to \c VARIANT_TRUE.
	///
	/// \sa get_ItemXBorder, put_ItemYBorder
	virtual HRESULT STDMETHODCALLTYPE put_ItemXBorder(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c ItemYBorder property</em>
	///
	/// Retrieves the thickness (in pixels) of the upper and lower borders of items.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is an undocumented feature of SysTreeView32.\n
	///          You should set this property to 0 if you want to set the \c BlendSelectedItemsIcons
	///          property to \c VARIANT_TRUE.
	///
	/// \sa put_ItemYBorder, get_ItemXBorder
	virtual HRESULT STDMETHODCALLTYPE get_ItemYBorder(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c ItemYBorder property</em>
	///
	/// Sets the thickness (in pixels) of the upper and lower borders of items.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is an undocumented feature of SysTreeView32.\n
	///          You should set this property to 0 if you want to set the \c BlendSelectedItemsIcons
	///          property to \c VARIANT_TRUE.
	///
	/// \sa get_ItemYBorder, put_ItemXBorder
	virtual HRESULT STDMETHODCALLTYPE put_ItemYBorder(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c LastRootItem property</em>
	///
	/// Retrieves the control's last top-level item.
	///
	/// \param[out] ppLastItem Receives the last item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FirstRootItem, get_LastVisibleItem, get_TreeItems
	virtual HRESULT STDMETHODCALLTYPE get_LastRootItem(ITreeViewItem** ppLastItem);
	/// \brief <em>Retrieves the current setting of the \c LastVisibleItem property</em>
	///
	/// Retrieves the control's last item, that is entirely located in the control's client area and
	/// therefore visible to the user.
	///
	/// \param[out] ppLastItem Receives the last visible item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_FirstVisibleItem, get_LastRootItem, get_TreeItems
	virtual HRESULT STDMETHODCALLTYPE get_LastVisibleItem(ITreeViewItem** ppLastItem);
	/// \brief <em>Retrieves the current setting of the \c LineColor property</em>
	///
	/// Retrieves the control's line color, which is used to draw the tree lines visualizing the hierarchy.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_LineColor, get_LineStyle, get_TreeViewStyle, get_BackColor, get_ForeColor,
	///     get_InsertMarkColor
	virtual HRESULT STDMETHODCALLTYPE get_LineColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c LineColor property</em>
	///
	/// Sets the control's line color, which is used to draw the tree lines visualizing the hierarchy.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_LineColor, put_LineStyle, put_TreeViewStyle, put_BackColor, put_ForeColor,
	///     put_InsertMarkColor
	virtual HRESULT STDMETHODCALLTYPE put_LineColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c LineStyle property</em>
	///
	/// Retrieves the control's line-style. Any of the values defined by the \c LineStyleConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_LineStyle, get_LineColor, get_TreeViewStyle, ExTVwLibU::LineStyleConstants
	/// \else
	///   \sa put_LineStyle, get_LineColor, get_TreeViewStyle, ExTVwLibA::LineStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_LineStyle(LineStyleConstants* pValue);
	/// \brief <em>Sets the \c LineStyle property</em>
	///
	/// Sets the control's line-style. Any of the values defined by the \c LineStyleConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_LineStyle, put_LineColor, put_TreeViewStyle, ExTVwLibU::LineStyleConstants
	/// \else
	///   \sa get_LineStyle, put_LineColor, put_TreeViewStyle, ExTVwLibA::LineStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_LineStyle(LineStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Locale property</em>
	///
	/// Retrieves the unique ID of the locale to use when sorting the tree view items.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The locale is used when sorting using the \c sobNumericIntText, \c sobNumericFloatText,
	///          \c sobCurrencyText or \c sobDateTimeText sorting criterion.
	///
	/// \if UNICODE
	///   \sa put_Locale, get_TextParsingFlags, SortSubItems, ExTVwLibU::SortByConstants
	/// \else
	///   \sa put_Locale, get_TextParsingFlags, SortSubItems, ExTVwLibA::SortByConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Locale(LONG* pValue);
	/// \brief <em>Sets the \c Locale property</em>
	///
	/// Sets the unique ID of the locale to use when sorting the tree view items.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The locale is used when sorting using the \c sobNumericIntText, \c sobNumericFloatText,
	///          \c sobCurrencyText or \c sobDateTimeText sorting criterion.
	///
	/// \if UNICODE
	///   \sa get_Locale, put_TextParsingFlags, SortSubItems, ExTVwLibU::SortByConstants
	/// \else
	///   \sa get_Locale, put_TextParsingFlags, SortSubItems, ExTVwLibA::SortByConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Locale(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c MaxScrollTime property</em>
	///
	/// Retrieves the number of milliseconds the control may need to scroll its content.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MaxScrollTime
	virtual HRESULT STDMETHODCALLTYPE get_MaxScrollTime(LONG* pValue);
	/// \brief <em>Sets the \c MaxScrollTime property</em>
	///
	/// Sets the number of milliseconds the control may need to scroll its content.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MaxScrollTime
	virtual HRESULT STDMETHODCALLTYPE put_MaxScrollTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c MouseIcon property</em>
	///
	/// Retrieves a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[out] ppMouseIcon Receives the picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, ExTVwLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, ExTVwLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MouseIcon(IPictureDisp** ppMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewMouseIcon is cloned.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, ExTVwLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, ExTVwLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, ExTVwLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, ExTVwLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE putref_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Retrieves the current setting of the \c MousePointer property</em>
	///
	/// Retrieves the cursor's type that's used if the mouse cursor is placed within the control's client
	/// area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MousePointer, get_MouseIcon, ExTVwLibU::MousePointerConstants
	/// \else
	///   \sa put_MousePointer, get_MouseIcon, ExTVwLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MousePointer(MousePointerConstants* pValue);
	/// \brief <em>Sets the \c MousePointer property</em>
	///
	/// Sets the cursor's type that's used if the mouse cursor is placed within the control's client
	/// area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MousePointer, put_MouseIcon, ExTVwLibU::MousePointerConstants
	/// \else
	///   \sa get_MousePointer, put_MouseIcon, ExTVwLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MousePointer(MousePointerConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c OLEDragImageStyle property</em>
	///
	/// Retrieves the appearance of the OLE drag images generated by the control. Any of the values defined
	/// by the \c OLEDragImageStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_OLEDragImageStyle, get_SupportOLEDragImages, OLEDrag, ExTVwLibU::OLEDragImageStyleConstants
	/// \else
	///   \sa put_OLEDragImageStyle, get_SupportOLEDragImages, OLEDrag, ExTVwLibA::OLEDragImageStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue);
	/// \brief <em>Sets the \c OLEDragImageStyle property</em>
	///
	/// Sets the appearance of the OLE drag images generated by the control. Any of the values defined
	/// by the \c OLEDragImageStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_OLEDragImageStyle, put_SupportOLEDragImages, OLEDrag, ExTVwLibU::OLEDragImageStyleConstants
	/// \else
	///   \sa get_OLEDragImageStyle, put_SupportOLEDragImages, OLEDrag, ExTVwLibA::OLEDragImageStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_OLEDragImageStyle(OLEDragImageStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ProcessContextMenuKeys property</em>
	///
	/// Retrieves whether the control fires the \c ContextMenu and \c EditContextMenu events if the
	/// user presses [SHIFT]+[F10] or [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the events
	/// are fired; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ProcessContextMenuKeys, Raise_ContextMenu, Raise_EditContextMenu
	virtual HRESULT STDMETHODCALLTYPE get_ProcessContextMenuKeys(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ProcessContextMenuKeys property</em>
	///
	/// Sets whether the control fires the \c ContextMenu and \c EditContextMenu events if the
	/// user presses [SHIFT]+[F10] or [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the events
	/// are fired; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ProcessContextMenuKeys, Raise_ContextMenu, Raise_EditContextMenu
	virtual HRESULT STDMETHODCALLTYPE put_ProcessContextMenuKeys(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's programmer(s)</em>
	///
	/// Retrieves the name(s) of the control's programmer(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The programmer.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Programmer(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c RegisterForOLEDragDrop property</em>
	///
	/// Retrieves whether the control is registered as a target for OLE drag'n'drop. If set to
	/// \c VARIANT_TRUE, the control accepts OLE drag'n'drop actions; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RegisterForOLEDragDrop, get_AllowDragDrop, get_SupportOLEDragImages, Raise_OLEDragEnter
	virtual HRESULT STDMETHODCALLTYPE get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c RegisterForOLEDragDrop property</em>
	///
	/// Sets whether the control is registered as a target for OLE drag'n'drop. If set to
	/// \c VARIANT_TRUE, the control accepts OLE drag'n'drop actions; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RegisterForOLEDragDrop, put_AllowDragDrop, put_SupportOLEDragImages, Raise_OLEDragEnter
	virtual HRESULT STDMETHODCALLTYPE put_RegisterForOLEDragDrop(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c RichToolTips property</em>
	///
	/// Retrieves whether the control's tooltips are "rich formatted", i.e. they include the item's icon.
	/// If set to \c VARIANT_TRUE, the tooltips include the item's icon; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_RichToolTips, get_ShowToolTips, get_hWndToolTip, Raise_ItemGetInfoTipText
	virtual HRESULT STDMETHODCALLTYPE get_RichToolTips(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c RichToolTips property</em>
	///
	/// Sets whether the control's tooltips are "rich formatted", i.e. they include the item's icon.
	/// If set to \c VARIANT_TRUE, the tooltips include the item's icon; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_RichToolTips, put_ShowToolTips, put_hWndToolTip, Raise_ItemGetInfoTipText
	virtual HRESULT STDMETHODCALLTYPE put_RichToolTips(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c RightToLeft property</em>
	///
	/// Retrieves whether bidirectional features are enabled or disabled. Any combination of the values
	/// defined by the \c RightToLeftConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_RightToLeft, get_IMEMode, ExTVwLibU::RightToLeftConstants
	/// \else
	///   \sa put_RightToLeft, get_IMEMode, ExTVwLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RightToLeft(RightToLeftConstants* pValue);
	/// \brief <em>Sets the \c RightToLeft property</em>
	///
	/// Enables or disables bidirectional features. Any combination of the values defined by the
	/// \c RightToLeftConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Setting or clearing the \c rtlLayout flag destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_RightToLeft, put_IMEMode, ExTVwLibU::RightToLeftConstants
	/// \else
	///   \sa get_RightToLeft, put_IMEMode, ExTVwLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeft(RightToLeftConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c RootItems property</em>
	///
	/// Retrieves a collection object wrapping the control's top-level items.
	///
	/// \param[out] ppItems Receives the collection object's \c ITreeViewItems implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_TreeItems, TreeViewItems
	virtual HRESULT STDMETHODCALLTYPE get_RootItems(ITreeViewItems** ppItems = NULL);
	/// \brief <em>Retrieves the current setting of the \c ScrollBars property</em>
	///
	/// Retrieves the scrollbars to show if necessary. Any of the values defined by the
	/// \c ScrollBarsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property may destroy and recreate the control window.
	///
	/// \if UNICODE
	///   \sa put_ScrollBars, get_AutoHScroll, ExTVwLibU::ScrollBarsConstants
	/// \else
	///   \sa put_ScrollBars, get_AutoHScroll, ExTVwLibA::ScrollBarsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ScrollBars(ScrollBarsConstants* pValue);
	/// \brief <em>Sets the \c ScrollBars property</em>
	///
	/// Sets the scrollbars to show if necessary. Any of the values defined by the \c ScrollBarsConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property may destroy and recreate the control window.
	///
	/// \if UNICODE
	///   \sa get_ScrollBars, put_AutoHScroll, ExTVwLibU::ScrollBarsConstants
	/// \else
	///   \sa get_ScrollBars, put_AutoHScroll, ExTVwLibA::ScrollBarsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ScrollBars(ScrollBarsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowDragImage property</em>
	///
	/// Retrieves whether the drag image is visible or hidden. If set to \c VARIANT_TRUE, it is visible;
	/// otherwise hidden.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowDragImage, get_hDragImageList, get_SupportOLEDragImages, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_ShowDragImage(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowDragImage property</em>
	///
	/// Sets whether the drag image is visible or hidden. If set to \c VARIANT_TRUE, it is visible; otherwise
	/// hidden.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowDragImage, get_hDragImageList, put_SupportOLEDragImages, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_ShowDragImage(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowStateImages property</em>
	///
	/// Retrieves whether the control draws a state image (usually a checkbox) to the left of each
	/// item. If set to \c VARIANT_TRUE, the state images are drawn; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property may destroy and recreate the control window.
	///
	/// \sa put_ShowStateImages, get_BuiltInStateImages, get_hImageList, get_IndentStateImages,
	///     get_BlendSelectedItemsIcons, TreeViewItem::get_StateImageIndex
	virtual HRESULT STDMETHODCALLTYPE get_ShowStateImages(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowStateImages property</em>
	///
	/// Sets whether the control draws a state image (usually a checkbox) to the left of each item.
	/// If set to \c VARIANT_TRUE, the state images are drawn; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property may destroy and recreate the control window.
	///
	/// \sa get_ShowStateImages, put_BuiltInStateImages, put_hImageList, put_IndentStateImages,
	///     put_BlendSelectedItemsIcons, TreeViewItem::put_StateImageIndex
	virtual HRESULT STDMETHODCALLTYPE put_ShowStateImages(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowToolTips property</em>
	///
	/// Retrieves whether the control displays any tooltips. A treeview knows two kinds of tooltips:
	/// Tooltips that display the complete text of truncated items and so-called info tips displaying
	/// details about the item that the mouse cursor is placed over.\n
	/// If this property is set to \c VARIANT_TRUE, the truncated items' tooltips are displayed and the
	/// \c ItemGetInfoTipText event is fired for the item below the mouse cursor. Otherwise no tooltips
	/// are shown.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowToolTips, get_hWndToolTip, get_RichToolTips, TreeViewItem::DisplayInfoTip,
	///     Raise_ItemGetInfoTipText
	virtual HRESULT STDMETHODCALLTYPE get_ShowToolTips(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowToolTips property</em>
	///
	/// Sets whether the control displays any tooltips. A treeview knows two kinds of tooltips:
	/// Tooltips that display the complete text of truncated items and so-called info tips displaying
	/// details about the item that the mouse cursor is placed over.\n
	/// If this property is set to \c VARIANT_TRUE, the truncated items' tooltips are displayed and the
	/// \c ItemGetInfoTipText event is fired for the item below the mouse cursor. Otherwise no tooltips
	/// are shown.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowToolTips, put_hWndToolTip, put_RichToolTips, TreeViewItem::DisplayInfoTip,
	///     Raise_ItemGetInfoTipText
	virtual HRESULT STDMETHODCALLTYPE put_ShowToolTips(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SingleExpand property</em>
	///
	/// Retrieves the way items are expanded and collapsed. Any of the values defined by the
	/// \c SingleExpandConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_SingleExpand, Raise_SingleExpanding, get_FullRowSelect, get_HotTracking,
	///       ExTVwLibU::SingleExpandConstants
	/// \else
	///   \sa put_SingleExpand, Raise_SingleExpanding, get_FullRowSelect, get_HotTracking,
	///       ExTVwLibA::SingleExpandConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SingleExpand(SingleExpandConstants* pValue);
	/// \brief <em>Sets the \c SingleExpand property</em>
	///
	/// Sets the way items are expanded and collapsed. Any of the values defined by the
	/// \c SingleExpandConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_SingleExpand, Raise_SingleExpanding, put_FullRowSelect, put_HotTracking,
	///       ExTVwLibU::SingleExpandConstants
	/// \else
	///   \sa get_SingleExpand, Raise_SingleExpanding, put_FullRowSelect, put_HotTracking,
	///       ExTVwLibA::SingleExpandConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_SingleExpand(SingleExpandConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c SortOrder property</em>
	///
	/// Retrieves the direction items are sorted in. Any of the values defined by the \c SortOrderConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_SortOrder, SortItems, Raise_ChangingSortOrder, Raise_ChangedSortOrder,
	///       ExTVwLibU::SortOrderConstants
	/// \else
	///   \sa put_SortOrder, SortItems, Raise_ChangingSortOrder, Raise_ChangedSortOrder,
	///       ExTVwLibA::SortOrderConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SortOrder(SortOrderConstants* pValue);
	/// \brief <em>Sets the \c SortOrder property</em>
	///
	/// Sets the direction items are sorted in. Any of the values defined by the \c SortOrderConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_SortOrder, SortItems, Raise_ChangingSortOrder, Raise_ChangedSortOrder,
	///       ExTVwLibU::SortOrderConstants
	/// \else
	///   \sa get_SortOrder, SortItems, Raise_ChangingSortOrder, Raise_ChangedSortOrder,
	///       ExTVwLibA::SortOrderConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_SortOrder(SortOrderConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c SupportOLEDragImages property</em>
	///
	/// Retrieves whether the control creates \c IDropTargetHelper and \c IDragSourceHelper objects, so
	/// that a drag image can be shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control
	/// creates the objects; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa put_SupportOLEDragImages, get_RegisterForOLEDragDrop, get_hImageList, get_ShowDragImage,
	///     get_OLEDragImageStyle, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646441.aspx">IDragSourceHelper</a>
	virtual HRESULT STDMETHODCALLTYPE get_SupportOLEDragImages(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SupportOLEDragImages property</em>
	///
	/// Sets whether the control creates \c IDropTargetHelper and \c IDragSourceHelper objects, so
	/// that a drag image can be shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control
	/// creates the objects; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa get_SupportOLEDragImages, put_RegisterForOLEDragDrop, put_hImageList, put_ShowDragImage,
	///     put_OLEDragImageStyle, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646441.aspx">IDragSourceHelper</a>
	virtual HRESULT STDMETHODCALLTYPE put_SupportOLEDragImages(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's tester(s)</em>
	///
	/// Retrieves the name(s) of the control's tester(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The name(s) of the tester(s).
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease,
	///     get_Programmer
	virtual HRESULT STDMETHODCALLTYPE get_Tester(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c TextParsingFlags property</em>
	///
	/// Retrieves the options to apply when parsing the item text into a numerical or date value. The parsing
	/// results may be used when sorting the tree view items.
	///
	/// \param[in] parsingFunction Specifies the parsing function for which to retrieve the options. Any of
	///            the values defined by the \c TextParsingFunctionConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The parsing options are used when sorting using the \c sobNumericIntText,
	///          \c sobNumericFloatText, \c sobCurrencyText or \c sobDateTimeText sorting criterion.\n
	///          They are also used when sorting using the \c sobText criterion if a specific locale
	///          identifier has been set.
	///
	/// \if UNICODE
	///   \sa put_TextParsingFlags, get_Locale, SortSubItems, ExTVwLibU::SortByConstants,
	///       ExTVwLibU::TextParsingFunctionConstants
	/// \else
	///   \sa put_TextParsingFlags, get_Locale, SortSubItems, ExTVwLibA::SortByConstants,
	///       ExTVwLibA::TextParsingFunctionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TextParsingFlags(TextParsingFunctionConstants parsingFunction, LONG* pValue);
	/// \brief <em>Sets the \c TextParsingFlags property</em>
	///
	/// Sets the options to apply when parsing the item text into a numerical or date value. The parsing
	/// results may be used when sorting the tree view items.
	///
	/// \param[in] parsingFunction Specifies the parsing function for which to set the options. Any of
	///            the values defined by the \c TextParsingFunctionConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The parsing options are used when sorting using the \c sobNumericIntText,
	///          \c sobNumericFloatText, \c sobCurrencyText or \c sobDateTimeText sorting criterion.\n
	///          They are also used when sorting using the \c sobText criterion if a specific locale
	///          identifier has been set.
	///
	/// \if UNICODE
	///   \sa get_TextParsingFlags, put_Locale, SortSubItems, ExTVwLibU::SortByConstants,
	///       ExTVwLibU::TextParsingFunctionConstants
	/// \else
	///   \sa get_TextParsingFlags, put_Locale, SortSubItems, ExTVwLibA::SortByConstants,
	///       ExTVwLibA::TextParsingFunctionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_TextParsingFlags(TextParsingFunctionConstants parsingFunction, LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c TreeItems property</em>
	///
	/// Retrieves a collection object wrapping the control's items.
	///
	/// \param[out] ppItems Receives the collection object's \c ITreeViewItems implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_RootItems, TreeViewItems
	virtual HRESULT STDMETHODCALLTYPE get_TreeItems(ITreeViewItems** ppItems);
	/// \brief <em>Retrieves the current setting of the \c TreeViewStyle property</em>
	///
	/// Retrieves whether some of the control's visual elements are enabled or disabled. Any
	/// combination of the values defined by the \c TreeViewStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_TreeViewStyle, get_LineStyle, TreeViewItem::get_HasExpando,
	///       ExTVwLibU::TreeViewStyleConstants
	/// \else
	///   \sa put_TreeViewStyle, get_LineStyle, TreeViewItem::get_HasExpando,
	///       ExTVwLibA::TreeViewStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TreeViewStyle(TreeViewStyleConstants* pValue);
	/// \brief <em>Sets the \c TreeViewStyle property</em>
	///
	/// Enables or disables some of the control's visual elements. Any combination of the values
	/// defined by the \c TreeViewStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to a value that includes \c tvsLines while the \c FullRowSelect
	///          property is set to \c VARIANT_TRUE, the \c FullRowSelect property will be changed to
	///          \c VARIANT_FALSE.
	///
	/// \if UNICODE
	///   \sa get_TreeViewStyle, put_LineStyle, put_FullRowSelect, TreeViewItem::put_HasExpando,
	///       ExTVwLibU::TreeViewStyleConstants
	/// \else
	///   \sa get_TreeViewStyle, put_LineStyle, put_FullRowSelect, TreeViewItem::put_HasExpando,
	///       ExTVwLibA::TreeViewStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_TreeViewStyle(TreeViewStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c UseSystemFont property</em>
	///
	/// Retrieves whether the control uses the system's default treeview font or the font specified by the
	/// \c Font property. If set to \c VARIANT_TRUE, the system font; otherwise the specified font is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseSystemFont, get_Font
	virtual HRESULT STDMETHODCALLTYPE get_UseSystemFont(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseSystemFont property</em>
	///
	/// Sets whether the control uses the system's default treeview font or the font specified by the
	/// \c Font property. If set to \c VARIANT_TRUE, the system font; otherwise the specified font is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseSystemFont, put_Font, putref_Font
	virtual HRESULT STDMETHODCALLTYPE put_UseSystemFont(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \param[out] pValue The control's version.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Version(BSTR* pValue);

	/// \brief <em>Displays the control's credits</em>
	///
	/// Displays some information about this control and its author.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AboutDlg
	virtual HRESULT STDMETHODCALLTYPE About(void);
	/// \brief <em>Enters drag'n'drop mode</em>
	///
	/// \param[in] pDraggedItems The dragged items collection object's \c ITreeViewItemContainer
	///            implementation.
	/// \param[in] hDragImageList The imagelist containing the drag image that shall be used to
	///            visualize the drag'n'drop operation. If -1, the method creates the drag image itself;
	///            if \c NULL, no drag image is used.
	/// \param[in,out] pXHotSpot The x-coordinate (in pixels) of the drag image's hotspot relative to the
	///                drag image's upper-left corner. If the \c hDragImageList parameter is set to -1 or
	///                \c NULL, this parameter is ignored. This parameter will be changed to the value that
	///                finally was used by the method.
	/// \param[in,out] pYHotSpot The y-coordinate (in pixels) of the drag image's hotspot relative to the
	///                drag image's upper-left corner. If the \c hDragImageList parameter is set to -1 or
	///                \c NULL, this parameter is ignored. This parameter will be changed to the value that
	///                finally was used by the method.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OLEDrag, get_DraggedItems, EndDrag, get_hDragImageList, Raise_ItemBeginDrag,
	///     Raise_ItemBeginRDrag, TreeViewItem::CreateDragImage, TreeViewItemContainer::CreateDragImage
	virtual HRESULT STDMETHODCALLTYPE BeginDrag(ITreeViewItemContainer* pDraggedItems, OLE_HANDLE hDragImageList = NULL, OLE_XPOS_PIXELS* pXHotSpot = NULL, OLE_YPOS_PIXELS* pYHotSpot = NULL);
	/// \brief <em>Calculates the maximum number of entirely visible items in the treeview</em>
	///
	/// Retrieves the number of items that fit entirely into the control's client area.
	///
	/// \param[out] pValue The maximum number of simultaneous visible items.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa TreeViewItems::Count
	virtual HRESULT STDMETHODCALLTYPE CountVisible(LONG* pValue);
	/// \brief <em>Creates a new \c TreeViewItemContainer object</em>
	///
	/// Retrieves a new \c TreeViewItemContainer object and fills it with the specified items.
	///
	/// \param[in] items The item(s) to add to the collection. May be either \c Empty, an item ID, a
	///            \c TreeViewItem object or a \c TreeViewItems collection.
	/// \param[out] ppContainer Receives the new collection object's \c ITreeViewItemContainer
	///             implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa TreeViewItemContainer::Clone, TreeViewItemContainer::Add
	virtual HRESULT STDMETHODCALLTYPE CreateItemContainer(VARIANT items = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), ITreeViewItemContainer** ppContainer = NULL);
	/// \brief <em>Exits drag'n'drop mode</em>
	///
	/// \param[in] abort If \c VARIANT_TRUE, the drag'n'drop operation will be treated as aborted;
	///            otherwise it will be treated as a drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DraggedItems, BeginDrag, Raise_AbortedDrag, Raise_Drop
	virtual HRESULT STDMETHODCALLTYPE EndDrag(VARIANT_BOOL abort);
	/// \brief <em>Ends label-editing</em>
	///
	/// Ends editing the item's text.
	///
	/// \param[in] discard Specifies whether to apply or discard the text contained by the label-edit
	///            control at the time this method is called. If set to \c VARIANT_TRUE, the text is
	///            discarded and the edited item's text remains the same; otherwise the edited item's
	///            text is changed to the text contained by the label-edit control at the time this
	///            method is called.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa TreeViewItem::StartLabelEditing, Raise_RenamedItem, Raise_ItemSetText
	virtual HRESULT STDMETHODCALLTYPE EndLabelEdit(VARIANT_BOOL discard);
	/// \brief <em>Finishes a pending drop operation</em>
	///
	/// During a drag'n'drop operation the drag image is displayed until the \c OLEDragDrop event has been
	/// handled. This order is intended by Microsoft Windows. However, if a message box is displayed from
	/// within the \c OLEDragDrop event, or the drop operation cannot be performed asynchronously and takes
	/// a long time, it may be desirable to remove the drag image earlier.\n
	/// This method will break the intended order and finish the drag'n'drop operation (including removal
	/// of the drag image) immediately.
	///
	/// \remarks This method will fail if not called from the \c OLEDragDrop event handler or if no drag
	///          images are used.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_OLEDragDrop, get_SupportOLEDragImages
	virtual HRESULT STDMETHODCALLTYPE FinishOLEDragDrop(void);
	/// \brief <em>Retrieves the position of the control's insertion mark</em>
	///
	/// \param[out] pRelativePosition The insertion mark's position relative to the specified item. The
	///             following values, defined by the \c InsertMarkPositionConstants enumeration, are
	///             valid: \c impBefore, \c impAfter, \c impNowhere.
	/// \param[out] ppTreeItem Receives the \c ITreeViewItem implementation of the item, at which
	///             the insertion mark is shown.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa SetInsertMarkPosition, ExTVwLibU::InsertMarkPositionConstants
	/// \else
	///   \sa SetInsertMarkPosition, ExTVwLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetInsertMarkPosition(InsertMarkPositionConstants* pRelativePosition, ITreeViewItem** ppTreeItem);
	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the control's parts that lie below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in,out] pHitTestDetails Receives a value specifying the exact part of the control the
	///                specified point lies in. Any of the values defined by the \c HitTestConstants
	///                enumeration is valid.
	/// \param[out] ppHitItem Receives the "hit" item's \c ITreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ItemBoundingBoxDefinition, ExTVwLibU::HitTestConstants, HitTest
	/// \else
	///   \sa get_ItemBoundingBoxDefinition, ExTVwLibA::HitTestConstants, HitTest
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, ITreeViewItem** ppHitItem);
	/// \brief <em>Loads the control's settings from the specified file</em>
	///
	/// \param[in] file The file to read from.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveSettingsToFile
	virtual HRESULT STDMETHODCALLTYPE LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Enters OLE drag'n'drop mode</em>
	///
	/// \param[in] pDataObject A pointer to the \c IDataObject implementation to use during OLE
	///            drag'n'drop. If not specified, the control's own implementation is used.
	/// \param[in] supportedEffects A bit field defining all drop effects the client wants to support.
	///            Any combination of the values defined by the \c OLEDropEffectConstants enumeration
	///            (except \c odeScroll) is valid.
	/// \param[in] hWndToAskForDragImage The handle of the window, that is awaiting the
	///            \c DI_GETDRAGIMAGE message to specify the drag image to use. If -1, the method
	///            creates the drag image itself.\n
	///            If \c SupportOLEDragImages is set to \c VARIANT_FALSE, no drag image is used.
	/// \param[in] pDraggedItems The dragged items collection object's \c ITreeViewItemContainer
	///            implementation. It is used to generate the drag image and is ignored if
	///            \c hWndToAskForDragImage is not -1.
	/// \param[in] itemCountToDisplay The number to display in the item count label of Aero drag images.
	///            If set to 0 or 1, no item count label is displayed. If set to -1, the number of items
	///            contained in the \c pDraggedItems collection is displayed in the item count label. If
	///            set to any value larger than 1, this value is displayed in the item count label.
	/// \param[out] pPerformedEffects The performed drop effect. Any of the values defined by the
	///             \c OLEDropEffectConstants enumeration (except \c odeScroll) is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa BeginDrag, Raise_ItemBeginDrag, Raise_ItemBeginRDrag, Raise_OLEStartDrag,
	///       get_SupportOLEDragImages, get_OLEDragImageStyle, ExTVwLibU::OLEDropEffectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	/// \else
	///   \sa BeginDrag, Raise_ItemBeginDrag, Raise_ItemBeginRDrag, Raise_OLEStartDrag,
	///       get_SupportOLEDragImages, get_OLEDragImageStyle, ExTVwLibA::OLEDropEffectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE OLEDrag(LONG* pDataObject = NULL, OLEDropEffectConstants supportedEffects = odeCopyOrMove, OLE_HANDLE hWndToAskForDragImage = -1, ITreeViewItemContainer* pDraggedItems = NULL, LONG itemCountToDisplay = -1, OLEDropEffectConstants* pPerformedEffects = NULL);
	/// \brief <em>Advises the control to redraw itself</em>
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Refresh(void);
	/// \brief <em>Saves the control's settings to the specified file</em>
	///
	/// \param[in] file The file to write to.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadSettingsFromFile
	virtual HRESULT STDMETHODCALLTYPE SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Sets the position of the control's insertion mark</em>
	///
	/// \param[in] relativePosition The insertion mark's position relative to the specified item. Any of
	///            the values defined by the \c InsertMarkPositionConstants enumeration is valid.
	/// \param[in] pTreeItem The \c ITreeViewItem implementation of the item, at which the insertion
	///            mark will be shown. If set to \c NULL, the insertion mark will be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa GetInsertMarkPosition, get_InsertMarkColor, get_AllowDragDrop, get_RegisterForOLEDragDrop,
	///       ExTVwLibU::InsertMarkPositionConstants
	/// \else
	///   \sa GetInsertMarkPosition, get_InsertMarkColor, get_AllowDragDrop, get_RegisterForOLEDragDrop,
	///       ExTVwLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE SetInsertMarkPosition(InsertMarkPositionConstants relativePosition, ITreeViewItem* pTreeItem);
	/// \brief <em>Sorts the control's top-level items</em>
	///
	/// Sorts the control's top-level items by up to 5 criteria.
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
	/// \param[in] recurse If \c VARIANT_FALSE, only the top-level items will be sorted; otherwise
	///            all items will be sorted recursively.
	/// \param[in] caseSensitive If \c VARIANT_TRUE, text comparisons will be case sensitive; otherwise
	///            not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_SortOrder, TreeViewItem::SortSubItems, Raise_CompareItems, ExTVwLibU::SortByConstants
	/// \else
	///   \sa get_SortOrder, TreeViewItem::SortSubItems, Raise_CompareItems, ExTVwLibA::SortByConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE SortItems(SortByConstants firstCriterion = sobShell, SortByConstants secondCriterion = sobText, SortByConstants thirdCriterion = sobNone, SortByConstants fourthCriterion = sobNone, SortByConstants fifthCriterion = sobNone, VARIANT_BOOL recurse = VARIANT_FALSE, VARIANT_BOOL caseSensitive = VARIANT_FALSE);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Property object changes
	///
	//@{
	/// \brief <em>Will be called after a property object was changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that was changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnChanged
	HRESULT OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/);
	/// \brief <em>Will be called before a property object is changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that is about to be changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnRequestEdit
	HRESULT OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Called to create the control window</em>
	///
	/// Called to create the control window. This method overrides \c CWindowImpl::Create() and is
	/// used to customize the window styles.
	///
	/// \param[in] hWndParent The control's parent window.
	/// \param[in] rect The control's bounding rectangle.
	/// \param[in] szWindowName The control's window name.
	/// \param[in] dwStyle The control's window style. Will be ignored.
	/// \param[in] dwExStyle The control's extended window style. Will be ignored.
	/// \param[in] MenuOrID The control's ID.
	/// \param[in] lpCreateParam The window creation data. Will be passed to the created window.
	///
	/// \return The created window's handle.
	///
	/// \sa OnCreate, GetStyleBits, GetExStyleBits
	HWND Create(HWND hWndParent, ATL::_U_RECT rect = NULL, LPCTSTR szWindowName = NULL, DWORD dwStyle = 0, DWORD dwExStyle = 0, ATL::_U_MENUorID MenuOrID = 0U, LPVOID lpCreateParam = NULL);
	/// \brief <em>Called to draw the control</em>
	///
	/// Called to draw the control. This method overrides \c CComControlBase::OnDraw() and is used to prevent
	/// the "ATL 9.0" drawing in user mode and replace it in design mode.
	///
	/// \param[in] drawInfo Contains any details like the device context required for drawing.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT OnDraw(ATL_DRAWINFO& drawInfo);
	/// \brief <em>Called after receiving the last message (typically \c WM_NCDESTROY)</em>
	///
	/// \param[in] hWnd The window being destroyed.
	///
	/// \sa OnCreate, OnDestroy
	void OnFinalMessage(HWND /*hWnd*/);
	/// \brief <em>Informs an embedded object of its display location within its container</em>
	///
	/// \param[in] pClientSite The \c IOleClientSite implementation of the container application's
	///            client side.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms684013.aspx">IOleObject::SetClientSite</a>
	virtual HRESULT STDMETHODCALLTYPE IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite);
	/// \brief <em>Notifies the control when the container's document window is activated or deactivated</em>
	///
	/// ATL's implementation of \c OnDocWindowActivate calls \c IOleInPlaceObject_UIDeactivate if the control
	/// is deactivated. This causes a bug in MDI apps. If the control sits on a \c MDI child window and has
	/// the focus and the title bar of another top-level window (not the MDI parent window) of the same
	/// process is clicked, the focus is moved from the ATL based ActiveX control to the next control on the
	/// MDI child before it is moved to the other top-level window that was clicked. If the focus is set back
	/// to the MDI child, the ATL based control no longer has the focus.
	///
	/// \param[in] fActivate If \c TRUE, the document window is activated; otherwise deactivated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/0kz79wfc.aspx">IOleInPlaceActiveObjectImpl::OnDocWindowActivate</a>
	virtual HRESULT STDMETHODCALLTYPE OnDocWindowActivate(BOOL /*fActivate*/);

	/// \brief <em>A keyboard or mouse message needs to be translated</em>
	///
	/// The control's container calls this method if it receives a keyboard or mouse message. It gives
	/// us the chance to customize keystroke translation (i. e. to react to them in a non-default way).
	/// This method overrides \c CComControlBase::PreTranslateAccelerator.
	///
	/// \param[in] pMessage A \c MSG structure containing details about the received window message.
	/// \param[out] hReturnValue A reference parameter of type \c HRESULT which will be set to \c S_OK,
	///             if the message was translated, and to \c S_FALSE otherwise.
	///
	/// \return \c FALSE if the object's container should translate the message; otherwise \c TRUE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/hxa56938.aspx">CComControlBase::PreTranslateAccelerator</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646373.aspx">TranslateAccelerator</a>
	BOOL PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue);

	//////////////////////////////////////////////////////////////////////
	/// \name Drag image creation
	///
	//@{
	/// \brief <em>Creates a legacy drag image for the specified item</em>
	///
	/// Creates a drag image for the specified item in the style of Windows versions prior to Vista. The
	/// drag image is added to an imagelist which is returned.
	///
	/// \param[in] hItem The item for which to create the drag image.
	/// \param[out] pUpperLeftPoint Receives the coordinates (in pixels) of the drag image's upper-left
	///             corner relative to the control's upper-left corner.
	/// \param[out] pBoundingRectangle Receives the drag image's bounding rectangle (in pixels) relative to
	///             the control's upper-left corner.
	///
	/// \return An imagelist containing the drag image.
	///
	/// \remarks The caller is responsible for destroying the imagelist.
	///
	/// \sa OnCreateDragImage, CreateLegacyOLEDragImage, TreeViewItemContainer::CreateDragImage
	HIMAGELIST CreateLegacyDragImage(HTREEITEM hItem, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle);
	/// \brief <em>Creates a legacy OLE drag image for the specified items</em>
	///
	/// Creates an OLE drag image for the specified items in the style of Windows versions prior to Vista.
	///
	/// \param[in] pItems A \c ITreeViewItemContainer object wrapping the items for which to create the drag
	///            image.
	/// \param[out] pDragImage Receives the drag image including transparency information and the coordinates
	///             (in pixels) of the drag image's upper-left corner relative to the control's upper-left
	///             corner.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa OnGetDragImage, CreateVistaOLEDragImage, CreateLegacyDragImage, TreeViewItemContainer,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759778.aspx">SHDRAGIMAGE</a>
	BOOL CreateLegacyOLEDragImage(ITreeViewItemContainer* pItems, __in LPSHDRAGIMAGE pDragImage);
	/// \brief <em>Creates a Vista-like OLE drag image for the specified items</em>
	///
	/// Creates an OLE drag image for the specified items in the style of Windows Vista and newer.
	///
	/// \param[in] pItems A \c ITreeViewItemContainer object wrapping the items for which to create the drag
	///            image.
	/// \param[out] pDragImage Receives the drag image including transparency information and the coordinates
	///             (in pixels) of the drag image's upper-left corner relative to the control's upper-left
	///             corner.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa OnGetDragImage, CreateLegacyOLEDragImage, CreateLegacyDragImage, TreeViewItemContainer,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759778.aspx">SHDRAGIMAGE</a>
	BOOL CreateVistaOLEDragImage(ITreeViewItemContainer* pItems, __in LPSHDRAGIMAGE pDragImage);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>The control should sort its items</em>
	///
	/// Sorts an item's sub-items by up to 5 criteria.
	///
	/// \param[in] hParentItem The item whose sub-items shall be sorted. If set to \c TVI_ROOT the
	///            control's top-level items will be sorted.
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
	/// \param[in] recurse If \c FALSE, only the item's immediate sub-items will be sorted; otherwise all
	///            sub-items will be sorted recursively.
	/// \param[in] caseSensitive If \c TRUE, text comparisons will be case sensitive; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa sortingSettings, get_SortOrder, SortItems, IItemComparator, CompareItems,
	///       Raise_CompareItems, ExTVwLibU::SortByConstants
	/// \else
	///   \sa sortingSettings, get_SortOrder, SortItems, IItemComparator, CompareItems,
	///       Raise_CompareItems, ExTVwLibA::SortByConstants
	/// \endif
	HRESULT SortSubItems(HTREEITEM hParentItem, SortByConstants firstCriterion = sobShell, SortByConstants secondCriterion = sobText, SortByConstants thirdCriterion = sobNone, SortByConstants fourthCriterion = sobNone, SortByConstants fifthCriterion = sobNone, BOOL recurse = FALSE, BOOL caseSensitive = FALSE);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropTarget
	///
	//@{
	/// \brief <em>Indicates whether a drop can be accepted, and, if so, the effect of the drop</em>
	///
	/// This method is called by the \c DoDragDrop function to determine the target's preferred drop
	/// effect the first time the user moves the mouse into the control during OLE drag'n'Drop. The
	/// target communicates the \c DoDragDrop function the drop effect it wants to be used on drop.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the dragged
	///            data.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragOver, DragLeave, Drop, Raise_OLEDragEnter,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680106.aspx">IDropTarget::DragEnter</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Notifies the target that it no longer is the target of the current OLE drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse out of the
	/// control during OLE drag'n'Drop or if the user canceled the operation. The target must release
	/// any references it holds to the data object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, Drop, Raise_OLEDragLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680110.aspx">IDropTarget::DragLeave</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeave(void);
	/// \brief <em>Communicates the current drop effect to the \c DoDragDrop function</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse over the
	/// control during OLE drag'n'Drop. The target communicates the \c DoDragDrop function the drop
	/// effect it wants to be used on drop.
	///
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, the current drop effect (defined by the \c DROPEFFECT
	///                enumeration). On return, this paramter must be set to the drop effect that the
	///                target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragLeave, Drop, Raise_OLEDragMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680129.aspx">IDropTarget::DragOver</a>
	virtual HRESULT STDMETHODCALLTYPE DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Incorporates the source data into the target and completes the drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user completes the drag'n'drop
	/// operation. The target must incorporate the dragged data into itself and pass the used drop
	/// effect back to the \c DoDragDrop function. The target must release any references it holds to
	/// the data object.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the data
	///            to transfer.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target finally executed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, DragLeave, Raise_OLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687242.aspx">IDropTarget::Drop</a>
	virtual HRESULT STDMETHODCALLTYPE Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropSource
	///
	//@{
	/// \brief <em>Notifies the source of the current drop effect during OLE drag'n'drop</em>
	///
	/// This method is called frequently by the \c DoDragDrop function to notify the source of the
	/// last drop effect that the target has chosen. The source should set an appropriate mouse cursor.
	///
	/// \param[in] effect The drop effect chosen by the target. Any of the values defined by the
	///            \c DROPEFFECT enumeration is valid.
	///
	/// \return \c S_OK if the method has set a custom mouse cursor; \c DRAGDROP_S_USEDEFAULTCURSORS to
	///         use default mouse cursors; or an error code otherwise.
	///
	/// \sa QueryContinueDrag, Raise_OLEGiveFeedback,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693723.aspx">IDropSource::GiveFeedback</a>
	virtual HRESULT STDMETHODCALLTYPE GiveFeedback(DWORD effect);
	/// \brief <em>Determines whether a drag'n'drop operation should be continued, canceled or completed</em>
	///
	/// This method is called by the \c DoDragDrop function to determine whether a drag'n'drop
	/// operation should be continued, canceled or completed.
	///
	/// \param[in] pressedEscape Indicates whether the user has pressed the \c ESC key since the
	///            previous call of this method. If \c TRUE, the key has been pressed; otherwise not.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	///
	/// \return \c S_OK if the drag'n'drop operation should continue; \c DRAGDROP_S_DROP if it should
	///         be completed; \c DRAGDROP_S_CANCEL if it should be canceled; or an error code otherwise.
	///
	/// \sa GiveFeedback, Raise_OLEQueryContinueDrag,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690076.aspx">IDropSource::QueryContinueDrag</a>
	virtual HRESULT STDMETHODCALLTYPE QueryContinueDrag(BOOL pressedEscape, DWORD keyState);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropSourceNotify
	///
	//@{
	/// \brief <em>Notifies the source that the user drags the mouse cursor into a potential drop target window</em>
	///
	/// This method is called by the \c DoDragDrop function to notify the source that the user is dragging
	/// the mouse cursor into a potential drop target window.
	///
	/// \param[in] hWndTarget The potential drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragLeaveTarget, Raise_OLEDragEnterPotentialTarget,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa344050.aspx">IDropSourceNotify::DragEnterTarget</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnterTarget(HWND hWndTarget);
	/// \brief <em>Notifies the source that the user drags the mouse cursor out of a potential drop target window</em>
	///
	/// This method is called by the \c DoDragDrop function to notify the source that the user is dragging
	/// the mouse cursor out of a potential drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnterTarget, Raise_OLEDragLeavePotentialTarget,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa344050.aspx">IDropSourceNotify::DragLeaveTarget</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeaveTarget(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	#ifdef INCLUDESHELLBROWSERINTERFACE
		//////////////////////////////////////////////////////////////////////
		/// \name Implementation of IInternalMessageListener
		///
		//@{
		/// \brief <em>Processes the messages sent by the attached \c ShellBrowser control</em>
		///
		/// \param[in] message The message to process.
		/// \param[in] wParam Used to transfer data with the message.
		/// \param[in] lParam Used to transfer data with the message.
		///
		/// \return An \c HRESULT error code.
		virtual HRESULT ProcessMessage(UINT message, WPARAM wParam, LPARAM lParam);
		//@}
		//////////////////////////////////////////////////////////////////////
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICategorizeProperties
	///
	//@{
	/// \brief <em>Retrieves a category's name</em>
	///
	/// \param[in] category The ID of the category whose name is requested.
	/// \param[in] languageID The locale identifier identifying the language in which name should be
	///            provided.
	/// \param[out] pName The category's name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::GetCategoryName
	virtual HRESULT STDMETHODCALLTYPE GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName);
	/// \brief <em>Maps a property to a category</em>
	///
	/// \param[in] property The ID of the property whose category is requested.
	/// \param[out] pCategory The category's ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::MapPropertyToCategory
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToCategory(DISPID property, PROPCAT* pCategory);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICreditsProvider
	///
	//@{
	/// \brief <em>Retrieves the name of the control's authors</em>
	///
	/// \return A string containing the names of all authors.
	CAtlString GetAuthors(void);
	/// \brief <em>Retrieves the URL of the website that has information about the control</em>
	///
	/// \return A string containing the URL.
	CAtlString GetHomepage(void);
	/// \brief <em>Retrieves the URL of the website where users can donate via Paypal</em>
	///
	/// \return A string containing the URL.
	CAtlString GetPaypalLink(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank especially</em>
	///
	/// \return A string containing the special thanks.
	CAtlString GetSpecialThanks(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank</em>
	///
	/// \return A string containing the thanks.
	CAtlString GetThanks(void);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \return A string containing the version.
	CAtlString GetVersion(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IItemComparator
	///
	//@{
	/// \brief <em>Two items must be compared</em>
	///
	/// This method compares two treeview items.
	///
	/// \param[in] itemID1 The unique ID of the first item to compare.
	/// \param[in] itemID2 The unique ID of the second item to compare.
	///
	/// \return -1 if the first item should precede the second; 0 if the items are equal; 1 if the
	///         second item should precede the first.
	///
	/// \sa sortingSettings, SortSubItems, ::CompareItems
	int CompareItems(LONG itemID1, LONG itemID2);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IPerPropertyBrowsing
	///
	//@{
	/// \brief <em>A displayable string for a property's current value is required</em>
	///
	/// This method is called if the caller's user interface requests a user-friendly description of the
	/// specified property's current value that may be displayed instead of the value itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display name is requested.
	/// \param[out] pDescription The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688734.aspx">IPerPropertyBrowsing::GetDisplayString</a>
	virtual HRESULT STDMETHODCALLTYPE GetDisplayString(DISPID property, BSTR* pDescription);
	/// \brief <em>Displayable strings for a property's predefined values are required</em>
	///
	/// This method is called if the caller's user interface requests user-friendly descriptions of the
	/// specified property's predefined values that may be displayed instead of the values itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display names are requested.
	/// \param[in,out] pDescriptions A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of strings. This array will be
	///                filled with the display name strings.
	/// \param[in,out] pCookies A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of \c DWORD values. Each \c DWORD
	///                value identifies a predefined value of the property.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedValue, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms679724.aspx">IPerPropertyBrowsing::GetPredefinedStrings</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies);
	/// \brief <em>A property's predefined value identified by a token is required</em>
	///
	/// This method is called if the caller's user interface requires a property's predefined value that
	/// it has the token of. The token was returned by the \c GetPredefinedStrings method.
	/// We use this method for enumeration-type properties to transform strings like "1 - At Root"
	/// back to the underlying enumeration value (here: \c lsLinesAtRoot).
	///
	/// \param[in] property The ID of the property for which a predefined value is requested.
	/// \param[in] cookie Token identifying which value to return. The token was previously returned
	///            in the \c pCookies array filled by \c IPerPropertyBrowsing::GetPredefinedStrings.
	/// \param[out] pPropertyValue A \c VARIANT that will receive the predefined value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690401.aspx">IPerPropertyBrowsing::GetPredefinedValue</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue);
	/// \brief <em>A property's property page is required</em>
	///
	/// This method is called to request the \c CLSID of the property page used to edit the specified
	/// property.
	///
	/// \param[in] property The ID of the property whose property page is requested.
	/// \param[out] pPropertyPage The property page's \c CLSID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms694476.aspx">IPerPropertyBrowsing::MapPropertyToPage</a>
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToPage(DISPID property, CLSID* pPropertyPage);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Retrieves a displayable string for a specified setting of a specified property</em>
	///
	/// Retrieves a user-friendly description of the specified property's specified setting. This
	/// description may be displayed by the caller's user interface instead of the setting itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property for which to retrieve the display name.
	/// \param[in] cookie Token identifying the setting for which to retrieve a description.
	/// \param[out] description The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedStrings, GetResStringWithNumber
	HRESULT GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISpecifyPropertyPages
	///
	//@{
	/// \brief <em>The property pages to show are required</em>
	///
	/// This method is called if the property pages, that may be displayed for this object, are required.
	///
	/// \param[out] pPropertyPages A caller-allocated, counted array structure containing the element
	///             count and address of a callee-allocated array of \c GUID structures. Each \c GUID
	///             structure identifies a property page to display.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CommonProperties,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687276.aspx">ISpecifyPropertyPages::GetPages</a>
	virtual HRESULT STDMETHODCALLTYPE GetPages(CAUUID* pPropertyPages);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Message handlers
	///
	//@{
	/// \brief <em>\c WM_CHAR handler</em>
	///
	/// Will be called if a \c WM_KEYDOWN message was translated by \c TranslateMessage.
	/// We use this handler to raise the \c KeyPress event.
	///
	/// \sa OnKeyDown, OnEditChar, Raise_KeyPress,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_CHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_CONTEXTMENU handler</em>
	///
	/// Will be called if the control's context menu should be displayed.
	/// We use this handler to raise the \c ContextMenu event.
	///
	/// \sa OnRButtonDown, OnEditContextMenu, Raise_ContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647592.aspx">WM_CONTEXTMENU</a>
	LRESULT OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_CREATE handler</em>
	///
	/// Will be called right after the control was created.
	/// We use this handler to subclass the parent window, to configure the control window and to raise the
	/// \c RecreatedControlWindow event.
	///
	/// \sa OnDestroy, OnFinalMessage, Raise_RecreatedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632619.aspx">WM_CREATE</a>
	LRESULT OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_CTLCOLOREDIT handler</em>
	///
	/// Will be called if an edit control is about to be drawn.
	/// We use this handler to set the colors (background and text) for the label-edit control.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms672119.aspx">WM_CTLCOLOREDIT</a>
	LRESULT OnCtlColorEdit(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the control is being destroyed.
	/// We use this handler to tidy up and to raise the \c DestroyedControlWindow event.
	///
	/// \sa OnCreate, OnFinalMessage, Raise_DestroyedControlWindow, OnEditDestroy,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_INPUTLANGCHANGE handler</em>
	///
	/// Will be called after an application's input language has been changed.
	/// We use this handler to update the IME mode of the control itself and the label-edit control.
	///
	/// \sa get_IMEMode, get_EditIMEMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632629.aspx">WM_INPUTLANGCHANGE</a>
	LRESULT OnInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the control has the keyboard focus.
	/// We use this handler mainly for the \c SingleExpand feature and to raise the \c KeyDown event.
	///
	/// \sa OnKeyUp, OnChar, OnEditKeyDown, Raise_KeyDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_KEYUP handler</em>
	///
	/// Will be called if a nonsystem key is released while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyUp event.
	///
	/// \sa OnKeyDown, OnChar, OnEditKeyUp, Raise_KeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646281.aspx">WM_KEYUP</a>
	LRESULT OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_LBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the left mouse
	/// button.
	/// We use this handler to raise the \c DblClick event.
	///
	/// \sa OnDblClkNotification, OnMButtonDblClk, OnRDblClkNotification, OnRButtonDblClk, OnXButtonDblClk,
	///     OnEditLButtonDblClk, Raise_DblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645606.aspx">WM_LBUTTONDBLCLK</a>
	LRESULT OnLButtonDblClk(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler mainly for the \c SingleExpand feature and to raise the \c MouseDown event.
	///
	/// \sa OnClickNotification, OnMButtonDown, OnRButtonDown, OnXButtonDown, OnEditLButtonDown,
	///     Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnLButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnMButtonUp, OnRButtonUp, OnXButtonUp, OnEditLButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnLButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the middle mouse
	/// button.
	/// We use this handler to raise the \c MDblClick event.
	///
	/// \sa OnDblClkNotification, OnLButtonDblClk, OnRDblClkNotification, OnRButtonDblClk, OnXButtonDblClk,
	///     OnEditMButtonDblClk, Raise_MDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnRButtonDown, OnXButtonDown, OnEditMButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnRButtonUp, OnXButtonUp, OnEditMButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEACTIVATE handler</em>
	///
	/// Will be called if the control is inactive and the user clicked in its client area.
	/// We use this handler to hide the message from \c CComControl, so that activating the control by
	/// clicking on its caret item doesn't start label-editing mode.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms645612.aspx">WM_MOUSEACTIVATE</a>
	LRESULT OnMouseActivate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_MOUSEHOVER handler</em>
	///
	/// Will be called if the mouse cursor has been located over the control's client area for a previously
	/// specified number of milliseconds.
	/// We use this handler to raise the \c MouseHover event.
	///
	/// \sa OnEditMouseHover, Raise_MouseHover,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645613.aspx">WM_MOUSEHOVER</a>
	LRESULT OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the control's client area.
	/// We use this handler to raise the \c MouseLeave event.
	///
	/// \sa OnEditMouseLeave, Raise_MouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// client area.
	/// We use this handler to raise the \c MouseMove event.
	///
	/// \sa OnLButtonDown, OnLButtonUp, OnEditMouseMove, OnMouseWheel, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEWHEEL and \c WM_MOUSEHWHEEL handler</em>
	///
	/// Will be called if the user rotates the mouse wheel while the mouse cursor is located over the
	/// control's client area.
	/// We use this handler to raise the \c MouseWheel event.
	///
	/// \sa OnEditMouseWheel, OnMouseMove, Raise_MouseWheel,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645617.aspx">WM_MOUSEWHEEL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645614.aspx">WM_MOUSEHWHEEL</a>
	LRESULT OnMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NOTIFY handler</em>
	///
	/// Will be called if the control receives a notification from a child-window (probably the tooltip
	/// control).
	/// We use this handler to allow the client to cancel info tip popups.
	///
	/// \sa OnToolTipGetDispInfoNotificationW,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672614.aspx">WM_NOTIFY</a>
	LRESULT OnNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_PAINT and \c WM_PRINTCLIENT handler</em>
	///
	/// Will be called if the control needs to be drawn.
	/// We use this handler to avoid the control being drawn by \c CComControl. This makes Vista's graphic
	/// effects work.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms534901.aspx">WM_PAINT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534913.aspx">WM_PRINTCLIENT</a>
	LRESULT OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_PARENTNOTIFY handler</em>
	///
	/// Will be called if the label-edit control is created or destroyed.
	/// We use this handler to raise the \c CreatedEditControlWindow event.
	///
	/// \sa OnBeginLabelEditNotification, Raise_CreatedEditControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632638.aspx">WM_PARENTNOTIFY</a>
	LRESULT OnParentNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the right mouse
	/// button.
	/// We use this handler to raise the \c RDblClick event.
	///
	/// \sa OnRDblClkNotification, OnLButtonDblClk, OnDblClkNotification, OnMButtonDblClk, OnXButtonDblClk,
	///     OnEditRButtonDblClk, Raise_RDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646241.aspx">WM_RBUTTONDBLCLK</a>
	LRESULT OnRButtonDblClk(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler mainly to raise the \c MouseDown event.
	///
	/// \sa OnRClickNotification, OnLButtonDown, OnMButtonDown, OnXButtonDown, OnEditRButtonDown,
	///     Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnXButtonUp, OnEditRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFONT handler</em>
	///
	/// Will be called if the control's font shall be set.
	/// We use this handler to synchronize our font settings with the new font.
	///
	/// \sa get_Font,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632642.aspx">WM_SETFONT</a>
	LRESULT OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETREDRAW handler</em>
	///
	/// Will be called if the control's redraw state shall be changed.
	/// We use this handler for proper handling of the \c DontRedraw property.
	///
	/// \sa get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534853.aspx">WM_SETREDRAW</a>
	LRESULT OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETTINGCHANGE handler</em>
	///
	/// Will be called if a system setting was changed.
	/// We use this handler to update our appearance.
	///
	/// \attention This message is posted to top-level windows only, so actually we'll never receive it.
	///
	/// \sa OnThemeChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms725497.aspx">WM_SETTINGCHANGE</a>
	LRESULT OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_THEMECHANGED handler</em>
	///
	/// Will be called on themable systems if the theme was changed.
	/// We use this handler to update our appearance.
	///
	/// \sa OnSettingChange, Flags::usingThemes,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632650.aspx">WM_THEMECHANGED</a>
	LRESULT OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_TIMER handler</em>
	///
	/// Will be called when a timer expires that's associated with the control.
	/// We use this handler for various things, e. g. for the \c CaretChangedDelayTime feature.
	///
	/// \sa get_CaretChangedDelayTime,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644902.aspx">WM_TIMER</a>
	LRESULT OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_WINDOWPOSCHANGED handler</em>
	///
	/// Will be called if the control was moved.
	/// We use this handler to resize the control on COM level and to raise the \c ResizedControlWindow
	/// event.
	///
	/// \sa Raise_ResizedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632652.aspx">WM_WINDOWPOSCHANGED</a>
	LRESULT OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using one of the extended
	/// mouse buttons.
	/// We use this handler to raise the \c XDblClick event.
	///
	/// \sa OnDblClkNotification, OnLButtonDblClk, OnMButtonDblClk, OnRDblClkNotification,
	///     OnEditXButtonDblClk, Raise_XDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646244.aspx">WM_XBUTTONDBLCLK</a>
	LRESULT OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnRButtonDown, OnEditXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnRButtonUp, OnEditXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NOTIFY handler</em>
	///
	/// Will be called if the control's parent window receives a notification from the control.
	/// We use this handler for the custom draw support.
	///
	/// \sa OnCustomDrawNotification, OnReflectedNotifyFormat,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672614.aspx">WM_NOTIFY</a>
	LRESULT OnReflectedNotify(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NOTIFYFORMAT handler</em>
	///
	/// Will be called if the control asks its parent window which format (Unicode/ANSI) the \c WM_NOTIFY
	/// notifications should have.
	/// We use this handler for proper Unicode support.
	///
	/// \sa OnReflectedNotify,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672615.aspx">WM_NOTIFYFORMAT</a>
	LRESULT OnReflectedNotifyFormat(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	#ifdef INCLUDESHELLBROWSERINTERFACE
		/// \brief <em>\c SHBM_HANDSHAKE handler</em>
		///
		/// Will be called if a \c ShellTreeView control wants to attach to us.
		///
		/// \sa IMessageListener, IInternalMessageListener, IExplorerTreeView, SHBM_HANDSHAKE
		LRESULT OnHandshake(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	#endif
	/// \brief <em>\c DI_GETDRAGIMAGE handler</em>
	///
	/// Will be called during OLE drag'n'drop if the control is queried for a drag image.
	///
	/// \sa OLEDrag, CreateLegacyOLEDragImage, CreateVistaOLEDragImage,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	LRESULT OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c TVM_CREATEDRAGIMAGE handler</em>
	///
	/// Will be called if the control shall create a drag image.
	/// We use this handler to work-around a bug in \c SysTreeView32's implementation of
	/// \c TVM_CREATEDRAGIMAGE.
	///
	/// \sa OnCustomDrawNotification, CreateLegacyDragImage,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672615.aspx">TVM_CREATEDRAGIMAGE</a>,
	///     <a href="https://groups.google.de/group/microsoft.public.win32.programmer.ui/browse_frm/thread/2196a945786889bf/">Bug description</a>
	LRESULT OnCreateDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVM_DELETEITEM handler</em>
	///
	/// Will be called if a treeview item shall be removed from the control.
	/// We use this handler mainly to raise the \c RemovingItem and \c RemovedItem events.
	///
	/// \sa OnDeleteItemNotification, OnInsertItem, Raise_RemovingItem, Raise_RemovedItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650156.aspx">TVM_DELETEITEM</a>
	LRESULT OnDeleteItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVM_EDITLABEL handler</em>
	///
	/// Will be called if the control shall enter label-edit mode.
	/// We use this handler to set the \c LabelEditStatus::watchingForEdit flag.
	///
	/// \sa OnBeginLabelEditNotification, LabelEditStatus::watchingForEdit,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650157.aspx">TVM_EDITLABEL</a>
	LRESULT OnEditLabel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVM_EXPANDITEM handler</em>
	///
	/// Will be called if a treeview item shall be collapsed or expanded.
	/// We use this handler mainly to raise the \c CollapsingItem, \c CollapsedItem, \c ExpandingItem and
	/// \c ExpandedItem events.
	///
	/// \sa OnItemExpandingNotification, OnItemExpandedNotification, Raise_CollapsingItem,
	///     Raise_CollapsedItem, Raise_ExpandingItem, Raise_ExpandedItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650160.aspx">TVM_EXPAND</a>
	LRESULT OnExpandItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVM_GETITEM handler</em>
	///
	/// Will be called if some or all of a treeview item's attributes are requested from the control.
	/// We use this handler to hide the item's unique ID from the caller and return the item's
	/// associated data out of the \c itemParams hash map instead.
	///
	/// \sa itemParams, OnSetItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650168.aspx">TVM_GETITEM</a>
	LRESULT OnGetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVM_INSERTITEM handler</em>
	///
	/// Will be called if a new treeview item shall be inserted into the control.
	/// We use this handler mainly to raise the \c InsertingItem and \c InsertedItem events.
	///
	/// \sa OnDeleteItem, Raise_InsertingItem, Raise_InsertedItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650180.aspx">TVM_INSERTITEM</a>
	LRESULT OnInsertItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVM_SELECTITEM handler</em>
	///
	/// Will be called if any kind of item selection (\c Caret, \c DropHilite, \c FirstVisible) in the
	/// control shall be changed.
	/// We use this handler to set the \c LabelEditStatus::watchingForEdit flag.
	///
	/// \sa OnCaretChangingNotification, OnCaretChangedNotification, get_CaretItem, get_FirstVisibleItem,
	///     LabelEditStatus::watchingForEdit,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650183.aspx">TVM_SELECTITEM</a>
	LRESULT OnSelectItem(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c TVM_SETAUTOSCROLLINFO handler</em>
	///
	/// Will be called if the control's automatic scrolling features shall be changed.
	/// We use this handler to synchronize our cached settings.
	///
	/// \sa get_AutoHScrollPixelsPerSecond, get_AutoHScrollRedrawInterval,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa361686.aspx">TVM_SETAUTOSCROLLINFO</a>
	LRESULT OnSetAutoScrollInfo(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVM_SETBORDER handler</em>
	///
	/// Will be called if the control's item borders shall be changed.
	/// We use this handler to synchronize our cached settings.
	///
	/// \sa cachedSettings, get_ItemXBorder, get_ItemYBorder, TVM_SETBORDER
	LRESULT OnSetBorder(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVM_SETIMAGELIST handler</em>
	///
	/// Will be called if one of the control's imagelists shall be changed.
	/// We use this handler to synchronize our cached settings.
	///
	/// \sa cachedSettings, get_hImageList,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650185.aspx">TVM_SETIMAGELIST</a>
	LRESULT OnSetImageList(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVM_SETINSERTMARK handler</em>
	///
	/// Will be called if the control's insertion mark shall be repositioned.
	/// We use this handler to synchronize the \c insertMark variable.
	///
	/// \sa insertMark, GetInsertMarkPosition,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650187.aspx">TVM_SETINSERTMARK</a>
	LRESULT OnSetInsertMark(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c TVM_SETITEM handler</em>
	///
	/// Will be called if the control is requested to update some or all of a treeview item's attributes.
	/// We use this handler to avoid the item's unique ID is overwritten and store the item's
	/// associated data in the \c itemParams hash map instead.
	///
	/// \sa itemParams, OnGetItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650189.aspx">TVM_SETITEM</a>
	LRESULT OnSetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVM_SETITEMHEIGHT handler</em>
	///
	/// Will be called if the control's default item height shall be changed.
	/// We use this handler to synchronize our cached settings.
	///
	/// \sa cachedSettings, get_ItemHeight,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650190.aspx">TVM_SETITEMHEIGHT</a>
	LRESULT OnSetItemHeight(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_CHAR handler</em>
	///
	/// Will be called if a \c WM_KEYDOWN message was translated by \c TranslateMessage.
	/// We use this handler to raise the \c EditKeyPress event.
	///
	/// \sa OnEditChange, OnEditKeyDown, OnChar, Raise_EditKeyPress,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_CHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnEditChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_CONTEXTMENU handler</em>
	///
	/// Will be called if the label-edit control's context menu should be displayed.
	/// We use this handler to raise the \c EditContextMenu event.
	///
	/// \sa OnEditRButtonDown, OnContextMenu, Raise_EditContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647592.aspx">WM_CONTEXTMENU</a>
	LRESULT OnEditContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the label-edit control is being destroyed.
	/// We use this handler to raise the \c DestroyedEditControlWindow event.
	///
	/// \sa OnEndLabelEditNotification, OnDestroy, Raise_DestroyedEditControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnEditDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the label-edit control has the keyboard focus.
	/// We use this handler to raise the \c EditKeyDown event.
	///
	/// \sa OnEditKeyUp, OnEditChar, OnKeyDown, Raise_EditKeyDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnEditKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYUP handler</em>
	///
	/// Will be called if a nonsystem key is released while the label-edit control has the keyboard focus.
	/// We use this handler to raise the \c EditKeyUp event.
	///
	/// \sa OnEditKeyDown, OnEditChar, OnKeyUp, Raise_EditKeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646281.aspx">WM_KEYUP</a>
	LRESULT OnEditKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KILLFOCUS handler</em>
	///
	/// Will be called immediately before the label-edit control loses the keyboard focus.
	/// We use this handler to raise the \c EditLostFocus event.
	///
	/// \sa OnEditSetFocus, Raise_EditLostFocus,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646282.aspx">WM_KILLFOCUS</a>
	LRESULT OnEditKillFocus(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the label-edit control's client area using the
	/// left mouse button.
	/// We use this handler to raise the \c EditDblClick event.
	///
	/// \sa OnEditMButtonDblClk, OnEditRButtonDblClk, OnEditXButtonDblClk, OnDblClkNotification,
	///     Raise_EditDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645606.aspx">WM_LBUTTONDBLCLK</a>
	LRESULT OnEditLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the label-edit control's client area.
	/// We use this handler to raise the \c EditMouseDown event.
	///
	/// \sa OnEditMButtonDown, OnEditRButtonDown, OnEditXButtonDown, OnLButtonDown, Raise_EditMouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnEditLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the label-edit control's client area.
	/// We use this handler to raise the \c EditMouseUp event.
	///
	/// \sa OnEditMButtonUp, OnEditRButtonUp, OnEditXButtonUp, OnLButtonUp, Raise_EditMouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnEditLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the label-edit control's client area using the
	/// middle mouse button.
	/// We use this handler to raise the \c EditMDblClick event.
	///
	/// \sa OnEditLButtonDblClk, OnEditRButtonDblClk, OnEditXButtonDblClk, OnMButtonDblClk,
	///     Raise_EditMDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnEditMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the label-edit control's client area.
	/// We use this handler to raise the \c EditMouseDown event.
	///
	/// \sa OnEditLButtonDown, OnEditRButtonDown, OnEditXButtonDown, OnMButtonDown, Raise_EditMouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnEditMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the label-edit control's client area.
	/// We use this handler to raise the \c EditMouseUp event.
	///
	/// \sa OnEditLButtonUp, OnEditRButtonUp, OnEditXButtonUp, OnMButtonUp, Raise_EditMouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnEditMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEHOVER handler</em>
	///
	/// Will be called if the mouse cursor has been located over the label-edit control's client area
	/// for a previously specified number of milliseconds.
	/// We use this handler to raise the \c EditMouseHover event.
	///
	/// \sa OnMouseHover, Raise_EditMouseHover,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645613.aspx">WM_MOUSEHOVER</a>
	LRESULT OnEditMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the label-edit control's client area.
	/// We use this handler to raise the \c EditMouseLeave event.
	///
	/// \sa OnMouseLeave, Raise_EditMouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnEditMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the label-edit
	/// control's client area.
	/// We use this handler to raise the \c EditMouseMove event.
	///
	/// \sa OnEditLButtonDown, OnEditLButtonUp, OnMouseMove, OnEditMouseWheel, Raise_EditMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnEditMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEWHEEL and \c WM_MOUSEHWHEEL handler</em>
	///
	/// Will be called if the user rotates the mouse wheel while the mouse cursor is located over the
	/// contained edit control's client area.
	/// We use this handler to raise the \c EditMouseWheel event.
	///
	/// \sa OnMouseWheel, OnEditMouseMove, Raise_EditMouseWheel,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645617.aspx">WM_MOUSEWHEEL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645614.aspx">WM_MOUSEHWHEEL</a>
	LRESULT OnEditMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the label-edit control's client area using the
	/// right mouse button.
	/// We use this handler to raise the \c EditRDblClick event.
	///
	/// \sa OnEditLButtonDblClk, OnEditMButtonDblClk, OnEditXButtonDblClk, OnRDblClkNotification,
	///     Raise_EditRDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646241.aspx">WM_RBUTTONDBLCLK</a>
	LRESULT OnEditRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the label-edit control's client area.
	/// We use this handler to raise the \c EditMouseDown event.
	///
	/// \sa OnEditLButtonDown, OnEditMButtonDown, OnEditXButtonDown, OnRButtonDown, Raise_EditMouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnEditRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the label-edit control's client area.
	/// We use this handler to raise the \c EditMouseUp event.
	///
	/// \sa OnEditLButtonUp, OnEditMButtonUp, OnEditXButtonUp, OnRButtonUp, Raise_EditMouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnEditRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFOCUS handler</em>
	///
	/// Will be called after the label-edit control gained the keyboard focus.
	/// We use this handler to raise the \c EditGotFocus event and to initialize IME for the edit control.
	///
	/// \sa OnEditKillFocus, OnSetFocusNotification, get_EditIMEMode, Raise_EditGotFocus,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646283.aspx">WM_SETFOCUS</a>
	LRESULT OnEditSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the contained edit control's client area using one of
	/// the extended mouse buttons.
	/// We use this handler to raise the \c EditXDblClick event.
	///
	/// \sa OnEditLButtonDblClk, OnEditMButtonDblClk, OnEditRButtonDblClk, Raise_EditXDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646244.aspx">WM_XBUTTONDBLCLK</a>
	LRESULT OnEditXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the contained edit control's client area.
	/// We use this handler to raise the \c EditMouseDown event.
	///
	/// \sa OnEditLButtonDown, OnEditMButtonDown, OnEditRButtonDown, OnXButtonDown, Raise_EditMouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnEditXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the contained edit control's client area.
	/// We use this handler to raise the \c EditMouseUp event.
	///
	/// \sa OnEditLButtonUp, OnEditMButtonUp, OnEditRButtonUp, OnXButtonUp, Raise_EditMouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnEditXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	#ifdef INCLUDESHELLBROWSERINTERFACE
		//////////////////////////////////////////////////////////////////////
		/// \name Internal message handlers
		///
		//@{
		/// \brief <em>Handles the \c EXTVM_ADDITEM message</em>
		///
		/// \sa EXTVM_ADDITEM
		HRESULT OnInternalAddItem(UINT /*message*/, WPARAM wParam, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_ITEMHANDLETOID message</em>
		///
		/// \sa EXTVM_ITEMHANDLETOID
		HRESULT OnInternalItemHandleToID(UINT /*message*/, WPARAM wParam, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_IDTOITEMHANDLE message</em>
		///
		/// \sa EXTVM_IDTOITEMHANDLE
		HRESULT OnInternalIDToItemHandle(UINT /*message*/, WPARAM wParam, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_GETITEMBYHANDLE message</em>
		///
		/// \sa EXTVM_GETITEMBYHANDLE
		HRESULT OnInternalGetItemByHandle(UINT /*message*/, WPARAM wParam, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_CREATEITEMCONTAINER message</em>
		///
		/// \sa EXTVM_CREATEITEMCONTAINER
		HRESULT OnInternalCreateItemContainer(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_GETITEMHANDLESFROMVARIANT message</em>
		///
		/// \sa EXTVM_GETITEMHANDLESFROMVARIANT
		HRESULT OnInternalGetItemHandlesFromVariant(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_SETITEMICONINDEX message</em>
		///
		/// \sa EXTVM_SETITEMICONINDEX
		HRESULT OnInternalSetItemIconIndex(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_SETITEMSELECTEDICONINDEX message</em>
		///
		/// \sa EXTVM_SETITEMSELECTEDICONINDEX
		HRESULT OnInternalSetItemSelectedIconIndex(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_SETITEMEXPANDEDICONINDEX message</em>
		///
		/// \sa EXTVM_SETITEMEXPANDEDICONINDEX
		HRESULT OnInternalSetItemExpandedIconIndex(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_HITTEST message</em>
		///
		/// \sa EXTVM_HITTEST
		HRESULT OnInternalHitTest(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_SORTITEMS message</em>
		///
		/// \sa EXTVM_SORTITEMS
		HRESULT OnInternalSortItems(UINT /*message*/, WPARAM wParam, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_SETDROPHILITEDITEM message</em>
		///
		/// \sa EXTVM_SETDROPHILITEDITEM
		HRESULT OnInternalSetDropHilitedItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_FIREDRAGDROPEVENT message</em>
		///
		/// \sa EXTVM_FIREDRAGDROPEVENT
		HRESULT OnInternalFireDragDropEvent(UINT /*message*/, WPARAM wParam, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_OLEDRAG message</em>
		///
		/// \sa EXTVM_OLEDRAG
		HRESULT OnInternalOLEDrag(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_GETIMAGELIST message</em>
		///
		/// \sa EXTVM_GETIMAGELIST
		HRESULT OnInternalGetImageList(UINT /*message*/, WPARAM wParam, LPARAM lParam);
		/// \brief <em>Handles the \c EXTVM_SETIMAGELIST message</em>
		///
		/// \sa EXTVM_SETIMAGELIST
		HRESULT OnInternalSetImageList(UINT /*message*/, WPARAM wParam, LPARAM lParam);
		//@}
		//////////////////////////////////////////////////////////////////////
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	/// \brief <em>\c NM_CLICK handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user clicked into the control's
	/// client area using the left mouse button.
	/// We use this handler to raise the \c Click event.
	///
	/// \sa OnLButtonDown, OnRClickNotification, Raise_Click,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650031.aspx">NM_CLICK (tree view)</a>
	LRESULT OnClickNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c NM_DBLCLK handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user double-clicked into the
	/// control's client area using the left mouse button.
	/// We use this handler to raise the \c DblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnRDblClkNotification, OnRButtonDblClk, OnXButtonDblClk,
	///     OnEditLButtonDblClk, Raise_DblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650134.aspx">NM_DBLCLK (tree view)</a>
	LRESULT OnDblClkNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& wasHandled);
	/// \brief <em>\c NM_RCLICK handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user clicked into the control's
	/// client area using the right mouse button.
	/// We use this handler to raise the \c RClick event.
	///
	/// \sa OnRButtonDown, OnClickNotification, Raise_RClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650136.aspx">NM_RCLICK (tree view)</a>
	LRESULT OnRClickNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& wasHandled);
	/// \brief <em>\c NM_RDBLCLK handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user double-clicked into the
	/// control's client area using the right mouse button.
	/// We use this handler to raise the \c RDblClick event.
	///
	/// \sa OnRButtonDblClk, OnLButtonDblClk, OnDblClkNotification, OnMButtonDblClk, OnXButtonDblClk,
	///     OnEditRButtonDblClk, Raise_RDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650137.aspx">NM_RDBLCLK (tree view)</a>
	LRESULT OnRDblClkNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& wasHandled);
	/// \brief <em>\c NM_SETCURSOR handler</em>
	///
	/// Will be called if the control's parent window is notified, that it may customize the control's mouse
	/// cursor.
	/// We use this handler to set the mouse cursor.
	///
	/// \sa get_MouseIcon, get_MousePointer,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb773493.aspx">NM_SETCURSOR (tree view)</a>
	LRESULT OnSetCursorNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c NM_SETFOCUS handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control has gained the keyboard
	/// focus.
	/// We use this handler to initialize IME.
	///
	/// \sa OnEditSetFocus, get_IMEMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650140.aspx">NM_SETFOCUS (tree view)</a>
	LRESULT OnSetFocusNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& wasHandled);
	/// \brief <em>\c NM_TVSTATEIMAGECHANGING handler</em>
	///
	/// Will be called if the control's parent window is notified, that an item's state image is about to
	/// change.
	/// We use this handler on comctl32.dll version 6.10 and newer to raise the \c ItemStateImageChanging
	/// event.
	///
	/// \sa Raise_ItemStateImageChanging,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa361891.aspx">NM_TVSTATEIMAGECHANGING</a>
	LRESULT OnTVStateImageChangingNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVN_ASYNCDRAW handler</em>
	///
	/// Will be called if the control's parent window is notified, that drawing a treeview item's associated
	/// image has failed.
	/// We use this handler to raise the \c ItemAsynchronousDrawFailed event.
	///
	/// \sa Raise_ItemAsynchronousDrawFailed,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa361681.aspx">TVN_ASYNCDRAW</a>
	LRESULT OnAsyncDrawNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVN_BEGINDRAG handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user wants to drag a treeview
	/// item using the left mouse button.
	/// We use this handler to raise the \c ItemBeginDrag event.
	///
	/// \sa OnBeginRDragNotification, Raise_ItemBeginDrag,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650141.aspx">TVN_BEGINDRAG</a>
	LRESULT OnBeginDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVN_BEGINLABELEDIT handler</em>
	///
	/// Will be called if the control's parent window is notified, that label-edit mode is about to be
	/// entered.
	/// We use this handler to raise the \c StartingLabelEditing event.
	///
	/// \sa OnEndLabelEditNotification, OnParentNotify, Raise_StartingLabelEditing,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650142.aspx">TVN_BEGINLABELEDIT</a>
	LRESULT OnBeginLabelEditNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled);
	/// \brief <em>\c TVN_BEGINRDRAG handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user wants to drag a treeview
	/// item using the right mouse button.
	/// We use this handler to raise the \c ItemBeginRDrag event.
	///
	/// \sa OnBeginDragNotification, Raise_ItemBeginRDrag,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650143.aspx">TVN_BEGINRDRAG</a>
	LRESULT OnBeginRDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVN_DELETEITEM handler</em>
	///
	/// Will be called if the control's parent window is notified, that a treeview item is being removed.
	/// We use this handler to raise the \c FreeItemData event.
	///
	/// \sa OnDeleteItem, Raise_FreeItemData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650144.aspx">TVN_DELETEITEM</a>
	LRESULT OnDeleteItemNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled);
	/// \brief <em>\c TVN_ENDLABELEDIT handler</em>
	///
	/// Will be called if the control's parent window is notified, that label-edit mode was left.
	/// We use this handler to raise the \c RenamingItem event.
	///
	/// \sa OnBeginLabelEditNotification, Raise_RenamingItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650145.aspx">TVN_ENDLABELEDIT</a>
	LRESULT OnEndLabelEditNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVN_GETDISPINFO handler</em>
	///
	/// Will be called if the control requests display information about a treeview item from its parent
	/// window.
	/// We use this handler to raise the \c ItemGetDisplayInfo event.
	///
	/// \sa OnSetDispInfoNotification, Raise_ItemGetDisplayInfo,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650146.aspx">TVN_GETDISPINFO</a>
	LRESULT OnGetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVN_GETINFOTIP handler</em>
	///
	/// Will be called if the control requests a treeview item's tooltip text from its parent window.
	/// We use this handler to raise the \c ItemGetInfoTipText event.
	///
	/// \sa Raise_ItemGetInfoTipText,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650147.aspx">TVN_GETINFOTIP</a>
	LRESULT OnGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVN_ITEMCHANGED handler</em>
	///
	/// Will be called if the control's parent window is notified, that some of the attributes of the
	/// specified treeview item have been changed.
	/// We use this handler to raise the \c ItemSelectionChanged and \c ItemStateImageChanged events.
	///
	/// \sa OnItemChangingNotification, Raise_ItemSelectionChanged, Raise_ItemStateImageChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa361682.aspx">TVN_ITEMCHANGED</a>
	LRESULT OnItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVN_ITEMCHANGING handler</em>
	///
	/// Will be called if the control's parent window is notified, that some of the attributes of the
	/// specified treeview item are about to change.
	/// We use this handler to raise the \c ItemSelectionChanging and \c ItemStateImageChanging events.
	///
	/// \sa OnItemChangedNotification, Raise_ItemSelectionChanging, Raise_ItemStateImageChanging,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa361683.aspx">TVN_ITEMCHANGING</a>
	LRESULT OnItemChangingNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVN_ITEMEXPANDED handler</em>
	///
	/// Will be called if the control's parent window is notified, that a treeview item's sub-items were
	/// collapsed or expanded.
	/// We use this handler mainly to raise the \c CollapsedItem and \c ExpandedItem events.
	///
	/// \sa OnItemExpandingNotification, OnExpandItem, Raise_CollapsedItem, Raise_ExpandedItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650148.aspx">TVN_ITEMEXPANDED</a>
	LRESULT OnItemExpandedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled);
	/// \brief <em>\c TVN_ITEMEXPANDING handler</em>
	///
	/// Will be called if the control's parent window is notified, that a treeview item's sub-items are
	/// about to be collapsed or expanded.
	/// We use this handler to raise the \c CollapsingItem and \c ExpandingItem events.
	///
	/// \sa OnItemExpandedNotification, OnExpandItem, Raise_CollapsingItem, Raise_ExpandingItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650149.aspx">TVN_ITEMEXPANDING</a>
	LRESULT OnItemExpandingNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled);
	/// \brief <em>\c TVN_KEYDOWN handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user pressed a key while the
	/// control had the keyboard focus.
	/// We use this handler to raise the \c IncrementalSearchStringChanging event.
	///
	/// \sa OnKeyDown, Raise_IncrementalSearchStringChanging,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650150.aspx">TVN_KEYDOWN</a>
	LRESULT OnKeyDownNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled);
	/// \brief <em>\c TVN_SELCHANGED handler</em>
	///
	/// Will be called if the control's parent window is notified, that the caret item has changed.
	/// We use this handler mainly to raise the \c CaretChanged event.
	///
	/// \sa OnCaretChangingNotification, OnSelectItem, Raise_CaretChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650151.aspx">TVN_SELCHANGED</a>
	LRESULT OnCaretChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVN_SELCHANGING handler</em>
	///
	/// Will be called if the control's parent window is notified, that the caret item is about to change.
	/// We use this handler mainly to raise the \c CaretChanging event.
	///
	/// \sa OnCaretChangedNotification, OnSelectItem, Raise_CaretChanging,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650152.aspx">TVN_SELCHANGING</a>
	LRESULT OnCaretChangingNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVN_SETDISPINFO handler</em>
	///
	/// Will be called if the control's parent window is notified, that it should update the information it
	/// maintains about the specified treeview item.
	/// We use this handler to raise the \c ItemSetText event.
	///
	/// \sa OnGetDispInfoNotification, Raise_ItemSetText,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650153.aspx">TVN_SETDISPINFO</a>
	LRESULT OnSetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TVN_SINGLEEXPAND handler</em>
	///
	/// Will be called if the control's parent window is notified, that a treeview item was single-expanded.
	/// We use this handler for the \c SingleExpand feature.
	///
	/// \sa get_SingleExpand, Raise_SingleExpanding,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650154.aspx">TVN_SINGLEEXPAND</a>
	LRESULT OnSingleExpandNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	/// \brief <em>\c EN_CHANGE handler</em>
	///
	/// Will be called if the control is notified, that the label-edit control's text was changed.
	/// We use this handler to raise the \c EditChange event.
	///
	/// \sa OnEditChar, Raise_ItemSetText,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672110.aspx">EN_CHANGE</a>
	LRESULT OnEditChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>\c NM_CUSTOMDRAW handler</em>
	///
	/// Will be called by the \c OnReflectedNotify method if a custom draw notification was received.
	/// We use this handler to raise the \c CustomDraw event.
	///
	/// \sa OnReflectedNotify, Raise_CustomDraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650133.aspx">NM_CUSTOMDRAW (tree view)</a>
	LRESULT OnCustomDrawNotification(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TTN_GETDISPINFOA handler</em>
	///
	/// Will be called by the \c OnNotify method if the \c TTN_GETDISPINFO notification was received.
	/// We use this handler to allow the client to cancel info tip popups.
	///
	/// \sa OnToolTipGetDispInfoNotificationW, OnNotify,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650474.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationA(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TTN_GETDISPINFOW handler</em>
	///
	/// Will be called by the \c OnNotify method if the \c TTN_GETDISPINFO notification was received.
	/// We use this handler to allow the client to cancel info tip popups.
	///
	/// \sa OnToolTipGetDispInfoNotificationA, OnNotify,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650474.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationW(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c AbortedDrag event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_AbortedDrag,
	///       ExTVwLibU::_IExplorerTreeViewEvents::AbortedDrag, Raise_Drop, EndDrag
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_AbortedDrag,
	///       ExTVwLibA::_IExplorerTreeViewEvents::AbortedDrag, Raise_Drop, EndDrag
	/// \endif
	inline HRESULT Raise_AbortedDrag(void);
	/// \brief <em>Raises the \c CaretChanged event</em>
	///
	/// \param[in] hPreviousCaretItem The previous caret item.
	/// \param[in] hNewCaretItem The new caret item.
	/// \param[in] caretChangeReason The reason for the caret change. Any \c TVC_* value is valid.
	/// \param[in] delay If \c TRUE, the event is delayed by the number of milliseconds defined by
	///            the \c CaretChangedDelayTime property; otherwise it is fired immediately.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_CaretChangedDelayTime, Proxy_IExplorerTreeViewEvents::Fire_CaretChanged,
	///       ExTVwLibU::_IExplorerTreeViewEvents::CaretChanged, Raise_CaretChanging,
	///       Raise_ItemSelectionChanged, Raise_SingleExpanding, get_CaretItem
	/// \else
	///   \sa put_CaretChangedDelayTime, Proxy_IExplorerTreeViewEvents::Fire_CaretChanged,
	///       ExTVwLibA::_IExplorerTreeViewEvents::CaretChanged, Raise_CaretChanging,
	///       Raise_ItemSelectionChanged, Raise_SingleExpanding, get_CaretItem
	/// \endif
	inline HRESULT Raise_CaretChanged(HTREEITEM hPreviousCaretItem, HTREEITEM hNewCaretItem, UINT caretChangeReason, BOOL delay = TRUE);
	/// \brief <em>Raises the \c CaretChanging event</em>
	///
	/// \param[in] hPreviousCaretItem The previous caret item.
	/// \param[in] hNewCaretItem The new caret item.
	/// \param[in] caretChangeReason The reason for the caret change. Any \c TVC_* value is valid.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort the caret change; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_CaretChanging,
	///       ExTVwLibU::_IExplorerTreeViewEvents::CaretChanging, Raise_CaretChanged,
	///       Raise_ItemSelectionChanging, Raise_SingleExpanding, get_CaretItem
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_CaretChanging,
	///       ExTVwLibA::_IExplorerTreeViewEvents::CaretChanging, Raise_CaretChanged,
	///       Raise_ItemSelectionChanging, Raise_SingleExpanding, get_CaretItem
	/// \endif
	inline HRESULT Raise_CaretChanging(HTREEITEM hPreviousCaretItem, HTREEITEM hNewCaretItem, UINT caretChangeReason, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c ChangedSortOrder event</em>
	///
	/// \param[in] previousSortOrder The control's old sort order. Any of the values defined by the
	///            \c SortOrderConstants enumeration is valid.
	/// \param[in] newSortOrder The control's new sort order. Any of the values defined by the
	///            \c SortOrderConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ChangedSortOrder,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ChangedSortOrder, Raise_ChangingSortOrder,
	///       ExTVwLibU::SortOrderConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ChangedSortOrder,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ChangedSortOrder, Raise_ChangingSortOrder,
	///       ExTVwLibA::SortOrderConstants
	/// \endif
	inline HRESULT Raise_ChangedSortOrder(SortOrderConstants previousSortOrder, SortOrderConstants newSortOrder);
	/// \brief <em>Raises the \c ChangingSortOrder event</em>
	///
	/// \param[in] previousSortOrder The control's old sort order. Any of the values defined by the
	///            \c SortOrderConstants enumeration is valid.
	/// \param[in] newSortOrder The control's new sort order. Any of the values defined by the
	///            \c SortOrderConstants enumeration is valid.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort redefining; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ChangingSortOrder,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ChangingSortOrder, Raise_ChangedSortOrder,
	///       ExTVwLibU::SortOrderConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ChangingSortOrder,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ChangingSortOrder, Raise_ChangedSortOrder,
	///       ExTVwLibA::SortOrderConstants
	/// \endif
	inline HRESULT Raise_ChangingSortOrder(SortOrderConstants previousSortOrder, SortOrderConstants newSortOrder, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c Click event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_Click,
	///       ExTVwLibU::_IExplorerTreeViewEvents::Click, Raise_DblClick, Raise_MClick, Raise_RClick,
	///       Raise_XClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_Click,
	///       ExTVwLibA::_IExplorerTreeViewEvents::Click, Raise_DblClick, Raise_MClick, Raise_RClick,
	///       Raise_XClick
	/// \endif
	inline HRESULT Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CollapsedItem event</em>
	///
	/// \param[in] pTreeItem The immediate parent item of the items that were collapsed.
	/// \param[in] removedAllSubItems If \c VARIANT_TRUE, the item's sub-items were removed; otherwise
	///            they simply were collapsed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TreeViewItem::Collapse, Proxy_IExplorerTreeViewEvents::Fire_CollapsedItem,
	///       ExTVwLibU::_IExplorerTreeViewEvents::CollapsedItem, Raise_CollapsingItem,
	///       Raise_ExpandedItem
	/// \else
	///   \sa TreeViewItem::Collapse, Proxy_IExplorerTreeViewEvents::Fire_CollapsedItem,
	///       ExTVwLibA::_IExplorerTreeViewEvents::CollapsedItem, Raise_CollapsingItem,
	///       Raise_ExpandedItem
	/// \endif
	inline HRESULT Raise_CollapsedItem(ITreeViewItem* pTreeItem, VARIANT_BOOL removedAllSubItems);
	/// \brief <em>Raises the \c CollapsingItem event</em>
	///
	/// \param[in] pTreeItem The immediate parent item of the items that get collapsed.
	/// \param[in] removingAllSubItems If \c VARIANT_TRUE, the item's sub-items get removed; otherwise
	///            they simply get collapsed.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort the collapse action;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TreeViewItem::Collapse, Proxy_IExplorerTreeViewEvents::Fire_CollapsingItem,
	///       ExTVwLibU::_IExplorerTreeViewEvents::CollapsingItem, Raise_CollapsedItem,
	///       Raise_ExpandingItem
	/// \else
	///   \sa TreeViewItem::Collapse, Proxy_IExplorerTreeViewEvents::Fire_CollapsingItem,
	///       ExTVwLibA::_IExplorerTreeViewEvents::CollapsingItem, Raise_CollapsedItem,
	///       Raise_ExpandingItem
	/// \endif
	inline HRESULT Raise_CollapsingItem(ITreeViewItem* pTreeItem, VARIANT_BOOL removingAllSubItems, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c CompareItems event</em>
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
	///   \sa get_SortOrder, SortItems, TreeViewItem::SortSubItems,
	///       Proxy_IExplorerTreeViewEvents::Fire_CompareItems,
	///       ExTVwLibU::_IExplorerTreeViewEvents::CompareItems, ExTVwLibU::CompareResultConstants
	/// \else
	///   \sa get_SortOrder, SortItems, TreeViewItem::SortSubItems,
	///       Proxy_IExplorerTreeViewEvents::Fire_CompareItems,
	///       ExTVwLibA::_IExplorerTreeViewEvents::CompareItems, ExTVwLibA::CompareResultConstants
	/// \endif
	inline HRESULT Raise_CompareItems(ITreeViewItem* pFirstItem, ITreeViewItem* pSecondItem, CompareResultConstants* pResult);
	/// \brief <em>Raises the \c ContextMenu event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ContextMenu,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ContextMenu, Raise_RClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ContextMenu,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ContextMenu, Raise_RClick
	/// \endif
	inline HRESULT Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CreatedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The label-edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_CreatedEditControlWindow,
	///       ExTVwLibU::_IExplorerTreeViewEvents::CreatedEditControlWindow,
	///       Raise_DestroyedEditControlWindow, Raise_StartingLabelEditing, get_hWndEdit
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_CreatedEditControlWindow,
	///       ExTVwLibA::_IExplorerTreeViewEvents::CreatedEditControlWindow,
	///       Raise_DestroyedEditControlWindow, Raise_StartingLabelEditing, get_hWndEdit
	/// \endif
	inline HRESULT Raise_CreatedEditControlWindow(LONG hWndEdit);
	/// \brief <em>Raises the \c CustomDraw event</em>
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
	/// \param[in,out] pFurtherProcessing A value controlling further drawing. Most of the values
	///                defined by the \c CustomDrawReturnValuesConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_CustomDraw,
	///       ExTVwLibU::_IExplorerTreeViewEvents::CustomDraw, TreeViewItem::get_Level, get_ForeColor,
	///       get_BlendSelectedItemsIcons, get_FavoritesStyle, ExTVwLibU::RECTANGLE,
	///       ExTVwLibU::CustomDrawStageConstants, ExTVwLibU::CustomDrawItemStateConstants,
	///       ExTVwLibU::CustomDrawReturnValuesConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms650133.aspx">NM_CUSTOMDRAW (tree view)</a>
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_CustomDraw,
	///       ExTVwLibA::_IExplorerTreeViewEvents::CustomDraw, TreeViewItem::get_Level, get_ForeColor,
	///       get_BlendSelectedItemsIcons, get_FavoritesStyle, ExTVwLibA::RECTANGLE,
	///       ExTVwLibA::CustomDrawStageConstants, ExTVwLibA::CustomDrawItemStateConstants,
	///       ExTVwLibA::CustomDrawReturnValuesConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms650133.aspx">NM_CUSTOMDRAW (tree view)</a>
	/// \endif
	inline HRESULT Raise_CustomDraw(ITreeViewItem* pTreeItem, LONG* pItemLevel, OLE_COLOR* pTextColor, OLE_COLOR* pTextBackColor, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants itemState, LONG hDC, RECTANGLE* pDrawingRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing);
	/// \brief <em>Raises the \c DblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_DblClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::DblClick, Raise_Click, Raise_MDblClick,
	///       Raise_RDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_DblClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::DblClick, Raise_Click, Raise_MDblClick,
	///       Raise_RDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_DestroyedControlWindow,
	///       ExTVwLibU::_IExplorerTreeViewEvents::DestroyedControlWindow,
	///       Raise_RecreatedControlWindow, get_hWnd
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_DestroyedControlWindow,
	///       ExTVwLibA::_IExplorerTreeViewEvents::DestroyedControlWindow,
	///       Raise_RecreatedControlWindow, get_hWnd
	/// \endif
	inline HRESULT Raise_DestroyedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c DestroyedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The label-edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_DestroyedEditControlWindow,
	///       ExTVwLibU::_IExplorerTreeViewEvents::DestroyedEditControlWindow,
	///       Raise_CreatedEditControlWindow, Raise_StartingLabelEditing, get_hWndEdit, EndLabelEdit
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_DestroyedEditControlWindow,
	///       ExTVwLibA::_IExplorerTreeViewEvents::DestroyedEditControlWindow,
	///       Raise_CreatedEditControlWindow, Raise_StartingLabelEditing, get_hWndEdit, EndLabelEdit
	/// \endif
	inline HRESULT Raise_DestroyedEditControlWindow(LONG hWndEdit);
	/// \brief <em>Raises the \c DragMouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_DragMouseMove,
	///       ExTVwLibU::_IExplorerTreeViewEvents::DragMouseMove, Raise_MouseMove,
	///       Raise_OLEDragMouseMove, BeginDrag
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_DragMouseMove,
	///       ExTVwLibA::_IExplorerTreeViewEvents::DragMouseMove, Raise_MouseMove,
	///       Raise_OLEDragMouseMove, BeginDrag
	/// \endif
	inline HRESULT Raise_DragMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c Drop event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_Drop, ExTVwLibU::_IExplorerTreeViewEvents::Drop,
	///       Raise_AbortedDrag, EndDrag
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_Drop, ExTVwLibA::_IExplorerTreeViewEvents::Drop,
	///       Raise_AbortedDrag, EndDrag
	/// \endif
	inline HRESULT Raise_Drop(void);
	/// \brief <em>Raises the \c EditChange event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditChange,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditChange, Raise_EditKeyPress, get_EditText
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditChange,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditChange, Raise_EditKeyPress, get_EditText
	/// \endif
	inline HRESULT Raise_EditChange(void);
	/// \brief <em>Raises the \c EditClick event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditClick, Raise_EditDblClick, Raise_EditMClick,
	///       Raise_EditRClick, Raise_EditXClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditClick, Raise_EditDblClick, Raise_EditMClick,
	///       Raise_EditRClick, Raise_EditXClick
	/// \endif
	inline HRESULT Raise_EditClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c EditContextMenu event</em>
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
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditContextMenu,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditContextMenu, Raise_EditRClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditContextMenu,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditContextMenu, Raise_EditRClick
	/// \endif
	inline HRESULT Raise_EditContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu);
	/// \brief <em>Raises the \c EditDblClick event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditDblClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditDblClick, Raise_EditClick, Raise_EditMDblClick,
	///       Raise_EditRDblClick, Raise_EditXDblClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditDblClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditDblClick, Raise_EditClick, Raise_EditMDblClick,
	///       Raise_EditRDblClick, Raise_EditXDblClick
	/// \endif
	inline HRESULT Raise_EditDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c EditGotFocus event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditGotFocus,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditGotFocus, Raise_EditLostFocus
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditGotFocus,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditGotFocus, Raise_EditLostFocus
	/// \endif
	inline HRESULT Raise_EditGotFocus(void);
	/// \brief <em>Raises the \c EditKeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN
	///                message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditKeyDown,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditKeyDown, Raise_EditKeyUp, Raise_EditKeyPress
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditKeyDown,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditKeyDown, Raise_EditKeyUp, Raise_EditKeyPress
	/// \endif
	inline HRESULT Raise_EditKeyDown(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c EditKeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditKeyPress,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditKeyPress, Raise_EditKeyDown, Raise_EditKeyUp
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditKeyPress,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditKeyPress, Raise_EditKeyDown, Raise_EditKeyUp
	/// \endif
	inline HRESULT Raise_EditKeyPress(SHORT* pKeyAscii);
	/// \brief <em>Raises the \c EditKeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditKeyUp,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditKeyUp, Raise_EditKeyDown, Raise_EditKeyPress
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditKeyUp,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditKeyUp, Raise_EditKeyDown, Raise_EditKeyPress
	/// \endif
	inline HRESULT Raise_EditKeyUp(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c EditLostFocus event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditLostFocus,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditLostFocus, Raise_EditGotFocus
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditLostFocus,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditLostFocus, Raise_EditGotFocus
	/// \endif
	inline HRESULT Raise_EditLostFocus(void);
	/// \brief <em>Raises the \c EditMClick event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditMClick, Raise_EditMDblClick, Raise_EditClick,
	///       Raise_EditRClick, Raise_EditXClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditMClick, Raise_EditMDblClick, Raise_EditClick,
	///       Raise_EditRClick, Raise_EditXClick
	/// \endif
	inline HRESULT Raise_EditMClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c EditMDblClick event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMDblClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditMDblClick, Raise_EditMClick, Raise_EditDblClick,
	///       Raise_EditRDblClick, Raise_EditXDblClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMDblClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditMDblClick, Raise_EditMClick, Raise_EditDblClick,
	///       Raise_EditRDblClick, Raise_EditXDblClick
	/// \endif
	inline HRESULT Raise_EditMDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c EditMouseDown event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseDown,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditMouseDown, Raise_EditMouseUp, Raise_EditClick,
	///       Raise_EditMClick, Raise_EditRClick, Raise_EditXClick, ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseDown,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditMouseDown, Raise_EditMouseUp, Raise_EditClick,
	///       Raise_EditMClick, Raise_EditRClick, Raise_EditXClick, ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_EditMouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c EditMouseEnter event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseEnter,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditMouseEnter, Raise_EditMouseLeave,
	///       Raise_EditMouseHover, Raise_EditMouseMove, ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseEnter,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditMouseEnter, Raise_EditMouseLeave,
	///       Raise_EditMouseHover, Raise_EditMouseMove, ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_EditMouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c EditMouseHover event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseHover,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditMouseHover, Raise_EditMouseEnter,
	///       Raise_EditMouseLeave, Raise_EditMouseMove, put_EditHoverTime,
	///       ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseHover,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditMouseHover, Raise_EditMouseEnter,
	///       Raise_EditMouseLeave, Raise_EditMouseMove, put_EditHoverTime,
	///       ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_EditMouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c EditMouseLeave event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseLeave,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditMouseLeave, Raise_EditMouseEnter,
	///       Raise_EditMouseHover, Raise_EditMouseMove, ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseLeave,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditMouseLeave, Raise_EditMouseEnter,
	///       Raise_EditMouseHover, Raise_EditMouseMove, ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_EditMouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c EditMouseMove event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseMove,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditMouseMove, Raise_EditMouseEnter,
	///       Raise_EditMouseLeave, Raise_EditMouseDown, Raise_EditMouseUp, Raise_EditMouseWheel,
	///       ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseMove,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditMouseMove, Raise_EditMouseEnter,
	///       Raise_EditMouseLeave, Raise_EditMouseDown, Raise_EditMouseUp, Raise_EditMouseWheel,
	///       ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_EditMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c EditMouseUp event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseUp,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditMouseUp, Raise_EditMouseDown, Raise_EditClick,
	///       Raise_EditMClick, Raise_EditRClick, Raise_EditXClick, ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseUp,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditMouseUp, Raise_EditMouseDown, Raise_EditClick,
	///       Raise_EditMClick, Raise_EditRClick, Raise_EditXClick, ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_EditMouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c EditMouseWheel event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseWheel,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditMouseWheel, Raise_EditMouseMove, Raise_MouseWheel,
	///       ExTVwLibU::ExtendedMouseButtonConstants, ExTVwLibU::ScrollAxisConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditMouseWheel,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditMouseWheel, Raise_EditMouseMove, Raise_MouseWheel,
	///       ExTVwLibA::ExtendedMouseButtonConstants, ExTVwLibA::ScrollAxisConstants
	/// \endif
	inline HRESULT Raise_EditMouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta);
	/// \brief <em>Raises the \c EditRClick event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditRClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditRClick, Raise_EditContextMenu,
	///       Raise_EditRDblClick, Raise_EditClick, Raise_EditMClick, Raise_EditXClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::Fire_RClick, Raise_EditContextMenu,
	///       Raise_EditRDblClick, Raise_EditClick, Raise_EditMClick, Raise_EditXClick
	/// \endif
	inline HRESULT Raise_EditRClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c EditRDblClick event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditRDblClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditRDblClick, Raise_EditRClick, Raise_EditDblClick,
	///       Raise_EditMDblClick, Raise_EditXDblClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditRDblClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditRDblClick, Raise_EditRClick, Raise_EditDblClick,
	///       Raise_EditMDblClick, Raise_EditXDblClick
	/// \endif
	inline HRESULT Raise_EditRDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c EditXClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the contained edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the contained edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditXClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditXClick, Raise_EditXDblClick, Raise_EditClick,
	///       Raise_EditMClick, Raise_EditRClick, Raise_EditXClick, ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditXClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditXClick, Raise_EditXDblClick, Raise_EditClick,
	///       Raise_EditMClick, Raise_EditRClick, Raise_EditXClick, ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_EditXClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c EditXDblClick event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditXDblClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::EditXDblClick, Raise_EditXClick, Raise_EditDblClick,
	///       Raise_EditMDblClick, Raise_EditRDblClick, Raise_EditXDblClick,
	///       ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_EditXDblClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::EditXDblClick, Raise_EditXClick, Raise_EditDblClick,
	///       Raise_EditMDblClick, Raise_EditRDblClick, Raise_EditXDblClick,
	///       ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_EditXDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c ExpandedItem event</em>
	///
	/// \param[in] pTreeItem The expanded item.
	/// \param[in] expandedPartially If \c VARIANT_TRUE, the "+" button next to the item was replaced
	///            with a "-" button; otherwise the button remained a "+".
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TreeViewItem::Expand, Proxy_IExplorerTreeViewEvents::Fire_ExpandedItem,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ExpandedItem, Raise_ExpandingItem,
	///       Raise_CollapsedItem
	/// \else
	///   \sa TreeViewItem::Expand, Proxy_IExplorerTreeViewEvents::Fire_ExpandedItem,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ExpandedItem, Raise_ExpandingItem,
	///       Raise_CollapsedItem
	/// \endif
	inline HRESULT Raise_ExpandedItem(ITreeViewItem* pTreeItem, VARIANT_BOOL expandedPartially);
	/// \brief <em>Raises the \c ExpandingItem event</em>
	///
	/// \param[in] pTreeItem The expanding item.
	/// \param[in] expandingPartially If \c VARIANT_TRUE, the "+" button next to the item is replaced
	///            with a "-" button; otherwise the button remains a "+".
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort the expand action; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TreeViewItem::Expand, Proxy_IExplorerTreeViewEvents::Fire_ExpandingItem,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ExpandingItem, Raise_ExpandedItem,
	///       Raise_CollapsingItem
	/// \else
	///   \sa TreeViewItem::Expand, Proxy_IExplorerTreeViewEvents::Fire_ExpandingItem,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ExpandingItem, Raise_ExpandedItem,
	///       Raise_CollapsingItem
	/// \endif
	inline HRESULT Raise_ExpandingItem(ITreeViewItem* pTreeItem, VARIANT_BOOL expandingPartially, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c FreeItemData event</em>
	///
	/// \param[in] pTreeItem The item whose associated data shall be freed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_FreeItemData,
	///       ExTVwLibU::_IExplorerTreeViewEvents::FreeItemData, Raise_RemovingItem, Raise_RemovedItem,
	///       TreeViewItem::put_ItemData
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_FreeItemData,
	///       ExTVwLibA::_IExplorerTreeViewEvents::FreeItemData, Raise_RemovingItem, Raise_RemovedItem,
	///       TreeViewItem::put_ItemData
	/// \endif
	inline HRESULT Raise_FreeItemData(ITreeViewItem* pTreeItem);
	/// \brief <em>Raises the \c IncrementalSearchStringChanging event</em>
	///
	/// \param[in] currentSearchString The control's current incremental search-string.
	/// \param[in] keyCodeOfCharToBeAdded The key code of the character to be added to the search-string.
	///            Most of the values defined by VB's \c KeyCodeConstants enumeration are valid.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should discard the character and leave the
	///                search-string unchanged; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_IncrementalSearchStringChanging,
	///       ExTVwLibU::_IExplorerTreeViewEvents::IncrementalSearchStringChanging, Raise_KeyPress,
	///       get_IncrementalSearchString
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_IncrementalSearchStringChanging,
	///       ExTVwLibA::_IExplorerTreeViewEvents::IncrementalSearchStringChanging, Raise_KeyPress,
	///       get_IncrementalSearchString
	/// \endif
	inline HRESULT Raise_IncrementalSearchStringChanging(BSTR currentSearchString, SHORT keyCodeOfCharToBeAdded, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c InsertedItem event</em>
	///
	/// \param[in] pTreeItem The inserted item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_InsertedItem,
	///       ExTVwLibU::_IExplorerTreeViewEvents::InsertedItem, Raise_InsertingItem, Raise_RemovedItem
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_InsertedItem,
	///       ExTVwLibA::_IExplorerTreeViewEvents::InsertedItem, Raise_InsertingItem, Raise_RemovedItem
	/// \endif
	inline HRESULT Raise_InsertedItem(ITreeViewItem* pTreeItem);
	/// \brief <em>Raises the \c InsertingItem event</em>
	///
	/// \param[in] pTreeItem The item being inserted.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort insertion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_InsertingItem,
	///       ExTVwLibU::_IExplorerTreeViewEvents::InsertingItem, Raise_InsertedItem, Raise_RemovingItem
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_InsertingItem,
	///       ExTVwLibA::_IExplorerTreeViewEvents::InsertingItem, Raise_InsertedItem, Raise_RemovingItem
	/// \endif
	inline HRESULT Raise_InsertingItem(IVirtualTreeViewItem* pTreeItem, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c ItemAsynchronousDrawFailed event</em>
	///
	/// \param[in] pTreeItem The item whose image failed to be drawn.
	/// \param[in,out] pNotificationDetails A \c NMTVASYNCDRAW struct holding the notification details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemAsynchronousDrawFailed,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ItemAsynchronousDrawFailed, get_DrawImagesAsynchronously,
	///       ExTVwLibU::FAILEDIMAGEDETAILS,
	///       <a href="https://msdn.microsoft.com/en-us/library/aa361676.aspx">NMTVASYNCDRAW</a>
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemAsynchronousDrawFailed,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ItemAsynchronousDrawFailed, get_DrawImagesAsynchronously,
	///       ExTVwLibA::FAILEDIMAGEDETAILS,
	///       <a href="https://msdn.microsoft.com/en-us/library/aa361676.aspx">NMTVASYNCDRAW</a>
	/// \endif
	inline HRESULT Raise_ItemAsynchronousDrawFailed(ITreeViewItem* pTreeItem, NMTVASYNCDRAW* pNotificationDetails);
	/// \brief <em>Raises the \c ItemBeginDrag event</em>
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
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemBeginDrag,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ItemBeginDrag, BeginDrag, OLEDrag, get_AllowDragDrop,
	///       Raise_ItemBeginRDrag, ExTVwLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemBeginDrag,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ItemBeginDrag, BeginDrag, OLEDrag, get_AllowDragDrop,
	///       Raise_ItemBeginRDrag, ExTVwLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ItemBeginDrag(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c ItemBeginRDrag event</em>
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
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemBeginRDrag,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ItemBeginRDrag, BeginDrag, OLEDrag, get_AllowDragDrop,
	///       Raise_ItemBeginDrag, ExTVwLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemBeginRDrag,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ItemBeginRDrag, BeginDrag, OLEDrag, get_AllowDragDrop,
	///       Raise_ItemBeginDrag, ExTVwLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ItemBeginRDrag(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c ItemGetDisplayInfo event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemGetDisplayInfo,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ItemGetDisplayInfo, Raise_ItemSetText,
	///       TreeViewItem::put_IconIndex, TreeViewItem::put_SelectedIconIndex,
	///       TreeViewItem::put_ExpandedIconIndex, put_hImageList, TreeViewItem::put_HasExpando,
	///       put_TreeViewStyle, TreeViewItem::put_Text, ExTVwLibU::RequestedInfoConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemGetDisplayInfo,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ItemGetDisplayInfo, Raise_ItemSetText,
	///       TreeViewItem::put_IconIndex, TreeViewItem::put_SelectedIconIndex,
	///       TreeViewItem::put_ExpandedIconIndex, put_hImageList, TreeViewItem::put_HasExpando,
	///       put_TreeViewStyle, TreeViewItem::put_Text, ExTVwLibA::RequestedInfoConstants
	/// \endif
	inline HRESULT Raise_ItemGetDisplayInfo(ITreeViewItem* pTreeItem, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG* pSelectedIconIndex, LONG* pExpandedIconIndex, VARIANT_BOOL* pHasButton, LONG maxItemTextLength, BSTR* pItemText, VARIANT_BOOL* pDontAskAgain);
	/// \brief <em>Raises the \c ItemGetInfoTipText event</em>
	///
	/// \param[in] pTreeItem The item whose info tip text is requested.
	/// \param[in] maxInfoTipLength The maximum number of characters the info tip text may consist of.
	/// \param[out] pInfoTipText The item's info tip text.
	/// \param[in,out] pAbortToolTip If \c VARIANT_TRUE, the caller should NOT show the tooltip;
	///                otherwise it should.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemGetInfoTipText,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ItemGetInfoTipText, put_ShowToolTips,
	///       TreeViewItem::DisplayInfoTip
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemGetInfoTipText,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ItemGetInfoTipText, put_ShowToolTips,
	///       TreeViewItem::DisplayInfoTip
	/// \endif
	inline HRESULT Raise_ItemGetInfoTipText(ITreeViewItem* pTreeItem, LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip);
	/// \brief <em>Raises the \c ItemMouseEnter event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemMouseEnter,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ItemMouseEnter, Raise_ItemMouseLeave, Raise_MouseMove,
	///       ExTVwLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemMouseEnter,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ItemMouseEnter, Raise_ItemMouseLeave, Raise_MouseMove,
	///       ExTVwLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ItemMouseEnter(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c ItemMouseLeave event</em>
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
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemMouseLeave,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ItemMouseLeave, Raise_ItemMouseEnter, Raise_MouseMove,
	///       ExTVwLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemMouseLeave,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ItemMouseLeave, Raise_ItemMouseEnter, Raise_MouseMove,
	///       ExTVwLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ItemMouseLeave(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c ItemSelectionChanged event</em>
	///
	/// \param[in] hItem The item that was selected/deselected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemSelectionChanged,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ItemSelectionChanged, Raise_ItemSelectionChanging,
	///       Raise_CaretChanged
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemSelectionChanged,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ItemSelectionChanged, Raise_ItemSelectionChanging,
	///       Raise_CaretChanged
	/// \endif
	inline HRESULT Raise_ItemSelectionChanged(HTREEITEM hItem);
	/// \brief <em>Raises the \c ItemSelectionChanging event</em>
	///
	/// \param[in] hItem The item that will be selected/deselected.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should prevent the selection state from
	///                changing; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemSelectionChanging,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ItemSelectionChanging, Raise_ItemSelectionChanged,
	///       Raise_CaretChanging
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemSelectionChanging,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ItemSelectionChanging, Raise_ItemSelectionChanged,
	///       Raise_CaretChanging
	/// \endif
	inline HRESULT Raise_ItemSelectionChanging(HTREEITEM hItem, VARIANT_BOOL* pCancelChange);
	/// \brief <em>Raises the \c ItemSetText event</em>
	///
	/// \param[in] pTreeItem The item whose text was changed.
	/// \param[in] itemText The item's new text.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemSetText,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ItemSetText, Raise_ItemGetDisplayInfo,
	///       TreeViewItem::put_Text
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemSetText,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ItemSetText, Raise_ItemGetDisplayInfo,
	///       TreeViewItem::put_Text
	/// \endif
	inline HRESULT Raise_ItemSetText(ITreeViewItem* pTreeItem, BSTR itemText);
	/// \brief <em>Raises the \c ItemStateImageChanged event</em>
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
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemStateImageChanged,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ItemStateImageChanged, Raise_ItemStateImageChanging,
	///       TreeViewItem::get_StateImageIndex, ExTVwLibU::StateImageChangeCausedByConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemStateImageChanged,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ItemStateImageChanged, Raise_ItemStateImageChanging,
	///       TreeViewItem::get_StateImageIndex, ExTVwLibA::StateImageChangeCausedByConstants
	/// \endif
	inline HRESULT Raise_ItemStateImageChanged(ITreeViewItem* pTreeItem, LONG previousStateImageIndex, LONG newStateImageIndex, StateImageChangeCausedByConstants causedBy);
	/// \brief <em>Raises the \c ItemStateImageChanging event</em>
	///
	/// \param[in] pTreeItem The item whose state image shall be changed.
	/// \param[in] previousStateImageIndex The one-based index of the item's previous state image.
	/// \param[in,out] pNewStateImageIndex The one-based index of the item's new state image. The client
	///                may change this value.
	/// \param[in] causedBy The reason for the state image change. Any of the values defined by the
	///            \c StateImageChangeCausedByConstants enumeration is valid.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should prevent the state image from
	///                changing; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The user may change the new state image's index.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemStateImageChanging,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ItemStateImageChanging, Raise_ItemStateImageChanged,
	///       TreeViewItem::get_StateImageIndex, ExTVwLibU::StateImageChangeCausedByConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ItemStateImageChanging,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ItemStateImageChanging, Raise_ItemStateImageChanged,
	///       TreeViewItem::get_StateImageIndex, ExTVwLibA::StateImageChangeCausedByConstants
	/// \endif
	inline HRESULT Raise_ItemStateImageChanging(ITreeViewItem* pTreeItem, LONG previousStateImageIndex, LONG* pNewStateImageIndex, StateImageChangeCausedByConstants causedBy, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_KeyDown,
	///       ExTVwLibU::_IExplorerTreeViewEvents::KeyDown, Raise_IncrementalSearchStringChanging,
	///       Raise_KeyUp, Raise_KeyPress
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_KeyDown,
	///       ExTVwLibA::_IExplorerTreeViewEvents::KeyDown, Raise_IncrementalSearchStringChanging,
	///       Raise_KeyUp, Raise_KeyPress
	/// \endif
	inline HRESULT Raise_KeyDown(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_KeyPress,
	///       ExTVwLibU::_IExplorerTreeViewEvents::KeyPress, Raise_KeyDown, Raise_KeyUp
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_KeyPress,
	///       ExTVwLibA::_IExplorerTreeViewEvents::KeyPress, Raise_KeyDown, Raise_KeyUp
	/// \endif
	inline HRESULT Raise_KeyPress(SHORT* pKeyAscii);
	/// \brief <em>Raises the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_KeyUp,
	///       ExTVwLibU::_IExplorerTreeViewEvents::KeyUp, Raise_KeyDown, Raise_KeyPress
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_KeyUp,
	///       ExTVwLibA::_IExplorerTreeViewEvents::KeyUp, Raise_KeyDown, Raise_KeyPress
	/// \endif
	inline HRESULT Raise_KeyUp(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c MClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::MClick, Raise_MDblClick, Raise_Click, Raise_RClick,
	///       Raise_XClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::MClick, Raise_MDblClick, Raise_Click, Raise_RClick,
	///       Raise_XClick
	/// \endif
	inline HRESULT Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MDblClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::MDblClick, Raise_MClick, Raise_DblClick,
	///       Raise_RDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MDblClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::MDblClick, Raise_MClick, Raise_DblClick,
	///       Raise_RDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseDown event</em>
	///
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseDown,
	///       ExTVwLibU::_IExplorerTreeViewEvents::MouseDown, Raise_MouseUp, Raise_Click, Raise_MClick,
	///       Raise_RClick, Raise_XClick, ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseDown,
	///       ExTVwLibA::_IExplorerTreeViewEvents::MouseDown, Raise_MouseUp, Raise_Click, Raise_MClick,
	///       Raise_RClick, Raise_XClick, ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseEnter event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseEnter,
	///       ExTVwLibU::_IExplorerTreeViewEvents::MouseEnter, Raise_MouseLeave, Raise_ItemMouseEnter,
	///       Raise_MouseHover, Raise_MouseMove, ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseEnter,
	///       ExTVwLibA::_IExplorerTreeViewEvents::MouseEnter, Raise_MouseLeave, Raise_ItemMouseEnter,
	///       Raise_MouseHover, Raise_MouseMove, ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseHover event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseHover,
	///       ExTVwLibU::_IExplorerTreeViewEvents::MouseHover, Raise_MouseEnter, Raise_MouseLeave,
	///       Raise_MouseMove, put_HoverTime, ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseHover,
	///       ExTVwLibA::_IExplorerTreeViewEvents::MouseHover, Raise_MouseEnter, Raise_MouseLeave,
	///       Raise_MouseMove, put_HoverTime, ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseLeave event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseLeave,
	///       ExTVwLibU::_IExplorerTreeViewEvents::MouseLeave, Raise_MouseEnter, Raise_ItemMouseLeave,
	///       Raise_MouseHover, Raise_MouseMove, ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseLeave,
	///       ExTVwLibA::_IExplorerTreeViewEvents::MouseLeave, Raise_MouseEnter, Raise_ItemMouseLeave,
	///       Raise_MouseHover, Raise_MouseMove, ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseMove,
	///       ExTVwLibU::_IExplorerTreeViewEvents::MouseMove, Raise_MouseEnter, Raise_MouseLeave,
	///       Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel, ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseMove,
	///       ExTVwLibA::_IExplorerTreeViewEvents::MouseMove, Raise_MouseEnter, Raise_MouseLeave,
	///       Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel, ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseUp event</em>
	///
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseUp,
	///       ExTVwLibU::_IExplorerTreeViewEvents::MouseUp, Raise_MouseDown, Raise_Click, Raise_MClick,
	///       Raise_RClick, Raise_XClick, ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseUp,
	///       ExTVwLibA::_IExplorerTreeViewEvents::MouseUp, Raise_MouseDown, Raise_Click, Raise_MClick,
	///       Raise_RClick, Raise_XClick, ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseWheel event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseWheel,
	///       ExTVwLibU::_IExplorerTreeViewEvents::MouseWheel, Raise_MouseMove, Raise_EditMouseWheel,
	///       ExTVwLibU::ExtendedMouseButtonConstants, ExTVwLibU::ScrollAxisConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_MouseWheel,
	///       ExTVwLibA::_IExplorerTreeViewEvents::MouseWheel, Raise_MouseMove, Raise_EditMouseWheel,
	///       ExTVwLibA::ExtendedMouseButtonConstants, ExTVwLibA::ScrollAxisConstants
	/// \endif
	inline HRESULT Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta);
	/// \brief <em>Raises the \c OLECompleteDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] performedEffect The performed drop effect. Any of the values (except \c odeScroll)
	///            defined by the \c OLEDropEffectConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLECompleteDrag,
	///       ExTVwLibU::_IExplorerTreeViewEvents::OLECompleteDrag, Raise_OLEStartDrag,
	///       ExTVwLibU::IOLEDataObject::GetData, OLEDrag
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLECompleteDrag,
	///       ExTVwLibA::_IExplorerTreeViewEvents::OLECompleteDrag, Raise_OLEStartDrag,
	///       ExTVwLibA::IOLEDataObject::GetData, OLEDrag
	/// \endif
	inline HRESULT Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect);
	/// \brief <em>Raises the \c OLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client finally executed.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[out] pCallDropTargetHelper If set to \c TRUE, the caller should call the appropriate
	///             \c IDropTargetHelper method; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEDragDrop,
	///       ExTVwLibU::_IExplorerTreeViewEvents::OLEDragDrop, Raise_OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEDragDrop,
	///       ExTVwLibA::_IExplorerTreeViewEvents::OLEDragDrop, Raise_OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragDrop(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper);
	/// \brief <em>Raises the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[out] pCallDropTargetHelper If set to \c TRUE, the caller should call the appropriate
	///             \c IDropTargetHelper method; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEDragEnter,
	///       ExTVwLibU::_IExplorerTreeViewEvents::OLEDragEnter, Raise_OLEDragMouseMove,
	///       Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEDragEnter,
	///       ExTVwLibA::_IExplorerTreeViewEvents::OLEDragEnter, Raise_OLEDragMouseMove,
	///       Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragEnter(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper);
	/// \brief <em>Raises the \c OLEDragEnterPotentialTarget event</em>
	///
	/// \param[in] hWndPotentialTarget The potential drop target window's handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEDragEnterPotentialTarget,
	///       ExTVwLibU::_IExplorerTreeViewEvents::OLEDragEnterPotentialTarget,
	///       Raise_OLEDragLeavePotentialTarget, OLEDrag
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEDragEnterPotentialTarget,
	///       ExTVwLibA::_IExplorerTreeViewEvents::OLEDragEnterPotentialTarget,
	///       Raise_OLEDragLeavePotentialTarget, OLEDrag
	/// \endif
	inline HRESULT Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget);
	/// \brief <em>Raises the \c OLEDragLeave event</em>
	///
	/// \param[out] pCallDropTargetHelper If set to \c TRUE, the caller should call the appropriate
	///             \c IDropTargetHelper method; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEDragLeave,
	///       ExTVwLibU::_IExplorerTreeViewEvents::OLEDragLeave, Raise_OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEDragLeave,
	///       ExTVwLibA::_IExplorerTreeViewEvents::OLEDragLeave, Raise_OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \endif
	inline HRESULT Raise_OLEDragLeave(BOOL* pCallDropTargetHelper);
	/// \brief <em>Raises the \c OLEDragLeavePotentialTarget event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEDragLeavePotentialTarget,
	///       ExTVwLibU::_IExplorerTreeViewEvents::OLEDragLeavePotentialTarget,
	///       Raise_OLEDragEnterPotentialTarget, OLEDrag
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEDragLeavePotentialTarget,
	///       ExTVwLibA::_IExplorerTreeViewEvents::OLEDragLeavePotentialTarget,
	///       Raise_OLEDragEnterPotentialTarget, OLEDrag
	/// \endif
	inline HRESULT Raise_OLEDragLeavePotentialTarget(void);
	/// \brief <em>Raises the \c OLEDragMouseMove event</em>
	///
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[out] pCallDropTargetHelper If set to \c TRUE, the caller should call the appropriate
	///             \c IDropTargetHelper method; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEDragMouseMove,
	///       ExTVwLibU::_IExplorerTreeViewEvents::OLEDragMouseMove, Raise_OLEDragEnter,
	///       Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEDragMouseMove,
	///       ExTVwLibA::_IExplorerTreeViewEvents::OLEDragMouseMove, Raise_OLEDragEnter,
	///       Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper);
	/// \brief <em>Raises the \c OLEGiveFeedback event</em>
	///
	/// \param[in] effect The current drop effect. It is chosen by the potential drop target. Any of
	///            the values defined by the \c DROPEFFECT enumeration is valid.
	/// \param[in,out] pUseDefaultCursors If set to \c VARIANT_TRUE, the system's default mouse cursors
	///                shall be used to visualize the various drop effects. If set to \c VARIANT_FALSE,
	///                the client has set a custom mouse cursor.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEGiveFeedback,
	///       ExTVwLibU::_IExplorerTreeViewEvents::OLEGiveFeedback, Raise_OLEQueryContinueDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEGiveFeedback,
	///       ExTVwLibA::_IExplorerTreeViewEvents::OLEGiveFeedback, Raise_OLEQueryContinueDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors);
	/// \brief <em>Raises the \c OLEQueryContinueDrag event</em>
	///
	/// \param[in] pressedEscape If \c TRUE, the user has pressed the \c ESC key since the last time
	///            this event was raised; otherwise not.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in,out] pActionToContinueWith Indicates whether to continue (\c S_OK), cancel
	///                (\c DRAGDROP_S_CANCEL) or complete (\c DRAGDROP_S_DROP) the drag'n'drop operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEQueryContinueDrag,
	///       ExTVwLibU::_IExplorerTreeViewEvents::OLEQueryContinueDrag, Raise_OLEGiveFeedback
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEQueryContinueDrag,
	///       ExTVwLibA::_IExplorerTreeViewEvents::OLEQueryContinueDrag, Raise_OLEGiveFeedback
	/// \endif
	inline HRESULT Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith);
	/// \brief <em>Raises the \c OLEReceivedNewData event</em>
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
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEReceivedNewData,
	///       ExTVwLibU::_IExplorerTreeViewEvents::OLEReceivedNewData, Raise_OLESetData,
	///       SourceOLEDataObject::GetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEReceivedNewData,
	///       ExTVwLibA::_IExplorerTreeViewEvents::OLEReceivedNewData, Raise_OLESetData,
	///       SourceOLEDataObject::GetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \endif
	HRESULT Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect);
	/// \brief <em>Raises the \c OLESetData event</em>
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
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLESetData,
	///       ExTVwLibU::_IExplorerTreeViewEvents::OLESetData, Raise_OLEStartDrag,
	///       SourceOLEDataObject::SetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLESetData,
	///       ExTVwLibA::_IExplorerTreeViewEvents::OLESetData, Raise_OLEStartDrag,
	///       SourceOLEDataObject::SetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \endif
	HRESULT Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect);
	/// \brief <em>Raises the \c OLEStartDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEStartDrag,
	///       ExTVwLibU::_IExplorerTreeViewEvents::OLEStartDrag, Raise_OLESetData, Raise_OLECompleteDrag,
	///       SourceOLEDataObject::SetData, OLEDrag
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_OLEStartDrag,
	///       ExTVwLibA::_IExplorerTreeViewEvents::OLEStartDrag, Raise_OLESetData, Raise_OLECompleteDrag,
	///       SourceOLEDataObject::SetData, OLEDrag
	/// \endif
	inline HRESULT Raise_OLEStartDrag(IOLEDataObject* pData);
	/// \brief <em>Raises the \c RClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::RClick, Raise_ContextMenu, Raise_RDblClick,
	///       Raise_Click, Raise_MClick, Raise_XClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::RClick, Raise_ContextMenu, Raise_RDblClick,
	///       Raise_Click, Raise_MClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RDblClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::RDblClick, Raise_RClick, Raise_DblClick,
	///       Raise_MDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RDblClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::RDblClick, Raise_RClick, Raise_DblClick,
	///       Raise_MDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RecreatedControlWindow,
	///       ExTVwLibU::_IExplorerTreeViewEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RecreatedControlWindow,
	///       ExTVwLibA::_IExplorerTreeViewEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_RecreatedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c RemovedItem event</em>
	///
	/// \param[in] pTreeItem The removed item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RemovedItem,
	///       ExTVwLibU::_IExplorerTreeViewEvents::RemovedItem, Raise_RemovingItem, Raise_InsertedItem
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RemovedItem,
	///       ExTVwLibA::_IExplorerTreeViewEvents::RemovedItem, Raise_RemovingItem, Raise_InsertedItem
	/// \endif
	inline HRESULT Raise_RemovedItem(IVirtualTreeViewItem* pTreeItem);
	/// \brief <em>Raises the \c RemovingItem event</em>
	///
	/// \param[in] pTreeItem The item being removed.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort deletion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RemovingItem,
	///       ExTVwLibU::_IExplorerTreeViewEvents::RemovingItem, Raise_RemovedItem, Raise_InsertingItem
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RemovingItem,
	///       ExTVwLibA::_IExplorerTreeViewEvents::RemovingItem, Raise_RemovedItem, Raise_InsertingItem
	/// \endif
	inline HRESULT Raise_RemovingItem(ITreeViewItem* pTreeItem, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c RenamedItem event</em>
	///
	/// \param[in] pTreeItem The renamed item.
	/// \param[in] previousItemText The item's previous text.
	/// \param[in] newItemText The item's new text.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RenamedItem,
	///       ExTVwLibU::_IExplorerTreeViewEvents::RenamedItem, Raise_RenamingItem,
	///       Raise_StartingLabelEditing, Raise_DestroyedEditControlWindow
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RenamedItem,
	///       ExTVwLibA::_IExplorerTreeViewEvents::RenamedItem, Raise_RenamingItem,
	///       Raise_StartingLabelEditing, Raise_DestroyedEditControlWindow
	/// \endif
	inline HRESULT Raise_RenamedItem(ITreeViewItem* pTreeItem, BSTR previousItemText, BSTR newItemText);
	/// \brief <em>Raises the \c RenamingItem event</em>
	///
	/// \param[in] pTreeItem The item being renamed.
	/// \param[in] previousItemText The item's previous text.
	/// \param[in] newItemText The item's new text.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort renaming; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RenamingItem,
	///       ExTVwLibU::_IExplorerTreeViewEvents::RenamingItem, Raise_RenamedItem
	///       Raise_StartingLabelEditing, Raise_DestroyedEditControlWindow
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_RenamingItem,
	///       ExTVwLibA::_IExplorerTreeViewEvents::RenamingItem, Raise_RenamedItem
	///       Raise_StartingLabelEditing, Raise_DestroyedEditControlWindow
	/// \endif
	inline HRESULT Raise_RenamingItem(ITreeViewItem* pTreeItem, BSTR previousItemText, BSTR newItemText, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ResizedControlWindow,
	///       ExTVwLibU::_IExplorerTreeViewEvents::ResizedControlWindow
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_ResizedControlWindow,
	///       ExTVwLibA::_IExplorerTreeViewEvents::ResizedControlWindow
	/// \endif
	inline HRESULT Raise_ResizedControlWindow(void);
	/// \brief <em>Raises the \c SingleExpanding event</em>
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
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_SingleExpanding,
	///       ExTVwLibU::_IExplorerTreeViewEvents::SingleExpanding, Raise_CaretChanging
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_SingleExpanding,
	///       ExTVwLibA::_IExplorerTreeViewEvents::SingleExpanding, Raise_CaretChanging
	/// \endif
	inline HRESULT Raise_SingleExpanding(ITreeViewItem* pPreviousCaretItem, ITreeViewItem* pNewCaretItem, VARIANT_BOOL* pDontChangePreviousItem, VARIANT_BOOL* pDontChangeNewItem);
	/// \brief <em>Raises the \c StartingLabelEditing event</em>
	///
	/// \param[in] pTreeItem The item being edited.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should prevent label-editing; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_StartingLabelEditing,
	///       ExTVwLibU::_IExplorerTreeViewEvents::StartingLabelEditing, Raise_RenamingItem,
	///       Raise_CreatedEditControlWindow
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_StartingLabelEditing,
	///       ExTVwLibA::_IExplorerTreeViewEvents::StartingLabelEditing, Raise_RenamingItem,
	///       Raise_CreatedEditControlWindow
	/// \endif
	inline HRESULT Raise_StartingLabelEditing(ITreeViewItem* pTreeItem, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c XClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_XClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::XClick, Raise_XDblClick, Raise_Click, Raise_MClick,
	///       Raise_RClick, ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_XClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::XClick, Raise_XDblClick, Raise_Click, Raise_MClick,
	///       Raise_RClick, ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c XDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_XDblClick,
	///       ExTVwLibU::_IExplorerTreeViewEvents::XDblClick, Raise_XClick, Raise_DblClick,
	///       Raise_MDblClick, Raise_RDblClick, ExTVwLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IExplorerTreeViewEvents::Fire_XDblClick,
	///       ExTVwLibA::_IExplorerTreeViewEvents::XDblClick, Raise_XClick, Raise_DblClick,
	///       Raise_MDblClick, Raise_RDblClick, ExTVwLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Applies the font to the control</em>
	///
	/// This method sets the control's font to the font specified by the \c Font property.
	///
	/// \sa get_Font
	void ApplyFont(void);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleObject
	///
	//@{
	/// \brief <em>Implementation of \c IOleObject::DoVerb</em>
	///
	/// Will be called if one of the control's registered actions (verbs) shall be executed; e. g.
	/// executing the "About" verb will display the About dialog.
	///
	/// \param[in] verbID The requested verb's ID.
	/// \param[in] pMessage A \c MSG structure describing the event (such as a double-click) that
	///            invoked the verb.
	/// \param[in] pActiveSite The \c IOleClientSite implementation of the control's active client site
	///            where the event occurred that invoked the verb.
	/// \param[in] reserved Reserved; must be zero.
	/// \param[in] hWndParent The handle of the document window containing the control.
	/// \param[in] pBoundingRectangle A \c RECT structure containing the coordinates and size in pixels
	///            of the control within the window specified by \c hWndParent.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa EnumVerbs,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694508.aspx">IOleObject::DoVerb</a>
	virtual HRESULT STDMETHODCALLTYPE DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle);
	/// \brief <em>Implementation of \c IOleObject::EnumVerbs</em>
	///
	/// Will be called if the control's container requests the control's registered actions (verbs); e. g.
	/// we provide a verb "About" that will display the About dialog on execution.
	///
	/// \param[out] ppEnumerator Receives the \c IEnumOLEVERB implementation of the verbs' enumerator.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692781.aspx">IOleObject::EnumVerbs</a>
	virtual HRESULT STDMETHODCALLTYPE EnumVerbs(IEnumOLEVERB** ppEnumerator);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleControl
	///
	//@{
	/// \brief <em>Implementation of \c IOleControl::GetControlInfo</em>
	///
	/// Will be called if the container requests details about the control's keyboard mnemonics and keyboard
	/// behavior.
	///
	/// \param[in, out] pControlInfo The requested details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms693730.aspx">IOleControl::GetControlInfo</a>
	virtual HRESULT STDMETHODCALLTYPE GetControlInfo(LPCONTROLINFO pControlInfo);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Executes the About verb</em>
	///
	/// Will be called if the control's registered action (verb) "About" shall be executed. We'll
	/// display the About dialog.
	///
	/// \param[in] hWndParent The window to use as parent for any user interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb, About
	HRESULT DoVerbAbout(HWND hWndParent);

	//////////////////////////////////////////////////////////////////////
	/// \name Caret changes
	///
	//@{
	/// \brief <em>Adds a new \c ICaretChangeTask object to the task list</em>
	///
	/// Adds a new entry to the list of tasks that will be executed on caret changes.
	/// We use this method for the \c TreeViewItem::Move feature.
	///
	/// \param[in] executeOnCaretChanging If \c TRUE, the task is executed on receiving the
	///            \c TVN_SELCHANGING notification.
	/// \param[in] executeOnCaretChanged If \c TRUE, the task is executed on receiving the
	///            \c TVN_SELCHANGED notification.
	/// \param[in] pTask The task object's \c ICaretChangeTask implementation.
	///
	/// \sa CaretChangeTasks::AddTask, AddInternalCaretChange, TreeViewItem::Move
	void AddCaretChangeTask(BOOL executeOnCaretChanging, BOOL executeOnCaretChanged, ICaretChangeTask* pTask);
	/// \brief <em>Marks an outstanding caret change as an internal one</em>
	///
	/// Adds a new entry to the list of internal caret changes. An internal caret change is a caret
	/// change that was initiated by ourselves (usually by sending \c TVM_SELECTITEM with \c TVGN_CARET
	/// specified).
	///
	/// \param[in] hNewCaretItem The new caret item's handle which will be used to identify the caret
	///            change.
	/// \param[in] realCaretChangeCause A \c TVC_* constant specifying the caret change's reason that
	///            will be reported by the \c CaretChanging and \c CaretChanged events. If set to -1,
	///            no event will be raised at all.
	///
	/// \sa TVC_INTERNAL, RemoveInternalCaretChange, AddCaretChangeTask, Raise_CaretChanging,
	///     Raise_CaretChanged
	void AddInternalCaretChange(HTREEITEM hNewCaretItem, UINT realCaretChangeCause = static_cast<UINT>(-1));
	/// \brief <em>Removes an internal caret change from the list</em>
	///
	/// Removes an entry from the list of internal caret changes. An internal caret change is a caret
	/// change that was initiated by ourselves (usually by sending \c TVM_SELECTITEM with \c TVGN_CARET
	/// specified).
	///
	/// \param[in] hNewCaretItem The new caret item's handle which will be used to identify the caret
	///            change to remove.
	///
	/// \sa AddInternalCaretChange
	void RemoveInternalCaretChange(HTREEITEM hNewCaretItem);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Item insertion and deletion
	///
	//@{
	/// \brief <em>Increments the \c silentItemInsertions flag</em>
	///
	/// \sa LeaveSilentItemInsertionSection, Flags::silentItemInsertions, AddSilentItemDeletion,
	///     TreeViewItem::Move, Raise_InsertingItem, Raise_InsertedItem
	void EnterSilentItemInsertionSection(void);
	/// \brief <em>Decrements the \c silentItemInsertions flag</em>
	///
	/// \sa EnterSilentItemInsertionSection, Flags::silentItemInsertions, AddSilentItemDeletion,
	///     TreeViewItem::Move, Raise_InsertingItem, Raise_InsertedItem
	void LeaveSilentItemInsertionSection(void);
	/// \brief <em>Marks an item deletion as silent thus it won't raise any event</em>
	///
	/// Adds a new entry to the list of silent item deletions. A silent item deletion doesn't raise
	/// the \c RemovingItem and \c RemovedItem events.
	/// We use silent item deletions for the \c TreeViewItem::Move feature.
	///
	/// \param[in] hDeletedItem The handle of the item that is being removed.
	///
	/// \sa LeaveSilentItemInsertionSection, TreeViewItem::Move, Raise_RemovingItem, Raise_RemovedItem
	void AddSilentItemDeletion(HTREEITEM hDeletedItem);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Item selection
	///
	//@{
	/// \brief <em>Increments the \c dontSingleExpand flag</em>
	///
	/// \sa LeaveNoSingleExpandSection, Flags::dontSingleExpand, OnSingleExpandNotification
	void EnterNoSingleExpandSection(void);
	/// \brief <em>Decrements the \c dontSingleExpand flag</em>
	///
	/// \sa EnterNoSingleExpandSection, Flags::dontSingleExpand, OnSingleExpandNotification
	void LeaveNoSingleExpandSection(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name MFC clones
	///
	//@{
	/// \brief <em>A rewrite of MFC's \c COleControl::RecreateControlWindow</em>
	///
	/// Destroys and re-creates the control window.
	///
	/// \remarks This rewrite probably isn't complete.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/5tx8h2fd.aspx">COleControl::RecreateControlWindow</a>
	void RecreateControlWindow(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name IME support
	///
	//@{
	/// \brief <em>Retrieves a window's current IME context mode</em>
	///
	/// \param[in] hWnd The window whose IME context mode is requested.
	///
	/// \return A constant out of the \c IMEModeConstants enumeration specifying the IME context mode.
	///
	/// \if UNICODE
	///   \sa SetCurrentIMEContextMode, ExTVwLibU::IMEModeConstants, get_IMEMode, get_EditIMEMode
	/// \else
	///   \sa SetCurrentIMEContextMode, ExTVwLibA::IMEModeConstants, get_IMEMode, get_EditIMEMode
	/// \endif
	IMEModeConstants GetCurrentIMEContextMode(HWND hWnd);
	/// \brief <em>Sets a window's current IME context mode</em>
	///
	/// \param[in] hWnd The window whose IME context mode is set.
	/// \param[in] IMEMode A constant out of the \c IMEModeConstants enumeration specifying the IME
	///            context mode to apply.
	///
	/// \if UNICODE
	///   \sa GetCurrentIMEContextMode, ExTVwLibU::IMEModeConstants, put_IMEMode, put_EditIMEMode
	/// \else
	///   \sa GetCurrentIMEContextMode, ExTVwLibA::IMEModeConstants, put_IMEMode, put_EditIMEMode
	/// \endif
	void SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode);
	/// \brief <em>Retrieves the control's effective IME context mode</em>
	///
	/// Retrieves the IME context mode that is set for the control after resolving recursive modes like
	/// \c imeInherit.
	///
	/// \return A constant out of the \c IMEModeConstants enumeration specifying the IME context mode.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::IMEModeConstants, GetEffectiveIMEMode_Edit, get_IMEMode
	/// \else
	///   \sa ExTVwLibA::IMEModeConstants, GetEffectiveIMEMode_Edit, get_IMEMode
	/// \endif
	IMEModeConstants GetEffectiveIMEMode(void);
	/// \brief <em>Retrieves the label-edit control's effective IME context mode</em>
	///
	/// Retrieves the IME context mode that is set for the label-edit control after resolving recursive
	/// modes like \c imeInherit.
	///
	/// \return A constant out of the \c IMEModeConstants enumeration specifying the IME context mode.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::IMEModeConstants, GetEffectiveIMEMode, get_EditIMEMode
	/// \else
	///   \sa ExTVwLibA::IMEModeConstants, GetEffectiveIMEMode, get_EditIMEMode
	/// \endif
	IMEModeConstants GetEffectiveIMEMode_Edit(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Control window configuration
	///
	//@{
	/// \brief <em>Calculates the extended window style bits</em>
	///
	/// Calculates the extended window style bits that the control must have set to match the current
	/// property settings.
	///
	/// \return A bit field of \c WS_EX_* constants specifying the required extended window styles.
	///
	/// \sa GetStyleBits, SendConfigurationMessages, Create
	DWORD GetExStyleBits(void);
	/// \brief <em>Calculates the window style bits</em>
	///
	/// Calculates the window style bits that the control must have set to match the current property
	/// settings.
	///
	/// \return A bit field of \c WS_* and \c TVS_* constants specifying the required window styles.
	///
	/// \sa GetExStyleBits, SendConfigurationMessages, Create
	DWORD GetStyleBits(void);
	/// \brief <em>Configures the control window by sending messages</em>
	///
	/// Sends \c WM_* and \c TVM_* messages to the control window to make it match the current property
	/// settings. Will be called out of \c Raise_RecreatedControlWindow.
	///
	/// \sa GetExStyleBits, GetStyleBits, Raise_RecreatedControlWindow
	void SendConfigurationMessages(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Value translation
	///
	//@{
	/// \brief <em>Translates a \c MousePointerConstants type mouse cursor into a \c HCURSOR type mouse cursor</em>
	///
	/// \param[in] mousePointer The \c MousePointerConstants type mouse cursor to translate.
	///
	/// \return The translated \c HCURSOR type mouse cursor.
	///
	/// \if UNICODE
	///   \sa ExTVwLibU::MousePointerConstants, OnSetCursorNotification, put_MousePointer
	/// \else
	///   \sa ExTVwLibA::MousePointerConstants, OnSetCursorNotification, put_MousePointer
	/// \endif
	HCURSOR MousePointerConst2hCursor(MousePointerConstants mousePointer);
	#ifdef USE_STL
		/// \brief <em>Translates an item's unique ID to its handle</em>
		///
		/// \param[in] ID The unique ID of the item whose handle is requested.
		///
		/// \return The requested item's handle if successful; \c NULL otherwise.
		///
		/// \sa itemHandles, ItemHandleToID, TreeViewItem::get_ID, TreeViewItem::get_Handle,
		///     TreeViewItems::get_Item, TreeViewItems::Remove
		HTREEITEM IDToItemHandle(std::unordered_map<LONG, HTREEITEM>::key_type ID);
	#else
		/// \brief <em>Translates an item's unique ID to its handle</em>
		///
		/// \param[in] ID The unique ID of the item whose handle is requested.
		///
		/// \return The requested item's handle if successful; \c NULL otherwise.
		///
		/// \sa itemHandles, ItemHandleToID, TreeViewItem::get_ID, TreeViewItem::get_Handle,
		///     TreeViewItems::get_Item, TreeViewItems::Remove
		HTREEITEM IDToItemHandle(CAtlMap<LONG, HTREEITEM>::KINARGTYPE ID);
	#endif
	/// \brief <em>Translates an item's handle to its unique ID</em>
	///
	/// \param[in] hItem The handle of the item whose unique ID is requested.
	///
	/// \return The requested item's unique ID if successful; 0 otherwise.
	///
	/// \sa itemHandles, IDToItemHandle, TreeViewItem::get_ID, TreeViewItem::get_Handle
	LONG ItemHandleToID(HTREEITEM hItem);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the item that lies below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[out] pFlags A bit field of TVHT_* flags, that holds further details about the control's
	///             part below the specified point.
	/// \param[in] ignoreBoundingBoxDefinition If \c TRUE the setting of the \c ItemBoundingBoxDefinition
	///            property is ignored; otherwise the returned item handle is set to \c NULL, if the
	///            \c ItemBoundingBoxDefinition property's setting and the \c pFlags parameter don't
	///            match.
	///
	/// \return The "hit" item's handle.
	///
	/// \sa HitTest, get_ItemBoundingBoxDefinition
	HTREEITEM HitTest(LONG x, LONG y, UINT* pFlags, BOOL ignoreBoundingBoxDefinition = FALSE);
	/// \brief <em>Retrieves whether we're in design mode or in user mode</em>
	///
	/// \return \c TRUE if the control is in design mode (i. e. is placed on a window which is edited
	///         by a form editor); \c FALSE if the control is in user mode (i. e. is placed on a window
	///         being used by an end-user).
	BOOL IsInDesignMode(void);
	/// \brief <em>Replaces an item handle within the \c itemHandles and \c itemParams hash maps with another item handle</em>
	///
	/// \param[in] oldItemHandle The item handle to replace.
	/// \param[in] newItemHandle The item handle to replace with.
	///
	/// \sa itemHandles, itemParams, TreeViewItem::Move, IDToItemHandle, ItemHandleToID
	void ReplaceItemHandle(HTREEITEM oldItemHandle, HTREEITEM newItemHandle);
	/// \brief <em>Auto-scrolls the control</em>
	///
	/// \sa OnTimer, Raise_DragMouseMove, DragDropStatus::AutoScrolling
	void AutoScroll(void);

	/// \brief <em>Retrieves whether a given item has at least one sub-item</em>
	///
	/// \param[in] hItem The item to check.
	///
	/// \return \c TRUE if the item has at least one sub-item; otherwise \c FALSE.
	BOOL HasSubItems(HTREEITEM hItem);
	/// \brief <em>Retrieves whether two items are immediate sub-items of the same parent item</em>
	///
	/// \param[in] hItem1 The first item to check.
	/// \param[in] hItem2 The second item to check.
	///
	/// \return \c TRUE if the items have the same immediate parent item; otherwise \c FALSE.
	BOOL HaveSameParentItem(HTREEITEM hItem1, HTREEITEM hItem2);
	/// \brief <em>Retrieves whether a given item is expanded</em>
	///
	/// \param[in] hItem The item to check.
	/// \param[in] allowPartialExpansion If set to \c TRUE, the method will return \c TRUE, if the item
	///            is either partially (\c TVIS_EXPANDPARTIAL) or entirely (\c TVIS_EXPANDED) expanded;
	///            otherwise it <strong>must</strong> be entirely expanded.
	///
	/// \return \c TRUE if the item is entirely or partial expanded; otherwise \c FALSE.
	///
	/// \sa TreeViewItem::get_ExpansionState
	BOOL IsExpanded(HTREEITEM hItem, BOOL allowPartialExpansion = TRUE);
	/// \brief <em>Retrieves whether the logical left mouse button is held down</em>
	///
	/// \return \c TRUE if the logical left mouse button is held down; otherwise \c FALSE.
	///
	/// \sa IsRightMouseButtonDown
	BOOL IsLeftMouseButtonDown(void);
	/// \brief <em>Retrieves whether the logical right mouse button is held down</em>
	///
	/// \return \c TRUE if the logical right mouse button is held down; otherwise \c FALSE.
	///
	/// \sa IsLeftMouseButtonDown
	BOOL IsRightMouseButtonDown(void);


	/// \brief <em>Holds constants and flags used with IME support</em>
	struct IMEFlags
	{
	protected:
		/// \brief <em>A table of IME modes to use for Chinese input language</em>
		///
		/// \sa GetIMECountryTable, japaneseIMETable, koreanIMETable
		static IMEModeConstants chineseIMETable[10];
		/// \brief <em>A table of IME modes to use for Japanese input language</em>
		///
		/// \sa GetIMECountryTable, chineseIMETable, koreanIMETable
		static IMEModeConstants japaneseIMETable[10];
		/// \brief <em>A table of IME modes to use for Korean input language</em>
		///
		/// \sa GetIMECountryTable, chineseIMETable, koreanIMETable
		static IMEModeConstants koreanIMETable[10];

	public:
		/// \brief <em>The handle of the default IME context</em>
		HIMC hDefaultIMC;

		IMEFlags()
		{
			hDefaultIMC = NULL;
		}

		/// \brief <em>Retrieves a table of IME modes to use for the current keyboard layout</em>
		///
		/// Retrieves a table of IME modes which can be used to map \c IME_CMODE_* constants to
		/// \c IMEModeConstants constants. The table depends on the current keyboard layout.
		///
		/// \param[in,out] table The IME mode table for the currently active keyboard layout.
		///
		/// \if UNICODE
		///   \sa ExTVwLibU::IMEModeConstants, GetCurrentIMEContextMode
		/// \else
		///   \sa ExTVwLibA::IMEModeConstants, GetCurrentIMEContextMode
		/// \endif
		static void GetIMECountryTable(IMEModeConstants table[10])
		{
			WORD languageID = LOWORD(GetKeyboardLayout(0));
			if(languageID <= MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED)) {
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_TRADITIONAL)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				switch(languageID) {
					case MAKELANGID(LANG_JAPANESE, SUBLANG_DEFAULT):
						CopyMemory(table, japaneseIMETable, sizeof(japaneseIMETable));
						return;
						break;
					case MAKELANGID(LANG_KOREAN, SUBLANG_DEFAULT):
						CopyMemory(table, koreanIMETable, sizeof(koreanIMETable));
						return;
						break;
				}
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			if(languageID <= MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_HONGKONG)) {
				if(languageID == MAKELANGID(LANG_KOREAN, SUBLANG_SYS_DEFAULT)) {
					CopyMemory(table, koreanIMETable, sizeof(koreanIMETable));
					return;
				}
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_HONGKONG)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			if((languageID != MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SINGAPORE)) && (languageID != MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_MACAU))) {
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
		}
	} IMEFlags;

	/// \brief <em>Holds a control instance's properties' settings</em>
	typedef struct Properties
	{
		/// \brief <em>Holds a font property's settings</em>
		typedef struct FontProperty
		{
		protected:
			/// \brief <em>Holds the control's default font</em>
			///
			/// \sa GetDefaultFont
			static FONTDESC defaultFont;

		public:
			/// \brief <em>Holds whether we're listening for events fired by the font object</em>
			///
			/// If greater than 0, we're advised to the \c IFontDisp object identified by \c pFontDisp. I. e.
			/// we will be notified if a property of the font object changes. If 0, we won't receive any events
			/// fired by the \c IFontDisp object.
			///
			/// \sa pFontDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>Flag telling \c OnSetFont not to retrieve the current font if set to \c TRUE</em>
			///
			/// \sa OnSetFont
			UINT dontGetFontObject : 1;
			/// \brief <em>The control's current font</em>
			///
			/// \sa OnSetFont, ApplyFont, owningFontResource
			CFont currentFont;
			/// \brief <em>If \c TRUE, \c currentFont may destroy the font resource; otherwise not</em>
			///
			/// \sa currentFont
			UINT owningFontResource : 1;
			/// \brief <em>A pointer to the font object's implementation of \c IFontDisp</em>
			IFontDisp* pFontDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<ExplorerTreeView> >* pPropertyNotifySink;

			FontProperty()
			{
				watching = 0;
				dontGetFontObject = FALSE;
				owningFontResource = TRUE;
				pFontDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~FontProperty()
			{
				Release();
			}

			FontProperty& operator =(const FontProperty& source)
			{
				Release();

				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				pFontDisp = source.pFontDisp;
				if(pFontDisp) {
					pFontDisp->AddRef();
				}
				owningFontResource = source.owningFontResource;
				if(!owningFontResource) {
					currentFont.Attach(source.currentFont.m_hFont);
				}
				dontGetFontObject = source.dontGetFontObject;

				if(source.watching > 0) {
					StartWatching();
				}

				return *this;
			}

			/// \brief <em>Retrieves a default font that may be used</em>
			///
			/// \return A \c FONTDESC structure containing the default font.
			///
			/// \sa defaultFont
			static FONTDESC GetDefaultFont(void)
			{
				return defaultFont;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(ExplorerTreeView* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<ExplorerTreeView> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(owningFontResource) {
					if(!currentFont.IsNull()) {
						currentFont.DeleteObject();
					}
				} else {
					currentFont.Detach();
				}
				if(pFontDisp) {
					pFontDisp->Release();
					pFontDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pFontDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} FontProperty;

		/// \brief <em>Holds a picture property's settings</em>
		typedef struct PictureProperty
		{
			/// \brief <em>Holds whether we're listening for events fired by the picture object</em>
			///
			/// If greater than 0, we're advised to the \c IPictureDisp object identified by \c pPictureDisp.
			/// I. e. we will be notified if a property of the picture object changes. If 0, we won't receive any
			/// events fired by the \c IPictureDisp object.
			///
			/// \sa pPictureDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>A pointer to the picture object's implementation of \c IPictureDisp</em>
			IPictureDisp* pPictureDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<ExplorerTreeView> >* pPropertyNotifySink;

			PictureProperty()
			{
				watching = 0;
				pPictureDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~PictureProperty()
			{
				Release();
			}

			PictureProperty& operator =(const PictureProperty& source)
			{
				Release();

				pPictureDisp = source.pPictureDisp;
				if(pPictureDisp) {
					pPictureDisp->AddRef();
				}
				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				if(source.watching > 0) {
					StartWatching();
				}
				return *this;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(ExplorerTreeView* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<ExplorerTreeView> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(pPictureDisp) {
					pPictureDisp->Release();
					pPictureDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pPictureDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} PictureProperty;

		/// \brief <em>The autoscroll-zone's default width</em>
		///
		/// The default width (in pixels) of the border around the control's client area, that's
		/// sensitive for auto-scrolling during a drag'n'drop operation. If the mouse cursor's position
		/// lies within this area during a drag'n'drop operation, the control will be auto-scrolled.
		///
		/// \sa dragScrollTimeBase, Raise_DragMouseMove
		static const int DRAGSCROLLZONEWIDTH = 16;

		/// \brief <em>Holds the \c AllowDragDrop property's setting</em>
		///
		/// \sa get_AllowDragDrop, put_AllowDragDrop
		UINT allowDragDrop : 1;
		/// \brief <em>Holds the \c AllowLabelEditing property's setting</em>
		///
		/// \sa get_AllowLabelEditing, put_AllowLabelEditing
		UINT allowLabelEditing : 1;
		/// \brief <em>Holds the \c AlwaysShowSelection property's setting</em>
		///
		/// \sa get_AlwaysShowSelection, put_AlwaysShowSelection
		UINT alwaysShowSelection : 1;
		/// \brief <em>Holds the \c Appearance property's setting</em>
		///
		/// \sa get_Appearance, put_Appearance
		AppearanceConstants appearance;
		/// \brief <em>Holds the \c AutoHScroll property's setting</em>
		///
		/// \sa get_AutoHScroll, put_AutoHScroll
		UINT autoHScroll : 1;
		/// \brief <em>Holds the \c AutoHScrollPixelsPerSecond property's setting</em>
		///
		/// \sa get_AutoHScrollPixelsPerSecond, put_AutoHScrollPixelsPerSecond
		long autoHScrollPixelsPerSecond;
		/// \brief <em>Holds the \c AutoHScrollRedrawInterval property's setting</em>
		///
		/// \sa get_AutoHScrollRedrawInterval, put_AutoHScrollRedrawInterval
		long autoHScrollRedrawInterval;
		/// \brief <em>Holds the \c BackColor property's setting</em>
		///
		/// \sa get_BackColor, put_BackColor
		OLE_COLOR backColor;
		/// \brief <em>Holds the \c BlendSelectedItemsIcons property's setting</em>
		///
		/// \sa get_BlendSelectedItemsIcons, put_BlendSelectedItemsIcons
		UINT blendSelectedItemsIcons : 1;
		/// \brief <em>Holds the \c BorderStyle property's setting</em>
		///
		/// \sa get_BorderStyle, put_BorderStyle
		BorderStyleConstants borderStyle;
		/// \brief <em>Holds the \c BuiltInStateImages property's setting</em>
		///
		/// \sa get_BuiltInStateImages, put_BuiltInStateImages
		BuiltInStateImagesConstants builtInStateImages;
		/// \brief <em>Holds the \c CaretChangedDelayTime property's setting</em>
		///
		/// \sa get_CaretChangedDelayTime, put_CaretChangedDelayTime
		long caretChangedDelayTime;
		/// \brief <em>Holds the \c DisabledEvents property's setting</em>
		///
		/// \sa get_DisabledEvents, put_DisabledEvents
		DisabledEventsConstants disabledEvents;
		/// \brief <em>Holds the \c DontRedraw property's setting</em>
		///
		/// \sa get_DontRedraw, put_DontRedraw
		UINT dontRedraw : 1;
		/// \brief <em>Holds the \c DragExpandTime property's setting</em>
		///
		/// \sa get_DragExpandTime, put_DragExpandTime
		long dragExpandTime;
		/// \brief <em>Holds the \c DragScrollTimeBase property's setting</em>
		///
		/// \sa get_DragScrollTimeBase, put_DragScrollTimeBase
		long dragScrollTimeBase;
		/// \brief <em>Holds the \c DrawImagesAsynchronously property's setting</em>
		///
		/// \sa get_DrawImagesAsynchronously, put_DrawImagesAsynchronously
		UINT drawImagesAsynchronously : 1;
		/// \brief <em>Holds the \c EditBackColor property's setting</em>
		///
		/// \sa get_EditBackColor, put_EditBackColor
		OLE_COLOR editBackColor;
		/// \brief <em>Holds the \c EditForeColor property's setting</em>
		///
		/// \sa get_EditForeColor, put_EditForeColor
		OLE_COLOR editForeColor;
		/// \brief <em>Holds the \c EditHoverTime property's setting</em>
		///
		/// \sa get_EditHoverTime, put_EditHoverTime
		long editHoverTime;
		/// \brief <em>Holds the \c EditIMEMode property's setting</em>
		///
		/// \sa get_EditIMEMode, put_EditIMEMode
		IMEModeConstants editIMEMode;
		/// \brief <em>Holds the \c Enabled property's setting</em>
		///
		/// \sa get_Enabled, put_Enabled
		UINT enabled : 1;
		/// \brief <em>Holds the \c FadeExpandos property's setting</em>
		///
		/// \sa get_FadeExpandos, put_FadeExpandos
		UINT fadeExpandos : 1;
		/// \brief <em>Holds the \c FavoritesStyle property's setting</em>
		///
		/// \sa get_FavoritesStyle, put_FavoritesStyle
		UINT favoritesStyle : 1;
		/// \brief <em>Holds the \c Font property's settings</em>
		///
		/// \sa get_Font, put_Font, putref_Font
		FontProperty font;
		/// \brief <em>Holds the \c ForeColor property's setting</em>
		///
		/// \sa get_ForeColor, put_ForeColor
		OLE_COLOR foreColor;
		/// \brief <em>Holds the \c FullRowSelect property's setting</em>
		///
		/// \sa get_FullRowSelect, put_FullRowSelect
		UINT fullRowSelect : 1;
		/// \brief <em>Holds the \c GroupBoxColor property's setting</em>
		///
		/// \sa get_GroupBoxColor, put_GroupBoxColor
		OLE_COLOR groupBoxColor;
		/// \brief <em>Holds the \c HotTracking property's setting</em>
		///
		/// \sa get_HotTracking, put_HotTracking
		HotTrackingConstants hotTracking;
		/// \brief <em>Holds the \c HoverTime property's setting</em>
		///
		/// \sa get_HoverTime, put_HoverTime
		long hoverTime;
		/// \brief <em>Holds the contained label-edit control's window handle</em>
		///
		/// \sa get_hWndEdit, Raise_CreatedEditControlWindow, Raise_DestroyedEditControlWindow
		HWND hWndEdit;
		/// \brief <em>Holds the \c IMEMode property's setting</em>
		///
		/// \sa get_IMEMode, put_IMEMode
		IMEModeConstants IMEMode;
		/// \brief <em>Holds the \c Indent property's setting</em>
		///
		/// \sa get_Indent, put_Indent
		OLE_XSIZE_PIXELS indent;
		/// \brief <em>Holds the \c IndentStateImages property's setting</em>
		///
		/// \sa get_IndentStateImages, put_IndentStateImages
		UINT indentStateImages : 1;
		/// \brief <em>Holds the \c InsertMarkColor property's setting</em>
		///
		/// \sa get_InsertMarkColor, put_InsertMarkColor
		OLE_COLOR insertMarkColor;
		/// \brief <em>Holds the \c ItemXBorder and \c ItemYBorder properties' settings</em>
		///
		/// Stores the \c ItemXBorder property's setting in the lower word and the \c ItemYBorder property's
		/// setting in the upper word.
		///
		/// \sa get_ItemXBorder, put_ItemXBorder, get_ItemYBorder, put_ItemYBorder
		DWORD itemBorders;
		/// \brief <em>Holds the \c ItemBoundingBoxDefinition property's setting</em>
		///
		/// \sa get_ItemBoundingBoxDefinition, put_ItemBoundingBoxDefinition
		ItemBoundingBoxDefinitionConstants itemBoundingBoxDefinition;
		/// \brief <em>Holds the \c ItemHeight property's setting</em>
		///
		/// \sa get_ItemHeight, put_ItemHeight
		OLE_YSIZE_PIXELS itemHeight;
		/// \brief <em>Holds the \c LineColor property's setting</em>
		///
		/// \sa get_LineColor, put_LineColor
		OLE_COLOR lineColor;
		/// \brief <em>Holds the \c LineStyle property's setting</em>
		///
		/// \sa get_LineStyle, put_LineStyle
		LineStyleConstants lineStyle;
		/// \brief <em>Holds the \c Locale property's setting</em>
		///
		/// \sa get_Locale, put_Locale
		LCID locale;
		/// \brief <em>Holds the \c MaxScrollTime property's setting</em>
		///
		/// \sa get_MaxScrollTime, put_MaxScrollTime
		long maxScrollTime;
		/// \brief <em>Holds the \c MouseIcon property's settings</em>
		///
		/// \sa get_MouseIcon, put_MouseIcon, putref_MouseIcon
		PictureProperty mouseIcon;
		/// \brief <em>Holds the \c MousePointer property's setting</em>
		///
		/// \sa get_MousePointer, put_MousePointer
		MousePointerConstants mousePointer;
		/// \brief <em>Holds the \c OLEDragImageStyle property's setting</em>
		///
		/// \sa get_OLEDragImageStyle, put_OLEDragImageStyle
		OLEDragImageStyleConstants oleDragImageStyle;
		/// \brief <em>Holds the \c ProcessContextMenuKeys property's setting</em>
		///
		/// \sa get_ProcessContextMenuKeys, put_ProcessContextMenuKeys
		UINT processContextMenuKeys : 1;
		/// \brief <em>Holds the \c RegisterForOLEDragDrop property's setting</em>
		///
		/// \sa get_RegisterForOLEDragDrop, put_RegisterForOLEDragDrop
		UINT registerForOLEDragDrop : 1;
		/// \brief <em>Holds the \c RichToolTips property's setting</em>
		///
		/// \sa get_RichToolTips, put_RichToolTips
		UINT richToolTips : 1;
		/// \brief <em>Holds the \c RightToLeft property's setting</em>
		///
		/// \sa get_RightToLeft, put_RightToLeft
		RightToLeftConstants rightToLeft;
		/// \brief <em>Holds the \c ScrollBars property's setting</em>
		///
		/// \sa get_ScrollBars, put_ScrollBars
		ScrollBarsConstants scrollBars;
		/// \brief <em>Holds the \c ShowStateImages property's setting</em>
		///
		/// \sa get_ShowStateImages, put_ShowStateImages
		UINT showStateImages : 1;
		/// \brief <em>Holds the \c ShowToolTips property's setting</em>
		///
		/// \sa get_ShowToolTips, put_ShowToolTips
		UINT showToolTips : 1;
		/// \brief <em>Holds the \c SingleExpand property's setting</em>
		///
		/// \sa get_SingleExpand, put_SingleExpand
		SingleExpandConstants singleExpand;
		/// \brief <em>Holds the \c SortOrder property's setting</em>
		///
		/// \sa get_SortOrder, put_SortOrder
		SortOrderConstants sortOrder;
		/// \brief <em>Holds the \c SupportOLEDragImages property's setting</em>
		///
		/// \sa get_SupportOLEDragImages, put_SupportOLEDragImages
		UINT supportOLEDragImages : 1;
		/// \brief <em>Holds the \c TextParsingFlags property's setting</em>
		///
		/// \sa get_TextParsingFlags, put_TextParsingFlags
		DWORD textParsingFlagsForCompareString;
		/// \brief <em>Holds the \c TextParsingFlags property's setting</em>
		///
		/// \sa get_TextParsingFlags, put_TextParsingFlags
		ULONG textParsingFlagsForVarFromStr;
		/// \brief <em>Holds the \c TreeViewStyle property's setting</em>
		///
		/// \sa get_TreeViewStyle, put_TreeViewStyle
		TreeViewStyleConstants treeViewStyle;
		/// \brief <em>Holds the \c UseSystemFont property's setting</em>
		///
		/// \sa get_UseSystemFont, put_UseSystemFont
		UINT useSystemFont : 1;

		Properties()
		{
			ResetToDefaults();
		}

		~Properties()
		{
			Release();
		}

		/// \brief <em>Prepares the object for destruction</em>
		void Release(void)
		{
			font.Release();
			mouseIcon.Release();
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			allowDragDrop = TRUE;
			allowLabelEditing = TRUE;
			alwaysShowSelection = TRUE;
			appearance = a3D;
			autoHScroll = TRUE;
			autoHScrollPixelsPerSecond = 150;
			autoHScrollRedrawInterval = 15;
			backColor = 0x80000000 | COLOR_WINDOW;
			blendSelectedItemsIcons = FALSE;
			borderStyle = bsNone;
			builtInStateImages = bisiSelectedAndUnselected;
			caretChangedDelayTime = 500;
			disabledEvents = static_cast<DisabledEventsConstants>(deTreeMouseEvents | deTreeClickEvents | deEditMouseEvents | deTreeKeyboardEvents | deEditKeyboardEvents | deItemInsertionEvents | deItemDeletionEvents | deFreeItemData | deCustomDraw | deItemGetInfoTipText | deItemSelectionChangingEvents);
			dontRedraw = FALSE;
			dragExpandTime = -1;
			dragScrollTimeBase = -1;
			drawImagesAsynchronously = FALSE;
			editBackColor = 0x80000000 | COLOR_WINDOW;
			editForeColor = 0x80000000 | COLOR_WINDOWTEXT;
			editHoverTime = -1;
			editIMEMode = imeInherit;
			enabled = TRUE;
			fadeExpandos = FALSE;
			favoritesStyle = FALSE;
			foreColor = 0x80000000 | COLOR_WINDOWTEXT;
			fullRowSelect = FALSE;
			groupBoxColor = 0x80000000 | COLOR_BTNSHADOW;
			hotTracking = htrNone;
			hoverTime = -1;
			IMEMode = imeInherit;
			indent = 16;
			indentStateImages = TRUE;
			insertMarkColor = 0;
			itemBorders = MAKELONG(3, 0);
			itemBoundingBoxDefinition = static_cast<ItemBoundingBoxDefinitionConstants>(ibbdItemIcon | ibbdItemLabel | ibbdItemStateImage);
			itemHeight = 17;
			lineColor = 0x80000000 | COLOR_BTNSHADOW;
			lineStyle = lsLinesAtItem;
			locale = static_cast<LCID>(-1);
			maxScrollTime = 100;
			mousePointer = mpDefault;
			oleDragImageStyle = odistClassic;
			processContextMenuKeys = TRUE;
			registerForOLEDragDrop = FALSE;
			richToolTips = FALSE;
			rightToLeft = static_cast<RightToLeftConstants>(0);
			scrollBars = sbNormal;
			showStateImages = FALSE;
			showToolTips = TRUE;
			singleExpand = seNone;
			sortOrder = soAscending;
			supportOLEDragImages = TRUE;
			textParsingFlagsForCompareString = 0;
			textParsingFlagsForVarFromStr = 0;
			treeViewStyle = static_cast<TreeViewStyleConstants>(tvsExpandos | tvsLines);
			useSystemFont = TRUE;
		}
	} Properties;
	/// \brief <em>Holds the control's properties' settings</em>
	Properties properties;

	/// \brief <em>Holds some frequently used settings</em>
	///
	/// Holds some settings that otherwise would be requested from the control window very often.
	/// The cached settings are updated whenever the corresponding control window's settings change.
	struct CachedSettings
	{
		/// \brief <em>Holds the \c ilHighResolution imagelist</em>
		///
		/// \sa get_hImageList, put_hImageList
		HIMAGELIST hHighResImageList;
		/// \brief <em>Holds the \c ilItems imagelist</em>
		///
		/// \sa get_hImageList, put_hImageList, OnSetImageList
		HIMAGELIST hImageList;
		/// \brief <em>Holds the height of a icon of the imagelist defined by \c hImageList</em>
		///
		/// \sa hImageList, OnSetImageList
		int iconHeight;
		/// \brief <em>Holds the width of a icon of the imagelist defined by \c hImageList</em>
		///
		/// \sa hImageList, OnSetImageList
		int iconWidth;
		/// \brief <em>Holds the \c ilState imagelist</em>
		///
		/// \sa get_hImageList, put_hImageList, OnSetImageList
		HIMAGELIST hStateImageList;
		/// \brief <em>Holds the \c ItemHeight property's setting</em>
		///
		/// \sa get_ItemHeight, put_ItemHeight, OnSetItemHeight
		SHORT itemHeight;
		/// \brief <em>Holds the \c ItemXBorder property's setting</em>
		///
		/// \sa get_ItemXBorder, put_ItemXBorder, OnSetBorder
		WORD itemXBorder;
		/// \brief <em>Holds the \c ItemYBorder property's setting</em>
		///
		/// \sa get_ItemYBorder, put_ItemYBorder, OnSetBorder
		WORD itemYBorder;

		CachedSettings()
		{
			hHighResImageList = NULL;
			hImageList = NULL;
			iconHeight = 0;
			iconWidth = 0;
			hStateImageList = NULL;
			itemHeight = -1;
			itemXBorder = 3;
			itemYBorder = 0;
		}

		/// \brief <em>Fills the caches with the settings of the specified control window</em>
		///
		/// \param[in] hWndTvw The treeview window whose settings will be cached.
		void CacheSettings(HWND hWndTvw)
		{
			hHighResImageList = NULL;
			hImageList = reinterpret_cast<HIMAGELIST>(::SendMessage(hWndTvw, TVM_GETIMAGELIST, TVSIL_NORMAL, 0));
			ImageList_GetIconSize(hImageList, &iconWidth, &iconHeight);
			hStateImageList = reinterpret_cast<HIMAGELIST>(::SendMessage(hWndTvw, TVM_GETIMAGELIST, TVSIL_STATE, 0));
			itemHeight = static_cast<SHORT>(::SendMessage(hWndTvw, TVM_GETITEMHEIGHT, 0, 0));
			DWORD itemBorders = static_cast<DWORD>(::SendMessage(hWndTvw, TVM_GETBORDER, 0, 0));
			itemXBorder = LOWORD(itemBorders);
			itemYBorder = HIWORD(itemBorders);
		}
	} cachedSettings;

	/// \brief <em>Holds the control's flags</em>
	struct Flags
	{
		/// \brief <em>If not 0, we avoid items being single-expanded</em>
		///
		/// \sa EnterNoSingleExpandSection, LeaveNoSingleExpandSection, OnSingleExpandNotification
		int dontSingleExpand;
		/// \brief <em>If not 0, we won't raise \c Insert*Item events</em>
		///
		/// \sa EnterSilentItemInsertionSection, LeaveSilentItemInsertionSection, OnInsertItem,
		///     Raise_InsertingItem, Raise_InsertedItem
		int silentItemInsertions;
		/// \brief <em>If \c TRUE, we will avoid the info tip being displayed</em>
		///
		/// \sa OnGetInfoTipNotification, OnToolTipGetDispInfoNotificationW
		UINT cancelToolTip : 1;
		/// \brief <em>If \c TRUE, we currently are removing all items</em>
		///
		/// \sa OnDeleteItem, OnCaretChangingNotification, OnCaretChangedNotification
		UINT removingAllItems : 1;
		/// \brief <em>If \c TRUE, we won't raise the \c CustomDraw event</em>
		///
		/// When \c SysTreeView32 handles \c TVM_INSERTITEM, it will calculate the scroll position and
		/// therefore send a \c NM_CUSTOMDRAW notification for the newly inserted item. This is problematic,
		/// because the \c CustomDraw event handler may access the item's \c ItemData property. The value
		/// for this property is stored in a map, which is updated <strong>after</strong> \c TVM_INSERTITEM was processed
		/// by \c SysTreeView32.
		///
		/// \sa OnInsertItem, OnCustomDrawNotification, Raise_CustomDraw
		UINT noCustomDraw : 1;
		/// \brief <em>If \c TRUE, we're using themes</em>
		///
		/// \sa OnThemeChanged
		UINT usingThemes : 1;

		Flags()
		{
			silentItemInsertions = 0;
			dontSingleExpand = 0;
			cancelToolTip = FALSE;
			removingAllItems = FALSE;
			noCustomDraw = FALSE;
			usingThemes = FALSE;
		}
	} flags;

	/// \brief <em>Holds the settings of the current sorting process</em>
	///
	/// \sa SortSubItems, CompareItems
	struct SortingSettings
	{
		/// \brief <em>If \c TRUE, text comparisons should be case sensitive</em>
		UINT caseSensitive : 1;
		/// \brief <em>The five sorting criteria to use</em>
		///
		/// Holds the five sorting criteria to use. If the items are equivalent regarding the first one,
		/// the second will be used. If the result is equivalence again, the third one is used and so on.
		SortByConstants sortingCriteria[5];
		/// \brief <em>The locale identifier to use when comparing strings</em>
		///
		/// If set to -1, the current thread's locale identifier is used.
		LCID localeID;
		/// \brief <em>Holds the flags to specify when calling conversion routines, e.g. for converting strings to date time values</em>
		ULONG flagsForVarFromStr;
		/// \brief <em>Holds the flags to specify when calling \c CompareString</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/dd317759.aspx">CompareString</a>
		DWORD flagsForCompareString;

		SortingSettings()
		{
			caseSensitive = FALSE;
			sortingCriteria[0] = sobShell;
			sortingCriteria[1] = sobText;
			sortingCriteria[2] = sobNone;
			sortingCriteria[3] = sobNone;
			sortingCriteria[4] = sobNone;
			localeID = static_cast<LCID>(-1);
			flagsForVarFromStr = 0;
			flagsForCompareString = 0;
		}
	} sortingSettings;

	//////////////////////////////////////////////////////////////////////
	/// \name Item management
	///
	//@{
	/// \brief <em>Holds additional item data</em>
	typedef struct ItemData
	{
		/// \brief <em>Specifies an user-defined integer value associated with the item</em>
		LPARAM lParam;
	} ItemData;
	#ifdef USE_STL
		/// \brief <em>A map of all items</em>
		///
		/// Holds the handles of all treeview items in the control. The item's ID is stored as key, the
		/// item's handle is stored as value.
		///
		/// \sa OnInsertItem, OnDeleteItem, OnDeleteItemNotification, OnGetItem, OnSetItem,
		///     TreeViewItem::get_ID, ReplaceItemHandle
		std::unordered_map<LONG, HTREEITEM> itemHandles;
		/// \brief <em>A map of all items</em>
		///
		/// Holds additional data for each treeview item in the control. The item's handle is stored as key,
		/// the item's additional data is stored as value.
		///
		/// \sa OnInsertItem, OnDeleteItem, OnDeleteItemNotification, OnGetItem, OnSetItem,
		///     TreeViewItem::get_ID, ReplaceItemHandle
		std::unordered_map<HTREEITEM, ItemData> itemParams;
	#else
		/// \brief <em>A map of all items</em>
		///
		/// Holds the handles of all treeview items in the control. The item's ID is stored as key, the
		/// item's handle is stored as value.
		///
		/// \sa OnInsertItem, OnDeleteItem, OnDeleteItemNotification, OnGetItem, OnSetItem,
		///     TreeViewItem::get_ID, ReplaceItemHandle
		CAtlMap<LONG, HTREEITEM> itemHandles;
		/// \brief <em>A map of all items</em>
		///
		/// Holds additional data for each treeview item in the control. The item's handle is stored as key,
		/// the item's additional data is stored as value.
		///
		/// \sa OnInsertItem, OnDeleteItem, OnDeleteItemNotification, OnGetItem, OnSetItem,
		///     TreeViewItem::get_ID, ReplaceItemHandle
		CAtlMap<HTREEITEM, ItemData> itemParams;
	#endif
	/// \brief <em>Retrieves a new unique item ID at each call</em>
	///
	/// \return A new unique item ID.
	///
	/// \sa itemHandles, TreeViewItem::get_ID
	LONG GetNewItemID(void);
	#ifdef USE_STL
		/// \brief <em>A vector of items that we won't raise \c Remov*Item events for on deletion</em>
		///
		/// Holds the handles of items of whose deletion we won't notify the client. We use this vector for
		/// the \c TreeViewItem::Move feature.
		///
		/// \sa OnDeleteItem, OnDeleteItemNotification, AddSilentItemDeletion, TreeViewItem::Move
		std::vector<HTREEITEM> itemDeletionsToIgnore;
		/// \brief <em>A map of all \c TreeViewItemContainer objects that we've created</em>
		///
		/// Holds pointers to all \c TreeViewItemContainer objects that we've created. We use this map to
		/// inform the containers of removed items. The container's ID is stored as key; the container's
		/// \c IItemContainer implementation is stored as value.
		///
		/// \sa CreateItemContainer, RegisterItemContainer, TreeViewItemContainer
		std::unordered_map<DWORD, IItemContainer*> itemContainers;
	#else
		/// \brief <em>A vector of items that we won't raise \c Remov*Item events for on deletion</em>
		///
		/// Holds the handles of items of whose deletion we won't notify the client. We use this vector for
		/// the \c TreeViewItem::Move feature.
		///
		/// \sa OnDeleteItem, OnDeleteItemNotification, AddSilentItemDeletion, TreeViewItem::Move
		CAtlArray<HTREEITEM> itemDeletionsToIgnore;
		/// \brief <em>A map of all \c TreeViewItemContainer objects that we've created</em>
		///
		/// Holds pointers to all \c TreeViewItemContainer objects that we've created. We use this map to
		/// inform the containers of removed items. The container's ID is stored as key; the container's
		/// \c IItemContainer implementation is stored as value.
		///
		/// \sa CreateItemContainer, RegisterItemContainer, TreeViewItemContainer
		CAtlMap<DWORD, IItemContainer*> itemContainers;
	#endif
	/// \brief <em>Registers the specified \c TreeViewItemContainer collection</em>
	///
	/// Registers the specified \c TreeViewItemContainer collection so that it is informed of item deletions.
	///
	/// \param[in] pContainer The container's \c IItemContainer implementation.
	///
	/// \sa DeregisterItemContainer, itemContainers, RemoveItemFromItemContainers
	void RegisterItemContainer(IItemContainer* pContainer);
	/// \brief <em>De-registers the specified \c TreeViewItemContainer collection</em>
	///
	/// De-registers the specified \c TreeViewItemContainer collection so that it no longer is informed of
	/// item deletions.
	///
	/// \param[in] containerID The container's ID.
	///
	/// \sa RegisterItemContainer, itemContainers
	void DeregisterItemContainer(DWORD containerID);
	/// \brief <em>Removes the specified item from all registered \c TreeViewItemContainer collections</em>
	///
	/// \param[in] itemID The unique ID of the item to remove. If 0, all items are removed.
	///
	/// \sa itemContainers, RegisterItemContainer, OnDeleteItem, OnDeleteItemNotification,
	///     Raise_DestroyedControlWindow
	void RemoveItemFromItemContainers(LONG itemID);
	/// \brief <em>Validates an item handle</em>
	///
	/// \param[in] itemHandle The item handle to validate.
	///
	/// \return If \c TRUE, the handle identifies an item and is valid; otherwise it is invalid.
	///
	/// \sa OnDeleteItem
	BOOL ValidateItemHandle(HTREEITEM itemHandle);
	///@}
	//////////////////////////////////////////////////////////////////////


	#ifdef USE_STL
		/// \brief <em>A map of internal caret changes</em>
		///
		/// Holds the handles of items that we called \c TVM_SELECTITEM/TVGN_CARET for. The item handle
		/// is stored as key; the caret change's reason, that we should tell the client, is stored as value.
		///
		/// \sa AddInternalCaretChange, RemoveInternalCaretChange
		std::unordered_map<HTREEITEM, UINT> internalCaretChanges;
	#else
		/// \brief <em>A map of internal caret changes</em>
		///
		/// Holds the handles of items that we called \c TVM_SELECTITEM/TVGN_CARET for. The item handle
		/// is stored as key; the caret change's reason, that we should tell the client, is stored as value.
		///
		/// \sa AddInternalCaretChange, RemoveInternalCaretChange
		CAtlMap<HTREEITEM, UINT> internalCaretChanges;
	#endif
	/// \brief <em>Manages tasks that are executed on caret changes</em>
	///
	/// \sa AddCaretChangeTask, CaretChangeTasks
	CaretChangeTasks caretChangeTasks;
	/// \brief <em>Holds the handle of the item below the mouse cursor</em>
	///
	/// \attention This member is not reliable with \c deTreeMouseEvents being set.
	HTREEITEM hItemUnderMouse;

	/// \brief <em>Holds mouse status variables</em>
	typedef struct MouseStatus
	{
	protected:
		/// \brief <em>Holds all mouse buttons that may cause a click event in the immediate future</em>
		///
		/// A bit field of \c SHORT values representing those mouse buttons that are currently pressed and
		/// may cause a click event in the immediate future.
		///
		/// \sa StoreClickCandidate, IsClickCandidate, RemoveClickCandidate, Raise_Click, Raise_MClick,
		///     Raise_RClick, Raise_XClick
		SHORT clickCandidates;

	public:
		/// \brief <em>If \c TRUE, the \c MouseEnter event was already raised</em>
		///
		/// \attention This member is not reliable with \c deEditMouseEvents being set.
		///
		/// \sa Raise_MouseEnter, Raise_EditMouseEnter
		UINT enteredControl : 1;
		/// \brief <em>If \c TRUE, the \c MouseHover event was already raised</em>
		///
		/// \attention This member is not reliable with \c deTreeMouseEvents respectively
		///            \c deEditMouseEvents being set.
		///
		/// \sa Raise_MouseHover, Raise_EditMouseHover
		UINT hoveredControl : 1;
		/// \brief <em>Holds the mouse cursor's last position</em>
		///
		/// \attention This member is not reliable with \c deTreeMouseEvents respectively
		///            \c deEditMouseEvents set.
		POINT lastPosition;
		/// \brief <em>If \c TRUE, \c OnTree*ButtonUp should raise the \c *DblClick event</em>
		///
		/// \sa OnLButtonDblClk, OnDblClkNotification, OnLButtonUp, OnRButtonDblClk, OnRDblClkNotification,
		///     OnRButtonUp, Raise_DblClick, Raise_RDblClick
		UINT raiseDblClkEvent : 1;
		/// \brief <em>Holds the handle of the last clicked item</em>
		///
		/// Holds the handle of the last clicked item. We use this to ensure that the \c *DblClick events
		/// are not raised accidently.
		///
		/// \attention This member is not reliable with \c deTreeClickEvents being set.
		///
		/// \sa Raise_Click, Raise_DblClick, Raise_MClick, Raise_MDblClick, Raise_RClick, Raise_RDblClick,
		///     Raise_XClick, Raise_XDblClick
		HTREEITEM hLastClickedItem;

		MouseStatus()
		{
			clickCandidates = 0;
			enteredControl = FALSE;
			hoveredControl = FALSE;
			raiseDblClkEvent = FALSE;
			hLastClickedItem = NULL;
		}

		/// \brief <em>Changes flags to indicate the \c MouseEnter event was just raised</em>
		///
		/// \sa enteredControl, HoverControl, LeaveControl
		void EnterControl(void)
		{
			RemoveAllClickCandidates();
			enteredControl = TRUE;
			raiseDblClkEvent = FALSE;
			hLastClickedItem = NULL;
		}

		/// \brief <em>Changes flags to indicate the \c MouseHover event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl, LeaveControl
		void HoverControl(void)
		{
			enteredControl = TRUE;
			hoveredControl = TRUE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseLeave event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl
		void LeaveControl(void)
		{
			enteredControl = FALSE;
			hoveredControl = FALSE;
			raiseDblClkEvent = FALSE;
			hLastClickedItem = NULL;
		}

		/// \brief <em>Stores a mouse button as click candidate</em>
		///
		/// param[in] button The mouse button to store.
		///
		/// \sa clickCandidates, IsClickCandidate, RemoveClickCandidate
		void StoreClickCandidate(SHORT button)
		{
			// avoid combined click events
			if(clickCandidates == 0) {
				clickCandidates |= button;
			}
		}

		/// \brief <em>Retrieves whether a mouse button is a click candidate</em>
		///
		/// \param[in] button The mouse button to check.
		///
		/// \return \c TRUE if the button is stored as a click candidate; otherwise \c FALSE.
		///
		/// \attention This member is not reliable with \c deTreeMouseEvents respectively
		///            \c deEditMouseEvents being set.
		///
		/// \sa clickCandidates, StoreClickCandidate, RemoveClickCandidate
		BOOL IsClickCandidate(SHORT button)
		{
			return (clickCandidates & button);
		}

		/// \brief <em>Removes a mouse button from the list of click candidates</em>
		///
		/// \param[in] button The mouse button to remove.
		///
		/// \sa clickCandidates, RemoveAllClickCandidates, StoreClickCandidate, IsClickCandidate
		void RemoveClickCandidate(SHORT button)
		{
			clickCandidates &= ~button;
		}

		/// \brief <em>Clears the list of click candidates</em>
		///
		/// \sa clickCandidates, RemoveClickCandidate, StoreClickCandidate, IsClickCandidate
		void RemoveAllClickCandidates(void)
		{
			clickCandidates = 0;
		}
	} MouseStatus;

	/// \brief <em>Holds the control's mouse status</em>
	///
	/// \sa mouseStatus_Edit
	MouseStatus mouseStatus;
	/// \brief <em>Holds the label-edit control's mouse status</em>
	///
	/// \sa mouseStatus
	MouseStatus	mouseStatus_Edit;

	/// \brief <em>Holds data and flags related to the \c CaretChangedDelayTime feature</em>
	///
	/// \sa get_CaretChangedDelayTime
	struct CaretChangedDelayStatus
	{
		/// \brief <em>Used to indicated that a point of time was not specified</em>
		static const DWORD UNDEFINEDTIME = static_cast<DWORD>(-1);

		/// \brief <em>The point of time of the caret change</em>
		DWORD timeOfCaretChange;
		/// \brief <em>The current caret item's handle</em>
		HTREEITEM hNewCaretItem;
		/// \brief <em>The previous caret item's handle</em>
		HTREEITEM hPreviousCaretItem;

		CaretChangedDelayStatus()
		{
			timeOfCaretChange = UNDEFINEDTIME;
			hNewCaretItem = NULL;
			hPreviousCaretItem = NULL;
		}

		/// \brief <em>Initializes this struct with the given values</em>
		///
		/// \param[in] hNewCaretItem The item that became the new caret item in this caret change.
		/// \param[in] hPreviousCaretItem The item that was the caret item before this caret change.
		///
		/// \sa DelayTimePassed
		void Initialize(HTREEITEM hNewCaretItem, HTREEITEM hPreviousCaretItem)
		{
			this->hNewCaretItem = hNewCaretItem;
			if(timeOfCaretChange == UNDEFINEDTIME) {
				this->hPreviousCaretItem = hPreviousCaretItem;
			}
			timeOfCaretChange = GetMessageTime();
		}

		/// \brief <em>Retrieves whether the given period of time has past</em>
		///
		/// \param[in] delayTime The number of milliseconds that must have passed since the point of time
		///            identified by \c timeOfCaretChange.
		///
		/// \return \c TRUE if at least <\c delayTime> milliseconds passed since the point of time
		///         identified by \c timeOfCaretChange; otherwise \c FALSE.
		///
		/// \sa timeOfCaretChange, ResetTimeOfCaretChange
		BOOL DelayTimePassed(DWORD delayTime)
		{
			DWORD now = GetTickCount();
			if(now < timeOfCaretChange) {
				// the counter hit the DWORD bound and started from 0 again
				if(MAXDWORD - timeOfCaretChange + now >= delayTime) {
					return TRUE;
				}
			} else if(now - timeOfCaretChange >= delayTime) {
				return TRUE;
			}
			return FALSE;
		}

		/// \brief <em>Resets \c timeOfCaretChange to its initial state</em>
		///
		/// \sa timeOfCaretChange
		void ResetTimeOfCaretChange(void)
		{
			timeOfCaretChange = UNDEFINEDTIME;
		}
	} caretChangedDelayStatus;

	/// \brief <em>Holds the position of the control's insertion mark</em>
	struct InsertMark
	{
		/// \brief <em>The handle of the item at which the insertion mark is placed</em>
		HTREEITEM hItem;
		/// \brief <em>The insertion mark's position relative to the item</em>
		BOOL afterItem;

		InsertMark()
		{
			Reset();
		}

		/// \brief <em>Resets all member variables</em>
		void Reset(void)
		{
			hItem = NULL;
			afterItem = FALSE;
		}

		/// \brief <em>Processes the parameters of a \c TVM_SETINSERTMARK message to store the new insertion mark</em>
		///
		/// \sa OnSetInsertMark
		void ProcessSetInsertMark(WPARAM wParam, LPARAM lParam)
		{
			afterItem = static_cast<BOOL>(wParam);
			hItem = reinterpret_cast<HTREEITEM>(lParam);
		}
	} insertMark;

	/// \brief <em>Holds data and flags used with state image changes</em>
	///
	/// \sa OnLButtonDown, OnLButtonUp
	struct StateImageChangeFlags
	{
		/// \brief <em>The handle of the item whose state image is affected in a state image change using the mouse</em>
		HTREEITEM hClickedItem;
		/// \brief <em>The handle of the item whose state image is currently changed through \c TVM_SETITEM</em>
		///
		/// \sa OnSetItem, OnItemChangedNotification, OnItemChangingNotification
		HTREEITEM hItemCurrentInternalStateImageChange;
		/// \brief <em>The handle of the item whose state image shall be changed by \c OnItemChangedNotification or \c OnTimer</em>
		///
		/// \sa stateImageIndexToResetTo, OnItemChangedNotification, OnItemChangingNotification, OnTimer
		HTREEITEM hItemToResetStateImageFor;
		/// \brief <em>The zero-based index of the state image that \c hItemToResetStateImageFor will be set to</em>
		///
		/// \sa hItemToResetStateImageFor, OnItemChangedNotification, OnItemChangingNotification, OnTimer
		int stateImageIndexToResetTo;
		/// \brief <em>If \c TRUE, the state image change has been canceled in \c OnLButtonUp and \c OnTimer shouldn't raise the \c ItemStateImageChanged event</em>
		///
		/// \sa OnLButtonUp, OnTimer, Raise_ItemStateImageChanged
		UINT changeHasBeenCanceled : 1;
		/// \brief <em>The zero-based index of the previous state image of \c hItemToResetStateImageFor</em>
		///
		/// \sa hItemToResetStateImageFor, OnLButtonUp, OnTimer
		int previousStateImage;
		/// \brief <em>If not 0, we won't raise the \c ItemStateImageChanging and \c ItemStateImageChanged events</em>
		///
		/// \sa EnterNoStateImageChangeEventsSection, LeaveNoStateImageChangeEventsSection,
		///     Raise_ItemStateImageChanging, Raise_ItemStateImageChanged
		int silentStateImageChanges;
		/// \brief <em>Holds a value that is used by some notifications to report the client the correct cause for the state image change</em>
		///
		/// \remarks This member is useful with comctl32.dll version 6.10 or higher only.
		///
		/// \sa OnTVStateImageChangingNotification, OnItemChangedNotification, OnItemChangingNotification,
		///     OnKeyDown, OnLButtonDown, OnLButtonUp
		StateImageChangeCausedByConstants reasonForPotentialStateImageChange;

		StateImageChangeFlags()
		{
			Reset();
		}

		/// \brief <em>Increments the \c silentStateImageChanges flag</em>
		///
		/// \sa LeaveNoStateImageChangeEventsSection, silentStateImageChanges, Raise_ItemStateImageChanging,
		///     Raise_ItemStateImageChanged
		void EnterNoStateImageChangeEventsSection(void)
		{
			++silentStateImageChanges;
		}

		/// \brief <em>Decrements the \c silentStateImageChanges flag</em>
		///
		/// \sa EnterNoStateImageChangeEventsSection, silentStateImageChanges, Raise_ItemStateImageChanging,
		///     Raise_ItemStateImageChanged
		void LeaveNoStateImageChangeEventsSection(void)
		{
			--silentStateImageChanges;
		}

		/// \brief <em>Resets all member variables</em>
		void Reset(void)
		{
			hClickedItem = NULL;
			hItemCurrentInternalStateImageChange = NULL;
			hItemToResetStateImageFor = NULL;
			silentStateImageChanges = 0;
			reasonForPotentialStateImageChange = static_cast<StateImageChangeCausedByConstants>(-1);
		}
	} stateImageChangeFlags;

	//////////////////////////////////////////////////////////////////////
	/// \name Drag'n'Drop
	///
	//@{
	/// \brief <em>The \c CLSID_WICImagingFactory object used to create WIC objects that are required during drag image creation</em>
	///
	/// \sa OnGetDragImage, CreateThumbnail
	CComPtr<IWICImagingFactory> pWICImagingFactory;
	/// \brief <em>Creates a thumbnail of the specified icon in the specified size</em>
	///
	/// \param[in] hIcon The icon to create the thumbnail for.
	/// \param[in] size The thumbnail's size in pixels.
	/// \param[in,out] pBits The thumbnail's DIB bits.
	/// \param[in] doAlphaChannelPostProcessing WIC has problems to handle the alpha channel of the icon
	///            specified by \c hIcon. If this parameter is set to \c TRUE, some post-processing is done
	///            to correct the pixel failures. Otherwise the failures are not corrected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OnGetDragImage, pWICImagingFactory
	HRESULT CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing);

	/// \brief <em>Holds data and flags related to drag'n'drop</em>
	struct DragDropStatus
	{
		/// \brief <em>The \c ITreeViewItemContainer implementation of the collection of the dragged items</em>
		ITreeViewItemContainer* pDraggedItems;
		/// \brief <em>The handle of the imagelist containing the drag image</em>
		///
		/// \sa get_hDragImageList
		HIMAGELIST hDragImageList;
		/// \brief <em>Enables or disables auto-destruction of \c hDragImageList</em>
		///
		/// Controls whether the imagelist defined by \c hDragImageList is auto-destroyed. If set to
		/// \c TRUE, it is destroyed in \c EndDrag; otherwise not.
		///
		/// \sa hDragImageList, EndDrag
		UINT autoDestroyImgLst : 1;
		/// \brief <em>Indicates whether the drag image is visible or hidden</em>
		///
		/// If this value is 0, the drag image is visible; otherwise not.
		///
		/// \sa get_hDragImageList, get_ShowDragImage, put_ShowDragImage, ShowDragImage, HideDragImage,
		///     IsDragImageVisible
		int dragImageIsHidden;
		/// \brief <em>The handle of the last drop target</em>
		HTREEITEM hLastDropTarget;

		//////////////////////////////////////////////////////////////////////
		/// \name OLE Drag'n'Drop
		///
		//@{
		/// \brief <em>The currently dragged data</em>
		CComPtr<IOLEDataObject> pActiveDataObject;
		/// \brief <em>The currently dragged data for the case that the we're the drag source</em>
		CComPtr<IDataObject> pSourceDataObject;
		/// \brief <em>Holds the mouse cursors last position</em>
		POINTL lastMousePosition;
		/// \brief <em>The \c IDropTargetHelper object used for drag image support</em>
		///
		/// \sa put_SupportOLEDragImages,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
		IDropTargetHelper* pDropTargetHelper;
		/// \brief <em>Holds the mouse button (as \c MK_* constant) that the drag'n'drop operation is performed with</em>
		DWORD draggingMouseButton;
		/// \brief <em>If \c TRUE, we'll hide and re-show the drag image in \c IDropTarget::DragEnter so that the item count label is displayed</em>
		///
		/// \sa DragEnter, OLEDrag
		UINT useItemCountLabelHack : 1;
		/// \brief <em>Holds the \c IDataObject to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		IDataObject* drop_pDataObject;
		/// \brief <em>Holds the mouse position to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		POINT drop_mousePosition;
		/// \brief <em>Holds the drop effect to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		DWORD drop_effect;
		//@}
		//////////////////////////////////////////////////////////////////////

		/// \brief <em>Holds data and flags related to auto-scrolling</em>
		///
		/// \sa AutoScroll
		struct AutoScrolling
		{
			/// \brief <em>Holds the current speed multiplier used for horizontal auto-scrolling</em>
			LONG currentHScrollVelocity;
			/// \brief <em>Holds the current speed multiplier used for vertical auto-scrolling</em>
			LONG currentVScrollVelocity;
			/// \brief <em>Holds the current interval of the auto-scroll timer</em>
			LONG currentTimerInterval;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled downwards</em>
			DWORD lastScroll_Down;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled to the left</em>
			DWORD lastScroll_Left;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled to the right</em>
			DWORD lastScroll_Right;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled upwardly</em>
			DWORD lastScroll_Up;

			AutoScrolling()
			{
				Reset();
			}

			/// \brief <em>Resets all member variables to their defaults</em>
			void Reset(void)
			{
				currentHScrollVelocity = 0;
				currentVScrollVelocity = 0;
				currentTimerInterval = 0;
				lastScroll_Down = 0;
				lastScroll_Left = 0;
				lastScroll_Right = 0;
				lastScroll_Up = 0;
			}
		} autoScrolling;

		DragDropStatus()
		{
			pActiveDataObject = NULL;
			pSourceDataObject = NULL;
			pDropTargetHelper = NULL;
			draggingMouseButton = 0;
			useItemCountLabelHack = FALSE;
			drop_pDataObject = NULL;

			pDraggedItems = NULL;
			hDragImageList = NULL;
			autoDestroyImgLst = FALSE;
			dragImageIsHidden = 1;
			hLastDropTarget = NULL;
		}

		~DragDropStatus()
		{
			if(pDropTargetHelper) {
				pDropTargetHelper->Release();
			}
			ATLASSERT(!pDraggedItems);
		}

		/// \brief <em>Resets all member variables to their defaults</em>
		void Reset(void)
		{
			if(this->pDraggedItems) {
				this->pDraggedItems->Release();
				this->pDraggedItems = NULL;
			}
			if(hDragImageList && autoDestroyImgLst) {
				ImageList_Destroy(hDragImageList);
			}
			hDragImageList = NULL;
			autoDestroyImgLst = FALSE;
			dragImageIsHidden = 1;
			hLastDropTarget = NULL;

			if(this->pActiveDataObject) {
				this->pActiveDataObject = NULL;
			}
			if(this->pSourceDataObject) {
				this->pSourceDataObject = NULL;
			}
			draggingMouseButton = 0;
			useItemCountLabelHack = FALSE;
			drop_pDataObject = NULL;
		}

		/// \brief <em>Decrements the \c dragImageIsHidden flag</em>
		///
		/// \param[in] commonDragDropOnly If \c TRUE, the method does nothing if we're within an OLE
		///            drag'n'drop operation.
		///
		/// \sa dragImageIsHidden, HideDragImage, IsDragImageVisible
		void ShowDragImage(BOOL commonDragDropOnly)
		{
			if(hDragImageList) {
				--dragImageIsHidden;
				if(dragImageIsHidden == 0) {
					ImageList_DragShowNolock(TRUE);
				}
			} else if(pDropTargetHelper && !commonDragDropOnly) {
				--dragImageIsHidden;
				if(dragImageIsHidden == 0) {
					pDropTargetHelper->Show(TRUE);
				}
			}
		}

		/// \brief <em>Increments the \c dragImageIsHidden flag</em>
		///
		/// \param[in] commonDragDropOnly If \c TRUE, the method does nothing if we're within an OLE
		///            drag'n'drop operation.
		///
		/// \sa dragImageIsHidden, ShowDragImage, IsDragImageVisible
		void HideDragImage(BOOL commonDragDropOnly)
		{
			if(hDragImageList) {
				++dragImageIsHidden;
				if(dragImageIsHidden == 1) {
					ImageList_DragShowNolock(FALSE);
				}
			} else if(pDropTargetHelper && !commonDragDropOnly) {
				++dragImageIsHidden;
				if(dragImageIsHidden == 1) {
					pDropTargetHelper->Show(FALSE);
				}
			}
		}

		/// \brief <em>Retrieves whether we're currently displaying a drag image</em>
		///
		/// \return \c TRUE, if we're displaying a drag image; otherwise \c FALSE.
		///
		/// \sa dragImageIsHidden, ShowDragImage, HideDragImage
		BOOL IsDragImageVisible(void)
		{
			return (dragImageIsHidden == 0);
		}

		/// \brief <em>Performs any tasks that must be done after a drag'n'drop operation started</em>
		///
		/// \param[in] hWndTvw The treeview window, that the method will work on to calculate the position
		///            of the drag image's hotspot.
		/// \param[in] pDraggedItems The \c ITreeViewItemContainer implementation of the collection of
		///            the dragged items.
		/// \param[in] hDragImageList The imagelist containing the drag image that shall be used to
		///            visualize the drag'n'drop operation. If -1, the method will create the drag image
		///            itself; if \c NULL, no drag image will be displayed.
		/// \param[in,out] pXHotSpot The x-coordinate (in pixels) of the drag image's hotspot relative to the
		///                drag image's upper-left corner. If the \c hDragImageList parameter is set to
		///                \c NULL, this parameter is ignored. If the \c hDragImageList parameter is set to
		///                -1, this parameter is set to the hotspot calculated by the method.
		/// \param[in,out] pYHotSpot The y-coordinate (in pixels) of the drag image's hotspot relative to the
		///                drag image's upper-left corner. If the \c hDragImageList parameter is set to
		///                \c NULL, this parameter is ignored. If the \c hDragImageList parameter is set to
		///                -1, this parameter is set to the hotspot calculated by the method.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa EndDrag
		HRESULT BeginDrag(HWND hWndTvw, ITreeViewItemContainer* pDraggedItems, HIMAGELIST hDragImageList, PINT pXHotSpot, PINT pYHotSpot)
		{
			ATLASSUME(pDraggedItems);
			if(!pDraggedItems) {
				return E_INVALIDARG;
			}

			UINT b = FALSE;
			if(hDragImageList == static_cast<HIMAGELIST>(LongToHandle(-1))) {
				OLE_HANDLE h = NULL;
				OLE_XPOS_PIXELS xUpperLeft = 0;
				OLE_YPOS_PIXELS yUpperLeft = 0;
				if(FAILED(pDraggedItems->CreateDragImage(&xUpperLeft, &yUpperLeft, &h))) {
					return E_FAIL;
				}
				hDragImageList = static_cast<HIMAGELIST>(LongToHandle(h));
				b = TRUE;

				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				::ScreenToClient(hWndTvw, &mousePosition);
				if(CWindow(hWndTvw).GetExStyle() & WS_EX_LAYOUTRTL) {
					SIZE dragImageSize = {0};
					ImageList_GetIconSize(hDragImageList, reinterpret_cast<PINT>(&dragImageSize.cx), reinterpret_cast<PINT>(&dragImageSize.cy));
					*pXHotSpot = xUpperLeft + dragImageSize.cx - mousePosition.x;
				} else {
					*pXHotSpot = mousePosition.x - xUpperLeft;
				}
				*pYHotSpot = mousePosition.y - yUpperLeft;
			}

			if(this->hDragImageList && this->autoDestroyImgLst) {
				ImageList_Destroy(this->hDragImageList);
			}

			this->autoDestroyImgLst = b;
			this->hDragImageList = hDragImageList;
			if(this->pDraggedItems) {
				this->pDraggedItems->Release();
				this->pDraggedItems = NULL;
			}
			pDraggedItems->Clone(&this->pDraggedItems);
			ATLASSUME(this->pDraggedItems);
			this->hLastDropTarget = NULL;

			dragImageIsHidden = 1;
			autoScrolling.Reset();
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done after a drag'n'drop operation ended</em>
		///
		/// \sa BeginDrag
		void EndDrag(void)
		{
			if(pDraggedItems) {
				pDraggedItems->Release();
				pDraggedItems = NULL;
			}
			if(autoDestroyImgLst && hDragImageList) {
				ImageList_Destroy(hDragImageList);
			}
			hDragImageList = NULL;
			dragImageIsHidden = 1;
			hLastDropTarget = NULL;
			autoScrolling.Reset();
		}

		/// \brief <em>Retrieves whether we're in drag'n'drop mode</em>
		///
		/// \return \c TRUE if we're in drag'n'drop mode; otherwise \c FALSE.
		///
		/// \sa BeginDrag, EndDrag
		BOOL IsDragging(void)
		{
			return (pDraggedItems != NULL);
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa OLEDragLeaveOrDrop
		HRESULT OLEDragEnter(void)
		{
			this->hLastDropTarget = NULL;
			autoScrolling.Reset();
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa OLEDragEnter
		void OLEDragLeaveOrDrop(void)
		{
			this->hLastDropTarget = NULL;
			autoScrolling.Reset();
		}
	} dragDropStatus;
	///@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds data and flags related to label-editing</em>
	struct LabelEditStatus
	{
		/// \brief <em>The handle of the item that currently is edited</em>
		HTREEITEM hEditedItem;
		/// \brief <em>The edited item's text that it had before label-editing started</em>
		BSTR previousText;
		/// \brief <em>If \c FALSE, the control rejects the new item text and re-enters label-edit mode immediately</em>
		///
		/// \sa OnEndLabelEditNotification, Raise_DestroyedEditControlWindow
		BOOL acceptText;
		#ifdef INCLUDESHELLBROWSERINTERFACE
			/// \brief <em>A text with which the label-edit control should be initialized</em>
			TCHAR pTextToSetOnBegin[MAX_ITEMTEXTLENGTH + 1];
		#endif
		/// \brief <em>If \c TRUE, we're awaiting the control to enter label-edit mode</em>
		///
		/// \sa get_AllowLabelEditing, get_SingleExpand, OnBeginLabelEditNotification,
		///     OnClickNotification, OnDblClkNotification, OnTimer
		UINT watchingForEdit : 1;

		LabelEditStatus()
		{
			Reset();
			#ifdef INCLUDESHELLBROWSERINTERFACE
				pTextToSetOnBegin[0] = TEXT('\0');
			#endif
		}

		/// \brief <em>Resets all members to their defaults</em>
		///
		/// \param[in] resetWatchingFlag If \c TRUE, the \c watchingForEdit member is reset; otherwise not.
		void Reset(BOOL resetWatchingFlag = TRUE)
		{
			hEditedItem = NULL;
			previousText = NULL;
			acceptText = TRUE;
			if(resetWatchingFlag) {
				watchingForEdit = FALSE;
			}
		}
	} labelEditStatus;

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		/// \brief <em>The ID of the timer that is used to delay caret change events</em>
		static const UINT_PTR ID_DELAYEDCARETCHANGE = 11;
		/// \brief <em>The ID of the timer that is used to redraw the control window after recreation</em>
		static const UINT_PTR ID_REDRAW = 12;
		/// \brief <em>The ID of the timer that is used to auto-scroll the control window during drag'n'drop</em>
		static const UINT_PTR ID_DRAGSCROLL = 13;
		/// \brief <em>The ID of the timer that is used to auto-expand items during drag'n'drop</em>
		static const UINT_PTR ID_DRAGEXPAND = 14;
		/// \brief <em>The ID of the timer that is used to start label-editing if \c SingleExpand is activated</em>
		///
		/// \sa OnClickNotification
		static const UINT_PTR ID_EDITLABEL = 15;
		/// \brief <em>The ID of the timer that helps \c OnLButtonUp to set an item's state-image to 0</em>
		///
		/// \sa OnLButtonUp
		static const UINT_PTR ID_STATEIMAGECHANGEHELPER = 16;

		/// \brief <em>The interval of the timer that is used to redraw the control window after recreation</em>
		static const UINT INT_REDRAW = 10;
		/// \brief <em>The interval of the timer that helps \c OnLButtonUp to set an item's state-image to 0</em>
		static const UINT INT_STATEIMAGECHANGEHELPER = 1;
	} timers;

	/// \brief <em>The handle of the control's built in state imagelist</em>
	HIMAGELIST hBuiltInStateImageList;
	/// \brief <em>The handle of the brush that the label-edit control's background is drawn with</em>
	HBRUSH hEditBackColorBrush;

	//////////////////////////////////////////////////////////////////////
	/// \name Version information
	///
	//@{
	/// \brief <em>Retrieves whether we're using at least version 6.10 of comctl32.dll</em>
	///
	/// \return \c TRUE if we're using comctl32.dll version 6.10 or higher; otherwise \c FALSE.
	BOOL IsComctl32Version610OrNewer(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	#ifdef INCLUDESHELLBROWSERINTERFACE
		//////////////////////////////////////////////////////////////////////
		/// \name ShellBrowser interface
		///
		//@{
		/// \brief <em>Holds any data that is specific to the \c ShellBrowser interface</em>
		struct ShellBrowserInterface
		{
			/// \brief <em>Holds the \c IMessageListener implementation of \c ShellBrowser</em>
			IMessageListener* pMessageListener;
			/// \brief <em>Holds the \c IInternalMessageListener implementation of \c ShellBrowser</em>
			IInternalMessageListener* pInternalMessageListener;

			ShellBrowserInterface()
			{
				pMessageListener = NULL;
				pInternalMessageListener = NULL;
			}
		} shellBrowserInterface;
		//@}
		//////////////////////////////////////////////////////////////////////
	#endif

private:
	/// \brief <em>Holds the \c LPARAM parameter of the control's last \c WM_KEYDOWN message</em>
	///
	/// \attention This member is not reliable with \c deTreeKeyboardEvents being set.
	///
	/// \sa OnKeyDown, Raise_IncrementalSearchStringChanging
	LPARAM lastKeyDownLParam;
};     // ExplorerTreeView

OBJECT_ENTRY_AUTO(__uuidof(ExplorerTreeView), ExplorerTreeView)