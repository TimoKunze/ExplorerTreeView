// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////
#pragma once
#include <initguid.h>

#import <libid:1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5> version("2.6") raw_dispinterfaces
#import <libid:0725EC44-6E15-4e66-A561-A94FDA7E5DCB> version("2.6") raw_dispinterfaces

DEFINE_GUID(LIBID_ExTVwLibU, 0x1f8f0fe7, 0x2cfb, 0x4466, 0xa2, 0xbc, 0xab, 0xb4, 0x41, 0xad, 0xed, 0xd5);
DEFINE_GUID(LIBID_ExTVwLibA, 0x0725ec44, 0x6e15, 0x4e66, 0xa5, 0x61, 0xa9, 0x4f, 0xda, 0x7e, 0x5d, 0xcb);

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_EXTVWU, CMainDlg, &__uuidof(ExTVwLibU::_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 6>,
    public IDispEventImpl<IDC_EXTVWA, CMainDlg, &__uuidof(ExTVwLibA::_IExplorerTreeViewEvents), &LIBID_ExTVwLibA, 2, 6>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> extvwUContainerWnd;
	CContainedWindowT<CWindow> extvwUWnd;
	CContainedWindowT<CAxWindow> extvwAContainerWnd;
	CContainedWindowT<CWindow> extvwAWnd;

	BOOL extvwaIsFocused;

	CMainDlg() :
	    extvwUContainerWnd(this, 1),
	    extvwAContainerWnd(this, 2),
	    extvwUWnd(this, 3),
	    extvwAWnd(this, 4)
	{
	}

	struct Controls
	{
		CEdit logEdit;
		CImageList imageList;
		CComPtr<ExTVwLibU::IExplorerTreeView> extvwU;
		CComPtr<ExTVwLibA::IExplorerTreeView> extvwA;

		~Controls()
		{
			if(!imageList.IsNull()) {
				imageList.Destroy();
			}
		}
	} controls;

	BOOL IsComctl32Version600OrNewer(void);
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		ALT_MSG_MAP(2)
		ALT_MSG_MAP(3)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocusU)

		ALT_MSG_MAP(4)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocusA)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 1, AbortedDragExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 0, CaretChangedExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 4, CaretChangingExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 5, ClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 6, CollapsedItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 7, CollapsingItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 8, CompareItemsExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 9, ContextMenuExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 10, CreatedEditControlWindowExtvwu)
		//SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 11, CustomDrawExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 12, DblClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 13, DestroyedControlWindowExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 14, DestroyedEditControlWindowExtvwu)
		//SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 15, DragMouseMoveExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 16, DropExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 17, EditChangeExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 18, EditClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 19, EditContextMenuExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 20, EditDblClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 21, EditGotFocusExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 22, EditKeyDownExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 23, EditKeyPressExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 24, EditKeyUpExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 25, EditLostFocusExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 26, EditMClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 27, EditMDblClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 28, EditMouseDownExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 29, EditMouseEnterExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 30, EditMouseHoverExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 31, EditMouseLeaveExtvwu)
		//SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 32, EditMouseMoveExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 33, EditMouseUpExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 34, EditRClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 35, EditRDblClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 36, ExpandedItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 37, ExpandingItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 38, FreeItemDataExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 39, IncrementalSearchStringChangingExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 40, InsertedItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 41, InsertingItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 42, ItemBeginDragExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 43, ItemBeginRDragExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 44, ItemGetDisplayInfoExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 45, ItemGetInfoTipTextExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 46, ItemMouseEnterExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 47, ItemMouseLeaveExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 48, ItemSelectionChangedExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 49, ItemSetTextExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 50, ItemStateImageChangedExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 51, ItemStateImageChangingExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 52, KeyDownExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 53, KeyPressExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 54, KeyUpExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 55, MClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 56, MDblClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 57, MouseDownExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 58, MouseEnterExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 59, MouseHoverExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 60, MouseLeaveExtvwu)
		//SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 61, MouseMoveExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 62, MouseUpExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 63, OLECompleteDragExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 64, OLEDragDropExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 65, OLEDragEnterExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 66, OLEDragLeaveExtvwu)
		//SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 67, OLEDragMouseMoveExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 68, OLEGiveFeedbackExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 69, OLEQueryContinueDragExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 70, OLESetDataExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 71, OLEStartDragExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 72, RClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 73, RDblClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 74, RecreatedControlWindowExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 75, RemovedItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 76, RemovingItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 77, RenamedItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 78, RenamingItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 79, ResizedControlWindowExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 80, SingleExpandingExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 81, StartingLabelEditingExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 82, ItemAsynchronousDrawFailedExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 83, ItemSelectionChangingExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 84, OLEDragEnterPotentialTargetExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 85, OLEDragLeavePotentialTargetExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 86, OLEReceivedNewDataExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 87, ChangedSortOrderExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 88, ChangingSortOrderExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 89, EditMouseWheelExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 90, EditXClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 91, EditXDblClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 92, MouseWheelExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 93, XClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 94, XDblClickExtvwu)

		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 1, AbortedDragExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 0, CaretChangedExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 4, CaretChangingExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 5, ClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 6, CollapsedItemExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 7, CollapsingItemExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 8, CompareItemsExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 9, ContextMenuExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 10, CreatedEditControlWindowExtvwa)
		//SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 11, CustomDrawExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 12, DblClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 13, DestroyedControlWindowExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 14, DestroyedEditControlWindowExtvwa)
		//SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 15, DragMouseMoveExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 16, DropExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 17, EditChangeExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 18, EditClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 19, EditContextMenuExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 20, EditDblClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 21, EditGotFocusExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 22, EditKeyDownExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 23, EditKeyPressExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 24, EditKeyUpExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 25, EditLostFocusExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 26, EditMClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 27, EditMDblClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 28, EditMouseDownExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 29, EditMouseEnterExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 30, EditMouseHoverExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 31, EditMouseLeaveExtvwa)
		//SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 32, EditMouseMoveExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 33, EditMouseUpExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 34, EditRClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 35, EditRDblClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 36, ExpandedItemExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 37, ExpandingItemExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 38, FreeItemDataExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 39, IncrementalSearchStringChangingExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 40, InsertedItemExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 41, InsertingItemExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 42, ItemBeginDragExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 43, ItemBeginRDragExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 44, ItemGetDisplayInfoExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 45, ItemGetInfoTipTextExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 46, ItemMouseEnterExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 47, ItemMouseLeaveExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 48, ItemSelectionChangedExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 49, ItemSetTextExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 50, ItemStateImageChangedExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 51, ItemStateImageChangingExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 52, KeyDownExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 53, KeyPressExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 54, KeyUpExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 55, MClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 56, MDblClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 57, MouseDownExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 58, MouseEnterExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 59, MouseHoverExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 60, MouseLeaveExtvwa)
		//SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 61, MouseMoveExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 62, MouseUpExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 63, OLECompleteDragExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 64, OLEDragDropExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 65, OLEDragEnterExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 66, OLEDragLeaveExtvwa)
		//SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 67, OLEDragMouseMoveExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 68, OLEGiveFeedbackExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 69, OLEQueryContinueDragExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 70, OLESetDataExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 71, OLEStartDragExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 72, RClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 73, RDblClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 74, RecreatedControlWindowExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 75, RemovedItemExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 76, RemovingItemExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 77, RenamedItemExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 78, RenamingItemExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 79, ResizedControlWindowExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 80, SingleExpandingExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 81, StartingLabelEditingExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 82, ItemAsynchronousDrawFailedExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 83, ItemSelectionChangingExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 84, OLEDragEnterPotentialTargetExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 85, OLEDragLeavePotentialTargetExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 86, OLEReceivedNewDataExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 87, ChangedSortOrderExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 88, ChangingSortOrderExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 89, EditMouseWheelExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 90, EditXClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 91, EditXDblClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 92, MouseWheelExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 93, XClickExtvwa)
		SINK_ENTRY_EX(IDC_EXTVWA, __uuidof(ExTVwLibA::_IExplorerTreeViewEvents), 94, XDblClickExtvwa)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
		DLGRESIZE_CONTROL(IDC_EXTVWU, DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(IDC_EXTVWA, DLSZ_SIZE_Y | DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(IDC_EDITLOG, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_Y)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSetFocusU(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnSetFocusA(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);

	void AddLogEntry(CAtlString text);
	void CloseDialog(int nVal);
	void InsertItemsA(void);
	void InsertItemsU(void);

	void __stdcall AbortedDragExtvwu();
	void __stdcall CaretChangedExtvwu(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, long caretChangeReason);
	void __stdcall CaretChangingExtvwu(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, long caretChangeReason, VARIANT_BOOL* cancelCaretChange);
	void __stdcall ChangedSortOrderExtvwu(long previousSortOrder, long newSortOrder);
	void __stdcall ChangingSortOrderExtvwu(long previousSortOrder, long newSortOrder, VARIANT_BOOL* cancelChange);
	void __stdcall ClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall CollapsedItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL removedAllSubItems);
	void __stdcall CollapsingItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL removingAllSubItems, VARIANT_BOOL* cancelCollapse);
	void __stdcall CompareItemsExtvwu(LPDISPATCH firstItem, LPDISPATCH secondItem, long* result);
	void __stdcall ContextMenuExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CreatedEditControlWindowExtvwu(long hWndEdit);
	void __stdcall CustomDrawExtvwu(LPDISPATCH treeItem, long* itemLevel, OLE_COLOR* textColor, OLE_COLOR* textBackColor, ExTVwLibU::CustomDrawStageConstants drawStage, ExTVwLibU::CustomDrawItemStateConstants itemState, long hDC, ExTVwLibU::RECTANGLE* drawingRectangle, ExTVwLibU::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall DestroyedControlWindowExtvwu(long hWnd);
	void __stdcall DestroyedEditControlWindowExtvwu(long hWndEdit);
	void __stdcall DragMouseMoveExtvwu(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails, VARIANT_BOOL* autoExpandItem, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall DropExtvwu(LPDISPATCH dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails);
	void __stdcall EditChangeExtvwu();
	void __stdcall EditClickExtvwu(short button, short shift, long x, long y);
	void __stdcall EditContextMenuExtvwu(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall EditDblClickExtvwu(short button, short shift, long x, long y);
	void __stdcall EditGotFocusExtvwu();
	void __stdcall EditKeyDownExtvwu(short* keyCode, short shift);
	void __stdcall EditKeyPressExtvwu(short* keyAscii);
	void __stdcall EditKeyUpExtvwu(short* keyCode, short shift);
	void __stdcall EditLostFocusExtvwu();
	void __stdcall EditMClickExtvwu(short button, short shift, long x, long y);
	void __stdcall EditMDblClickExtvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseDownExtvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseEnterExtvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseHoverExtvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseLeaveExtvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseMoveExtvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseUpExtvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseWheelExtvwu(short button, short shift, long x, long y, long scrollAxis, short wheelDelta);
	void __stdcall EditRClickExtvwu(short button, short shift, long x, long y);
	void __stdcall EditRDblClickExtvwu(short button, short shift, long x, long y);
	void __stdcall EditXClickExtvwu(short button, short shift, long x, long y);
	void __stdcall EditXDblClickExtvwu(short button, short shift, long x, long y);
	void __stdcall ExpandedItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL expandedPartially);
	void __stdcall ExpandingItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL expandingPartially, VARIANT_BOOL* cancelExpansion);
	void __stdcall FreeItemDataExtvwu(LPDISPATCH treeItem);
	void __stdcall IncrementalSearchStringChangingExtvwu(BSTR currentSearchString, short keyCodeOfCharToBeAdded, VARIANT_BOOL* cancelChange);
	void __stdcall InsertedItemExtvwu(LPDISPATCH treeItem);
	void __stdcall InsertingItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL* cancelInsertion);
	void __stdcall ItemAsynchronousDrawFailedExtvwu(LPDISPATCH treeItem, ExTVwLibU::FAILEDIMAGEDETAILS* imageDetails, ExTVwLibU::ImageDrawingFailureReasonConstants failureReason, ExTVwLibU::FailedAsyncDrawReturnValuesConstants* furtherProcessing, long* newImageToDraw);
	void __stdcall ItemBeginDragExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemBeginRDragExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemGetDisplayInfoExtvwu(LPDISPATCH treeItem, long requestedInfo, long* iconIndex, long* selectedIconIndex, VARIANT_BOOL* hasExpando, long maxItemTextLength, BSTR* itemText, VARIANT_BOOL* dontAskAgain);
	void __stdcall ItemGetInfoTipTextExtvwu(LPDISPATCH treeItem, long maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip);
	void __stdcall ItemMouseEnterExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemMouseLeaveExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemSelectionChangedExtvwu(LPDISPATCH treeItem);
	void __stdcall ItemSelectionChangingExtvwu(LPDISPATCH treeItem, VARIANT_BOOL* cancelChange);
	void __stdcall ItemSetTextExtvwu(LPDISPATCH treeItem, BSTR itemText);
	void __stdcall ItemStateImageChangedExtvwu(LPDISPATCH treeItem, long previousStateImageIndex, long newStateImageIndex, long causedBy);
	void __stdcall ItemStateImageChangingExtvwu(LPDISPATCH treeItem, long previousStateImageIndex, long* newStateImageIndex, long causedBy, VARIANT_BOOL* cancelChange);
	void __stdcall KeyDownExtvwu(short* keyCode, short shift);
	void __stdcall KeyPressExtvwu(short* keyAscii);
	void __stdcall KeyUpExtvwu(short* keyCode, short shift);
	void __stdcall MClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MDblClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseDownExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseEnterExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseHoverExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseLeaveExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseMoveExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseUpExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseWheelExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails, long scrollAxis, short wheelDelta);
	void __stdcall OLECompleteDragExtvwu(LPDISPATCH data, long performedEffect);
	void __stdcall OLEDragDropExtvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails);
	void __stdcall OLEDragEnterExtvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails, VARIANT_BOOL* autoExpandItem, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEDragEnterPotentialTargetExtvwu(long hWndPotentialTarget);
	void __stdcall OLEDragLeaveExtvwu(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails);
	void __stdcall OLEDragLeavePotentialTargetExtvwu(void);
	void __stdcall OLEDragMouseMoveExtvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails, VARIANT_BOOL* autoExpandItem, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEGiveFeedbackExtvwu(long effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragExtvwu(VARIANT_BOOL pressedEscape, short button, short shift, long* actionToContinueWith);
	void __stdcall OLEReceivedNewDataExtvwu(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLESetDataExtvwu(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLEStartDragExtvwu(LPDISPATCH data);
	void __stdcall RClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RDblClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RecreatedControlWindowExtvwu(long hWnd);
	void __stdcall RemovedItemExtvwu(LPDISPATCH treeItem);
	void __stdcall RemovingItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL* cancelDeletion);
	void __stdcall RenamedItemExtvwu(LPDISPATCH treeItem, BSTR previousItemText, BSTR newItemText);
	void __stdcall RenamingItemExtvwu(LPDISPATCH treeItem, BSTR previousItemText, BSTR newItemText, VARIANT_BOOL* cancelRenaming);
	void __stdcall ResizedControlWindowExtvwu();
	void __stdcall SingleExpandingExtvwu(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, VARIANT_BOOL* dontChangePreviousItem, VARIANT_BOOL* dontChangeNewItem);
	void __stdcall StartingLabelEditingExtvwu(LPDISPATCH treeItem, VARIANT_BOOL* cancelEditing);
	void __stdcall XClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall XDblClickExtvwu(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);

	void __stdcall AbortedDragExtvwa();
	void __stdcall CaretChangedExtvwa(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, long caretChangeReason);
	void __stdcall CaretChangingExtvwa(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, long caretChangeReason, VARIANT_BOOL* cancelCaretChange);
	void __stdcall ChangedSortOrderExtvwa(long previousSortOrder, long newSortOrder);
	void __stdcall ChangingSortOrderExtvwa(long previousSortOrder, long newSortOrder, VARIANT_BOOL* cancelChange);
	void __stdcall ClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall CollapsedItemExtvwa(LPDISPATCH treeItem, VARIANT_BOOL removedAllSubItems);
	void __stdcall CollapsingItemExtvwa(LPDISPATCH treeItem, VARIANT_BOOL removingAllSubItems, VARIANT_BOOL* cancelCollapse);
	void __stdcall CompareItemsExtvwa(LPDISPATCH firstItem, LPDISPATCH secondItem, long* result);
	void __stdcall ContextMenuExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CreatedEditControlWindowExtvwa(long hWndEdit);
	void __stdcall CustomDrawExtvwa(LPDISPATCH treeItem, long* itemLevel, OLE_COLOR* textColor, OLE_COLOR* textBackColor, ExTVwLibA::CustomDrawStageConstants drawStage, ExTVwLibA::CustomDrawItemStateConstants itemState, long hDC, ExTVwLibA::RECTANGLE* drawingRectangle, ExTVwLibA::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall DestroyedControlWindowExtvwa(long hWnd);
	void __stdcall DestroyedEditControlWindowExtvwa(long hWndEdit);
	void __stdcall DragMouseMoveExtvwa(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails, VARIANT_BOOL* autoExpandItem, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall DropExtvwa(LPDISPATCH dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails);
	void __stdcall EditChangeExtvwa();
	void __stdcall EditClickExtvwa(short button, short shift, long x, long y);
	void __stdcall EditContextMenuExtvwa(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall EditDblClickExtvwa(short button, short shift, long x, long y);
	void __stdcall EditGotFocusExtvwa();
	void __stdcall EditKeyDownExtvwa(short* keyCode, short shift);
	void __stdcall EditKeyPressExtvwa(short* keyAscii);
	void __stdcall EditKeyUpExtvwa(short* keyCode, short shift);
	void __stdcall EditLostFocusExtvwa();
	void __stdcall EditMClickExtvwa(short button, short shift, long x, long y);
	void __stdcall EditMDblClickExtvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseDownExtvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseEnterExtvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseHoverExtvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseLeaveExtvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseMoveExtvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseUpExtvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseWheelExtvwa(short button, short shift, long x, long y, long scrollAxis, short wheelDelta);
	void __stdcall EditRClickExtvwa(short button, short shift, long x, long y);
	void __stdcall EditRDblClickExtvwa(short button, short shift, long x, long y);
	void __stdcall EditXClickExtvwa(short button, short shift, long x, long y);
	void __stdcall EditXDblClickExtvwa(short button, short shift, long x, long y);
	void __stdcall ExpandedItemExtvwa(LPDISPATCH treeItem, VARIANT_BOOL expandedPartially);
	void __stdcall ExpandingItemExtvwa(LPDISPATCH treeItem, VARIANT_BOOL expandingPartially, VARIANT_BOOL* cancelExpansion);
	void __stdcall FreeItemDataExtvwa(LPDISPATCH treeItem);
	void __stdcall IncrementalSearchStringChangingExtvwa(BSTR currentSearchString, short keyCodeOfCharToBeAdded, VARIANT_BOOL* cancelChange);
	void __stdcall InsertedItemExtvwa(LPDISPATCH treeItem);
	void __stdcall InsertingItemExtvwa(LPDISPATCH treeItem, VARIANT_BOOL* cancelInsertion);
	void __stdcall ItemAsynchronousDrawFailedExtvwa(LPDISPATCH treeItem, ExTVwLibA::FAILEDIMAGEDETAILS* imageDetails, ExTVwLibA::ImageDrawingFailureReasonConstants failureReason, ExTVwLibA::FailedAsyncDrawReturnValuesConstants* furtherProcessing, long* newImageToDraw);
	void __stdcall ItemBeginDragExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemBeginRDragExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemGetDisplayInfoExtvwa(LPDISPATCH treeItem, long requestedInfo, long* iconIndex, long* selectedIconIndex, VARIANT_BOOL* hasExpando, long maxItemTextLength, BSTR* itemText, VARIANT_BOOL* dontAskAgain);
	void __stdcall ItemGetInfoTipTextExtvwa(LPDISPATCH treeItem, long maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip);
	void __stdcall ItemMouseEnterExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemMouseLeaveExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemSelectionChangedExtvwa(LPDISPATCH treeItem);
	void __stdcall ItemSelectionChangingExtvwa(LPDISPATCH treeItem, VARIANT_BOOL* cancelChange);
	void __stdcall ItemSetTextExtvwa(LPDISPATCH treeItem, BSTR itemText);
	void __stdcall ItemStateImageChangedExtvwa(LPDISPATCH treeItem, long previousStateImageIndex, long newStateImageIndex, long causedBy);
	void __stdcall ItemStateImageChangingExtvwa(LPDISPATCH treeItem, long previousStateImageIndex, long* newStateImageIndex, long causedBy, VARIANT_BOOL* cancelChange);
	void __stdcall KeyDownExtvwa(short* keyCode, short shift);
	void __stdcall KeyPressExtvwa(short* keyAscii);
	void __stdcall KeyUpExtvwa(short* keyCode, short shift);
	void __stdcall MClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MDblClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseDownExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseEnterExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseHoverExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseLeaveExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseMoveExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseUpExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseWheelExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails, long scrollAxis, short wheelDelta);
	void __stdcall OLECompleteDragExtvwa(LPDISPATCH data, long performedEffect);
	void __stdcall OLEDragDropExtvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails);
	void __stdcall OLEDragEnterExtvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails, VARIANT_BOOL* autoExpandItem, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEDragEnterPotentialTargetExtvwa(long hWndPotentialTarget);
	void __stdcall OLEDragLeaveExtvwa(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails);
	void __stdcall OLEDragLeavePotentialTargetExtvwa(void);
	void __stdcall OLEDragMouseMoveExtvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long yToItemTop, long hitTestDetails, VARIANT_BOOL* autoExpandItem, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEGiveFeedbackExtvwa(long effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragExtvwa(VARIANT_BOOL pressedEscape, short button, short shift, long* actionToContinueWith);
	void __stdcall OLEReceivedNewDataExtvwa(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLESetDataExtvwa(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLEStartDragExtvwa(LPDISPATCH data);
	void __stdcall RClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RDblClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RecreatedControlWindowExtvwa(long hWnd);
	void __stdcall RemovedItemExtvwa(LPDISPATCH treeItem);
	void __stdcall RemovingItemExtvwa(LPDISPATCH treeItem, VARIANT_BOOL* cancelDeletion);
	void __stdcall RenamedItemExtvwa(LPDISPATCH treeItem, BSTR previousItemText, BSTR newItemText);
	void __stdcall RenamingItemExtvwa(LPDISPATCH treeItem, BSTR previousItemText, BSTR newItemText, VARIANT_BOOL* cancelRenaming);
	void __stdcall ResizedControlWindowExtvwa();
	void __stdcall SingleExpandingExtvwa(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, VARIANT_BOOL* dontChangePreviousItem, VARIANT_BOOL* dontChangeNewItem);
	void __stdcall StartingLabelEditingExtvwa(LPDISPATCH treeItem, VARIANT_BOOL* cancelEditing);
	void __stdcall XClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall XDblClickExtvwa(LPDISPATCH treeItem, short button, short shift, long x, long y, long hitTestDetails);
};
