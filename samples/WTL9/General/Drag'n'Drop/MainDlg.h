// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#import <libid:1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5> version("2.6") named_guids, no_namespace, raw_dispinterfaces

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_EXTVWU, CMainDlg, &__uuidof(_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 6>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> extvwUWnd;

	CMainDlg() :
	    extvwUWnd(this, 1)
	{
		CF_TARGETCLSID = static_cast<CLIPFORMAT>(RegisterClipboardFormat(CFSTR_TARGETCLSID));
	}

	BOOL bRightDrag;
	OLE_HANDLE hLastDropTarget;
	CComQIPtr<IPicture> pDraggedPicture;

	struct Controls
	{
		CButton oleDDCheck;
		CButton vistaStyleCheck;
		CImageList imageList;
		CComPtr<IExplorerTreeView> extvwU;

		~Controls()
		{
			if(!imageList.IsNull()) {
				imageList.Destroy();
			}
		}
	} controls;

	CLIPFORMAT CF_TARGETCLSID;

	BOOL IsComctl32Version600OrNewer(void);
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		COMMAND_ID_HANDLER(ID_OPTIONS_AERODRAGIMAGES, OnOptionsAeroDragImages)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 1, AbortedDragExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 15, DragMouseMoveExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 16, DropExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 42, ItemBeginDragExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 43, ItemBeginRDragExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 52, KeyDownExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 62, MouseUpExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 63, OLECompleteDragExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 64, OLEDragDropExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 65, OLEDragEnterExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 66, OLEDragLeaveExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 67, OLEDragMouseMoveExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 70, OLESetDataExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 71, OLEStartDragExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 74, RecreatedControlWindowExtvwu)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
		DLGRESIZE_CONTROL(IDC_EXTVWU, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(IDC_OLEDDCHECK, DLSZ_MOVE_X)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnOptionsAeroDragImages(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
	InsertMarkPositionConstants CalculateInsertMarkPosition(ITreeViewItem* pDropTarget, long yToItemTop);
	void InsertItems(void);
	BOOL IsSameItem(ITreeViewItem* pItem1, ITreeViewItem* pItem2);
	void UpdateMenu(void);

	void __stdcall AbortedDragExtvwu();
	void __stdcall DragMouseMoveExtvwu(LPDISPATCH* dropTarget, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long yToItemTop, long /*hitTestDetails*/, VARIANT_BOOL* /*autoExpandItem*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/);
	void __stdcall DropExtvwu(LPDISPATCH dropTarget, short /*button*/, short /*shift*/, long x, long y, long /*yToItemTop*/, long /*hitTestDetails*/);
	void __stdcall ItemBeginDragExtvwu(LPDISPATCH treeItem, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/);
	void __stdcall ItemBeginRDragExtvwu(LPDISPATCH treeItem, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/);
	void __stdcall KeyDownExtvwu(short* keyCode, short /*shift*/);
	void __stdcall MouseUpExtvwu(LPDISPATCH /*treeItem*/, short button, short /*shift*/, long /*x*/, long /*y*/, long hitTestDetails);
	void __stdcall OLECompleteDragExtvwu(LPDISPATCH data, long /*performedEffect*/);
	void __stdcall OLEDragDropExtvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long /*x*/, long /*y*/, long /*yToItemTop*/, long /*hitTestDetails*/);
	void __stdcall OLEDragEnterExtvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long /*x*/, long /*y*/, long yToItemTop, long /*hitTestDetails*/, VARIANT_BOOL* /*autoExpandItem*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/);
	void __stdcall OLEDragLeaveExtvwu(LPDISPATCH /*data*/, LPDISPATCH /*dropTarget*/, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*yToItemTop*/, long /*hitTestDetails*/);
	void __stdcall OLEDragMouseMoveExtvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long /*x*/, long /*y*/, long yToItemTop, long /*hitTestDetails*/, VARIANT_BOOL* /*autoExpandItem*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/);
	void __stdcall OLESetDataExtvwu(LPDISPATCH data, long formatID, long /*index*/, long /*dataOrViewAspect*/);
	void __stdcall OLEStartDragExtvwu(LPDISPATCH data);
	void __stdcall RecreatedControlWindowExtvwu(long /*hWnd*/);
};
