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
    public IDispEventImpl<IDC_EXTVWU1, CMainDlg, &__uuidof(_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 6>,
    public IDispEventImpl<IDC_EXTVWU2, CMainDlg, &__uuidof(_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 6>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> extvwU1Wnd;
	CContainedWindowT<CAxWindow> extvwU2Wnd;

	CMainDlg() :
	    extvwU1Wnd(this, 1),
	    extvwU2Wnd(this, 2)
	{
	}

	struct Controls
	{
		CButton enabledCheck;
		CImageList imageList;
		CComPtr<IExplorerTreeView> extvwU1;
		CComPtr<IExplorerTreeView> extvwU2;

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

		COMMAND_HANDLER(IDC_ENABLEDCHK, BN_CLICKED, OnBnClickedEnabledchk)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		ALT_MSG_MAP(2)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_EXTVWU1, __uuidof(_IExplorerTreeViewEvents), 0, CaretChangedExtvwu1)
		SINK_ENTRY_EX(IDC_EXTVWU1, __uuidof(_IExplorerTreeViewEvents), 7, CollapsingItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU1, __uuidof(_IExplorerTreeViewEvents), 11, CustomDrawExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU1, __uuidof(_IExplorerTreeViewEvents), 37, ExpandingItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU1, __uuidof(_IExplorerTreeViewEvents), 45, ItemGetInfoTipTextExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU1, __uuidof(_IExplorerTreeViewEvents), 74, RecreatedControlWindowExtvwu1)
		SINK_ENTRY_EX(IDC_EXTVWU1, __uuidof(_IExplorerTreeViewEvents), 81, StartingLabelEditingExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU2, __uuidof(_IExplorerTreeViewEvents), 4, CaretChangingExtvwu2)
		SINK_ENTRY_EX(IDC_EXTVWU2, __uuidof(_IExplorerTreeViewEvents), 7, CollapsingItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU2, __uuidof(_IExplorerTreeViewEvents), 11, CustomDrawExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU2, __uuidof(_IExplorerTreeViewEvents), 37, ExpandingItemExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU2, __uuidof(_IExplorerTreeViewEvents), 45, ItemGetInfoTipTextExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU2, __uuidof(_IExplorerTreeViewEvents), 74, RecreatedControlWindowExtvwu2)
		SINK_ENTRY_EX(IDC_EXTVWU2, __uuidof(_IExplorerTreeViewEvents), 81, StartingLabelEditingExtvwu)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnBnClickedEnabledchk(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
	void InsertItems1(void);
	void InsertItems2(void);

	void __stdcall CaretChangedExtvwu1(LPDISPATCH /*previousCaretItem*/, LPDISPATCH newCaretItem, long /*caretChangeReason*/);
	void __stdcall CaretChangingExtvwu2(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, long /*caretChangeReason*/, VARIANT_BOOL* cancelCaretChange);
	void __stdcall CollapsingItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL /*removingAllSubItems*/, VARIANT_BOOL* cancelCollapse);
	void __stdcall CustomDrawExtvwu(LPDISPATCH treeItem, long* /*itemLevel*/, OLE_COLOR* textColor, OLE_COLOR* textBackColor, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants itemState, long /*hDC*/, RECTANGLE* /*drawingRectangle*/, CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall ExpandingItemExtvwu(LPDISPATCH treeItem, VARIANT_BOOL /*expandingPartially*/, VARIANT_BOOL* cancelExpansion);
	void __stdcall ItemGetInfoTipTextExtvwu(LPDISPATCH treeItem, long /*maxInfoTipLength*/, BSTR* /*infoTipText*/, VARIANT_BOOL* abortToolTip);
	void __stdcall RecreatedControlWindowExtvwu1(long /*hWnd*/);
	void __stdcall RecreatedControlWindowExtvwu2(long /*hWnd*/);
	void __stdcall StartingLabelEditingExtvwu(LPDISPATCH treeItem, VARIANT_BOOL* cancelEditing);
};
