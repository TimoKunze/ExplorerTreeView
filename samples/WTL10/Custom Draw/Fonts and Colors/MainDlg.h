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

	typedef struct ENUMFONTPARAM
	{
		BOOL firstCall;
		LOGFONT lfTreeView;
		CComPtr<ITreeViewItem> pParentItem;
		CComPtr<ITreeViewItems> pSubItemsToAddTo;
	} ENUMFONTPARAM, *LPENUMFONTPARAM;
	static int CALLBACK EnumFontFamExProc(const LPENUMLOGFONTEX lpElfe, const NEWTEXTMETRICEX* /*lpntme*/, DWORD /*FontType*/, LPARAM lParam);

	CMainDlg() :
	    extvwUWnd(this, 1)
	{
	}

	HFONT hFontBold;

	struct Controls
	{
		CComPtr<IExplorerTreeView> extvwU;
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
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 11, CustomDrawExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 38, FreeItemDataExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 45, ItemGetInfoTipTextExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 74, RecreatedControlWindowExtvwu)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
	void InsertItems(void);

	void __stdcall CustomDrawExtvwu(LPDISPATCH treeItem, long* /*itemLevel*/, OLE_COLOR* textColor, OLE_COLOR* /*textBackColor*/, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants itemState, long hDC, RECTANGLE* /*drawingRectangle*/, CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall FreeItemDataExtvwu(LPDISPATCH treeItem);
	void __stdcall ItemGetInfoTipTextExtvwu(LPDISPATCH treeItem, long /*maxInfoTipLength*/, BSTR* infoTipText, VARIANT_BOOL* /*abortToolTip*/);
	void __stdcall RecreatedControlWindowExtvwu(long /*hWnd*/);
};
