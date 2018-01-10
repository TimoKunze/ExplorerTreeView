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
	}

	struct Controls
	{
		CImageList imageList;
		CComPtr<IExplorerTreeView> extvwU;

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
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 5, ClickExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 9, ContextMenuExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 52, KeyDownExtvwu)
		SINK_ENTRY_EX(IDC_EXTVWU, __uuidof(_IExplorerTreeViewEvents), 74, RecreatedControlWindowExtvwu)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
		DLGRESIZE_CONTROL(IDC_EXTVWU, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_X)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
	void InsertItems(void);

	void __stdcall ClickExtvwu(LPDISPATCH treeItem, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long hitTestDetails);
	void __stdcall ContextMenuExtvwu(LPDISPATCH treeItem, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/, VARIANT_BOOL* /*showDefaultMenu*/);
	void __stdcall KeyDownExtvwu(short* keyCode, short /*shift*/);
	void __stdcall RecreatedControlWindowExtvwu(long /*hWnd*/);
};
