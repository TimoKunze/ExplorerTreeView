// EditItemDlg.h : interface of the CEditItemDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once

#include "helpers.h"

class CEditItemDlg :
    public CDialogImpl<CEditItemDlg>
{
public:
	enum { IDD = IDD_EDITITEMDLG };

	CEditItemDlg()
	{
		pItemProps = NULL;
	}

	LPITEMPROPS pItemProps;

	struct Controls
	{
		CEdit textEdit;
		CButton ghostedCheckBox;
		CButton boldCheckBox;
		CButton expandoCheckBox;
		CEdit iconIndexEdit;
		CEdit itemDataEdit;
		CEdit selIconIndexEdit;
		CEdit heightIncrEdit;
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(IDOK, OnOK)
		COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*bHandled*/);
	LRESULT OnOK(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCancel(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
};
