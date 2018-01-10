// EditItemDlg.cpp : implementation of the CEditItemDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"
#include "EditItemDlg.h"

BOOL CEditItemDlg::PreTranslateMessage(MSG* pMsg)
{
	return CWindow::IsDialogMessage(pMsg);
}

LRESULT CEditItemDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*bHandled*/)
{
	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	controls.textEdit = GetDlgItem(IDC_TEXT);
	controls.ghostedCheckBox = GetDlgItem(IDC_GHOSTED);
	controls.boldCheckBox = GetDlgItem(IDC_BOLD);
	controls.expandoCheckBox = GetDlgItem(IDC_EXPANDO);
	controls.iconIndexEdit = GetDlgItem(IDC_ICONINDEX);
	controls.itemDataEdit = GetDlgItem(IDC_ITEMDATA);
	controls.selIconIndexEdit = GetDlgItem(IDC_SELECTEDICONINDEX);
	controls.heightIncrEdit = GetDlgItem(IDC_HEIGHTINCREMENT);

	pItemProps = reinterpret_cast<LPITEMPROPS>(lParam);
	controls.textEdit.SetWindowText(pItemProps->text);
	controls.ghostedCheckBox.SetCheck(pItemProps->ghosted ? BST_CHECKED : BST_UNCHECKED);
	controls.boldCheckBox.SetCheck(pItemProps->bold ? BST_CHECKED : BST_UNCHECKED);
	controls.expandoCheckBox.SetCheck(pItemProps->hasExpando ? BST_CHECKED : BST_UNCHECKED);
	CAtlString tmp;
	tmp.Format(TEXT("%i"), pItemProps->iconIndex);
	controls.iconIndexEdit.SetWindowText(tmp);
	tmp.Format(TEXT("%i"), pItemProps->itemData);
	controls.itemDataEdit.SetWindowText(tmp);
	tmp.Format(TEXT("%i"), pItemProps->selectedIconIndex);
	controls.selIconIndexEdit.SetWindowText(tmp);
	tmp.Format(TEXT("%i"), pItemProps->heightIncrement);
	controls.heightIncrEdit.SetWindowText(tmp);

	return TRUE;
}

LRESULT CEditItemDlg::OnOK(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	int len = controls.textEdit.GetWindowTextLength() + 1;
	LPTSTR pBuffer = reinterpret_cast<LPTSTR>(malloc(len * sizeof(TCHAR)));
	controls.textEdit.GetWindowText(pBuffer, len);
	pItemProps->text = pBuffer;
	free(pBuffer);

	pItemProps->ghosted = (controls.ghostedCheckBox.GetCheck() == BST_CHECKED);
	pItemProps->bold = (controls.boldCheckBox.GetCheck() == BST_CHECKED);
	pItemProps->hasExpando = (controls.expandoCheckBox.GetCheck() == BST_CHECKED);

	len = controls.iconIndexEdit.GetWindowTextLength() + 1;
	pBuffer = reinterpret_cast<LPTSTR>(malloc(len * sizeof(TCHAR)));
	controls.iconIndexEdit.GetWindowText(pBuffer, len);
	pItemProps->iconIndex = _ttoi(pBuffer);
	free(pBuffer);

	len = controls.itemDataEdit.GetWindowTextLength() + 1;
	pBuffer = reinterpret_cast<LPTSTR>(malloc(len * sizeof(TCHAR)));
	controls.itemDataEdit.GetWindowText(pBuffer, len);
	pItemProps->itemData = _ttoi(pBuffer);
	free(pBuffer);

	len = controls.selIconIndexEdit.GetWindowTextLength() + 1;
	pBuffer = reinterpret_cast<LPTSTR>(malloc(len * sizeof(TCHAR)));
	controls.selIconIndexEdit.GetWindowText(pBuffer, len);
	pItemProps->selectedIconIndex = _ttoi(pBuffer);
	free(pBuffer);

	len = controls.heightIncrEdit.GetWindowTextLength() + 1;
	pBuffer = reinterpret_cast<LPTSTR>(malloc(len * sizeof(TCHAR)));
	controls.heightIncrEdit.GetWindowText(pBuffer, len);
	pItemProps->heightIncrement = _ttoi(pBuffer);
	if(pItemProps->heightIncrement == 0) {
		pItemProps->heightIncrement = 1;
	}
	free(pBuffer);

	EndDialog(wID);
	return 0;
}

LRESULT CEditItemDlg::OnCancel(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	EndDialog(wID);
	return 0;
}
