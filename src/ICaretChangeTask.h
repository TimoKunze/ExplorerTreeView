//////////////////////////////////////////////////////////////////////
/// \class ICaretChangeTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Base interface for caret change tasks</em>
///
/// This interface must be implemented by each task that shall be executed on internal caret changes
///
/// \todo Improve documentation ("See also" sections).
///
/// \sa CaretChangeTasks, CCEnsureVisibleTask, ExplorerTreeView
//////////////////////////////////////////////////////////////////////


#pragma once

// foward declaration
class ExplorerTreeView;


class ICaretChangeTask
{
public:
	/// \brief <em>Returns an exact copy of this object</em>
	///
	/// \return The copy.
	virtual ICaretChangeTask* Clone(void) = 0;
	/// \brief <em>Executes this task</em>
	///
	/// \param[in] pExTvw The \c ExplorerTreeView object the method will work on.
	/// \param[in] hWndTvw The treeview window the method will work on.
	virtual void Execute(ExplorerTreeView* pExTvw, HWND hWndTvw) = 0;
};