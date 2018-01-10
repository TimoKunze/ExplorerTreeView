//////////////////////////////////////////////////////////////////////
/// \class CCEnsureVisibleTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A task that ensures a given item is visible</em>
///
/// This class provides a task usable with the \c CaretChangeTasks task manager. It ensures that a
/// given item in a given \c ExplorerTreeView is visible to the user.
///
/// \todo Improve documentation ("See also" sections).
///
/// \sa ICaretChangeTask, CaretChangeTasks, ExplorerTreeView
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#include "ExplorerTreeView.h"
#include "ICaretChangeTask.h"


class CCEnsureVisibleTask :
    public ICaretChangeTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// \param[in] hItem The item this class will work on.
	CCEnsureVisibleTask(HTREEITEM hItem);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICaretChangeTask
	///
	//@{
	/// \brief <em>Returns an exact copy of this object</em>
	///
	/// \return The copy.
	ICaretChangeTask* Clone(void);
	/// \brief <em>Executes this task</em>
	///
	/// \param[in] pExTvw The \c ExplorerTreeView object the method will work on.
	/// \param[in] hWndTvw The treeview window the method will work on.
	void Execute(ExplorerTreeView* pExTvw, HWND hWndTvw);
	//@}
	//////////////////////////////////////////////////////////////////////

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>Holds the item this class will work on</em>
		HTREEITEM hItem;
	} properties;
};     // CCEnsureVisibleTask