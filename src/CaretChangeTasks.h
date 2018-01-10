//////////////////////////////////////////////////////////////////////
/// \class CaretChangeTasks
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A task manager used during internal caret changes</em>
///
/// This class provides a task manager for tasks that are executed during internal caret changes.
///
/// \todo Improve documentation ("See also" sections).
///
/// \sa ICaretChangeTask, CCEnsureVisibleTask, ExplorerTreeView
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#include "ICaretChangeTask.h"


class CaretChangeTasks
{
public:
	/// \brief <em>The destructor of this class</em>
	///
	/// Used to destroy all remaining tasks.
	~CaretChangeTasks();

	/// \brief <em>Adds a new task</em>
	///
	/// Adds a new task to the list of tasks to execute.
	///
	/// \param[in] executeOnCaretChanging Execute the task on \c TVN_SELCHANGING notification.
	/// \param[in] executeOnCaretChanged Execute the task on \c TVN_SELCHANGED notification.
	/// \param[in] pTask An object that implements \c ICaretChangeTask.
	///
	/// \sa ExplorerTreeView::AddCaretChangeTask, ProcessOnCaretChangedTasks, ProcessOnCaretChangingTasks,
	///     RemoveAllTasks
	void AddTask(BOOL executeOnCaretChanging, BOOL executeOnCaretChanged, ICaretChangeTask* pTask);
	/// \brief <em>Executes all matching tasks</em>
	///
	/// Calls the \c ICaretChangeTask::Execute method for all tasks that were registered to be executed
	/// from the \c TVN_SELCHANGED notification handler.
	///
	/// \param[in] pExTvw The \c ExplorerTreeView object the method will work on.
	/// \param[in] hWndTvw The treeview window the method will work on.
	/// \param[in] autoRemoveTasks If set to \c TRUE, the task will be removed from the list after it
	///            got executed.
	///
	/// \sa ICaretChangeTask::Execute, AddTask, ProcessOnCaretChangingTasks, RemoveAllTasks
	void ProcessOnCaretChangedTasks(ExplorerTreeView* pExTvw, HWND hWndTvw, BOOL autoRemoveTasks = TRUE);
	/// \brief <em>Executes all matching tasks</em>
	///
	/// Calls the \c ICaretChangeTask::Execute method for all tasks that were registered to be executed
	/// from the \c TVN_SELCHANGING notification handler.
	///
	/// \param[in] pExTvw The \c ExplorerTreeView object the method will work on.
	/// \param[in] hWndTvw The treeview window the method will work on.
	/// \param[in] autoRemoveTasks If set to \c TRUE, the task will be removed from the list after it
	///            got executed.
	///
	/// \sa ICaretChangeTask::Execute, AddTask, ProcessOnCaretChangedTasks, RemoveAllTasks
	void ProcessOnCaretChangingTasks(ExplorerTreeView* pExTvw, HWND hWndTvw, BOOL autoRemoveTasks = TRUE);
	/// \brief <em>Removes all matching tasks</em>
	///
	/// Removes all matching tasks from the list of tasks to execute. The task objects get destroyed.
	///
	/// \param[in] onCaretChangingTasks If set to \c TRUE, tasks, that were registered for the
	///            \c TVN_SELCHANGING notification, get removed.
	/// \param[in] onCaretChangedTasks If set to \c TRUE, tasks, that were registered for the
	///            \c TVN_SELCHANGED notification, get removed.
	///
	/// \sa AddTask, ProcessOnCaretChangedTasks, ProcessOnCaretChangingTasks
	void RemoveAllTasks(BOOL onCaretChangingTasks, BOOL onCaretChangedTasks);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		#ifdef USE_STL
			/// \brief <em>Holds all tasks registered for the \c TVN_SELCHANGED notification</em>
			std::vector<ICaretChangeTask*> onCaretChangedTasks;
			/// \brief <em>Holds all tasks registered for the \c TVN_SELCHANGING notification</em>
			std::vector<ICaretChangeTask*> onCaretChangingTasks;
		#else
			/// \brief <em>Holds all tasks registered for the \c TVN_SELCHANGED notification</em>
			CAtlArray<ICaretChangeTask*> onCaretChangedTasks;
			/// \brief <em>Holds all tasks registered for the \c TVN_SELCHANGING notification</em>
			CAtlArray<ICaretChangeTask*> onCaretChangingTasks;
		#endif
	} properties;
};     // CaretChangeTasks