// CaretChangeTasks.cpp: task manager for tasks being executed during internal caret changes

#include "stdafx.h"
#include "res/resource.h"
#include "helpers.h"
#include "CaretChangeTasks.h"


CaretChangeTasks::~CaretChangeTasks()
{
	RemoveAllTasks(TRUE, TRUE);
}


void CaretChangeTasks::AddTask(BOOL executeOnCaretChanging, BOOL executeOnCaretChanged, ICaretChangeTask* pTask)
{
	if(executeOnCaretChanging) {
		#ifdef USE_STL
			properties.onCaretChangingTasks.push_back(pTask);
		#else
			properties.onCaretChangingTasks.Add(pTask);
		#endif
		if(executeOnCaretChanged) {
			/* If the task is inserted into more than 1 list, we must use clones. Otherwise we'll
			   run into trouble if the task is deleted from one list, because the object is deleted
			   in this case leaving illegal pointers in the other lists. */
			#ifdef USE_STL
				properties.onCaretChangedTasks.push_back(pTask->Clone());
			#else
				properties.onCaretChangedTasks.Add(pTask->Clone());
			#endif
		}
	} else if(executeOnCaretChanged) {
		#ifdef USE_STL
			properties.onCaretChangedTasks.push_back(pTask);
		#else
			properties.onCaretChangedTasks.Add(pTask);
		#endif
	}
}

void CaretChangeTasks::ProcessOnCaretChangedTasks(ExplorerTreeView* pExTvw, HWND hWndTvw, BOOL autoRemoveTasks/* = TRUE*/)
{
	#ifdef USE_STL
		for(std::vector<ICaretChangeTask*>::const_iterator iter = properties.onCaretChangedTasks.begin(); iter < properties.onCaretChangedTasks.end(); ++iter) {
			(*iter)->Execute(pExTvw, hWndTvw);
		}
	#else
		for(size_t i = 0; i < properties.onCaretChangedTasks.GetCount(); ++i) {
			properties.onCaretChangedTasks[i]->Execute(pExTvw, hWndTvw);
		}
	#endif
	if(autoRemoveTasks) {
		RemoveAllTasks(FALSE, TRUE);
	}
}

void CaretChangeTasks::ProcessOnCaretChangingTasks(ExplorerTreeView* pExTvw, HWND hWndTvw, BOOL autoRemoveTasks/* = TRUE*/)
{
	#ifdef USE_STL
		for(std::vector<ICaretChangeTask*>::const_iterator iter = properties.onCaretChangingTasks.begin(); iter < properties.onCaretChangingTasks.end(); ++iter) {
			(*iter)->Execute(pExTvw, hWndTvw);
		}
	#else
		for(size_t i = 0; i < properties.onCaretChangingTasks.GetCount(); ++i) {
			properties.onCaretChangingTasks[i]->Execute(pExTvw, hWndTvw);
		}
	#endif
	if(autoRemoveTasks) {
		RemoveAllTasks(TRUE, FALSE);
	}
}

void CaretChangeTasks::RemoveAllTasks(BOOL onCaretChangingTasks, BOOL onCaretChangedTasks)
{
	if(onCaretChangingTasks) {
		#ifdef USE_STL
			for(std::vector<ICaretChangeTask*>::const_iterator iter = properties.onCaretChangingTasks.begin(); iter < properties.onCaretChangingTasks.end(); ++iter) {
				delete *iter;
			}
			properties.onCaretChangingTasks.clear();
		#else
			for(size_t i = 0; i < properties.onCaretChangingTasks.GetCount(); ++i) {
				delete properties.onCaretChangingTasks[i];
			}
			properties.onCaretChangingTasks.RemoveAll();
		#endif
	}

	if(onCaretChangedTasks) {
		#ifdef USE_STL
			for(std::vector<ICaretChangeTask*>::const_iterator iter = properties.onCaretChangedTasks.begin(); iter < properties.onCaretChangedTasks.end(); ++iter) {
				delete *iter;
			}
			properties.onCaretChangedTasks.clear();
		#else
			for(size_t i = 0; i < properties.onCaretChangedTasks.GetCount(); ++i) {
				delete properties.onCaretChangedTasks[i];
			}
			properties.onCaretChangedTasks.RemoveAll();
		#endif
	}
}