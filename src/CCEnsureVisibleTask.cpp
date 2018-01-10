// CCEnsureVisibleTask.cpp: a task that ensures a given item is visible

#include "stdafx.h"
#include "res/resource.h"
#include "CCEnsureVisibleTask.h"
#include "ExplorerTreeView.h"


CCEnsureVisibleTask::CCEnsureVisibleTask(HTREEITEM hItem)
{
	properties.hItem = hItem;
}


ICaretChangeTask* CCEnsureVisibleTask::Clone(void)
{
	return new CCEnsureVisibleTask(properties.hItem);
}

void CCEnsureVisibleTask::Execute(ExplorerTreeView* /*pExTvw*/, HWND hWndTvw)
{
	ATLASSERT(IsWindow(hWndTvw));

	SendMessage(hWndTvw, TVM_ENSUREVISIBLE, 0, reinterpret_cast<LPARAM>(properties.hItem));
}