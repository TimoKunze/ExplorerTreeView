#pragma once

typedef struct ITEMPROPS
{
	int iconIndex;
	int selectedIconIndex;
	int itemData;
	int heightIncrement;
	CAtlString text;
	BOOL bold;
	BOOL ghosted;
	BOOL hasExpando;

	ITEMPROPS()
	{
		iconIndex = 0;
		selectedIconIndex = 0;
		itemData = 0;
		heightIncrement = 1;
		text = TEXT("");
		bold = FALSE;
		ghosted = FALSE;
		hasExpando = FALSE;
	}
} ITEMPROPS, *LPITEMPROPS;