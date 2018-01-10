//////////////////////////////////////////////////////////////////////
/// \class IItemComparator
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communication between the \c CompareItems callback method and the object that initiated the sorting</em>
///
/// This interface allows the \c CompareItems callback method to forward the message to the object
/// that initiated the sorting.
///
/// \sa ::CompareItems, ExplorerTreeView::SortSubItems
//////////////////////////////////////////////////////////////////////


#pragma once


class IItemComparator
{
public:
	/// \brief <em>Compares two items</em>
	///
	/// \param[in] itemID1 The unique ID of the first item to compare.
	/// \param[in] itemID2 The unique ID of the second item to compare.
	///
	/// \return -1 if the first item should precede the second; 0 if the items are equal; 1 if the
	///         second item should precede the first.
	///
	/// \sa ::CompareItems, ExplorerTreeView::SortSubItems
	virtual int CompareItems(LONG itemID1, LONG itemID2) = 0;
};     // IItemComparator