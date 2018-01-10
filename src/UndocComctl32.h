//////////////////////////////////////////////////////////////////////
/// \file UndocComctl32.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Undocumented comctl32.dll stuff</em>
///
/// Declaration of some comctl32.dll internals that we're using.
///
/// \todo Improve documentation ("See also" sections).
///
/// \sa ExplorerTreeView
//////////////////////////////////////////////////////////////////////

#pragma once


/// \brief <em>Retrieves a treeview control's internal borders</em>
///
/// This is an undocumented window message to retrieve a treeview's internal borders.
///
/// \param[in] wParam Must be set to 0.
/// \param[in] lParam Must be set to \c NULL.
///
/// \return A \c DWORD value containing the horizontal border in pixels in the lower \c WORD and the
///         vertical border in pixels in the upper \c WORD.
#define TVM_GETBORDER (TV_FIRST + 36)

/// \brief <em>The minimum indentation in pixels</em>
#define TV_MININDENT 5
/// \brief <em>The space between an item's icon and its text in pixels</em>
#define TV_TEXTINDENT 3
