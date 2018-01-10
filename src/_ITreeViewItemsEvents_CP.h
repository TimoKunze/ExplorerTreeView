//////////////////////////////////////////////////////////////////////
/// \class Proxy_ITreeViewItemsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _ITreeViewItemsEvents interface</em>
///
/// \if UNICODE
///   \sa TreeViewItems, ExTVwLibU::_ITreeViewItemsEvents
/// \else
///   \sa TreeViewItems, ExTVwLibA::_ITreeViewItemsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_ITreeViewItemsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_ITreeViewItemsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_ITreeViewItemsEvents