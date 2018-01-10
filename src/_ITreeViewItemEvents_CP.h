//////////////////////////////////////////////////////////////////////
/// \class Proxy_ITreeViewItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _ITreeViewItemEvents interface</em>
///
/// \if UNICODE
///   \sa TreeViewItem, ExTVwLibU::_ITreeViewItemEvents
/// \else
///   \sa TreeViewItem, ExTVwLibA::_ITreeViewItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_ITreeViewItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_ITreeViewItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_ITreeViewItemEvents