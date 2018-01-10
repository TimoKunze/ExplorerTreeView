//////////////////////////////////////////////////////////////////////
/// \class Proxy_ITreeViewItemContainerEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _ITreeViewItemContainerEvents interface</em>
///
/// \if UNICODE
///   \sa TreeViewItemContainer, ExTVwLibU::_ITreeViewItemContainerEvents
/// \else
///   \sa TreeViewItemContainer, ExTVwLibA::_ITreeViewItemContainerEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_ITreeViewItemContainerEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_ITreeViewItemContainerEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_ITreeViewItemContainerEvents