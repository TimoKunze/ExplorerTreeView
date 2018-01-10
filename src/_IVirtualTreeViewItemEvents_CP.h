//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualTreeViewItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualTreeViewItemEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualTreeViewItem, ExTVwLibU::_IVirtualTreeViewItemEvents
/// \else
///   \sa VirtualTreeViewItem, ExTVwLibA::_IVirtualTreeViewItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualTreeViewItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualTreeViewItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualTreeViewItemEvents