// ExplorerTreeView.cpp: Wraps and extends SysTreeView32.

#include "stdafx.h"
#include "ExplorerTreeView.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
IMEModeConstants ExplorerTreeView::IMEFlags::chineseIMETable[10] = {
    imeOff,
    imeOff,
    imeOff,
    imeOff,
    imeOn,
    imeOn,
    imeOn,
    imeOn,
    imeOn,
    imeOff
};

IMEModeConstants ExplorerTreeView::IMEFlags::japaneseIMETable[10] = {
    imeDisable,
    imeDisable,
    imeOff,
    imeOff,
    imeHiragana,
    imeHiragana,
    imeKatakana,
    imeKatakanaHalf,
    imeAlphaFull,
    imeAlpha
};

IMEModeConstants ExplorerTreeView::IMEFlags::koreanIMETable[10] = {
    imeDisable,
    imeDisable,
    imeAlpha,
    imeAlpha,
    imeHangulFull,
    imeHangul,
    imeHangulFull,
    imeHangul,
    imeAlphaFull,
    imeAlpha
};

FONTDESC ExplorerTreeView::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


#pragma warning(push)
#pragma warning(disable: 4355)     // passing "this" to a member constructor
ExplorerTreeView::ExplorerTreeView() :
    containedEdit(WC_EDIT, this, 1)
{
	properties.font.InitializePropertyWatcher(this, DISPID_EXTVW_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_EXTVW_MOUSEICON);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	// initialize
	hItemUnderMouse = NULL;
	hEditBackColorBrush = NULL;
	hBuiltInStateImageList = NULL;

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	if(RunTimeHelper::IsVista()) {
		CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pWICImagingFactory));
		ATLASSUME(pWICImagingFactory);
	}
}
#pragma warning(pop)


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ExplorerTreeView::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IExplorerTreeView, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP ExplorerTreeView::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 35 VT_I4 properties...
	pSize->QuadPart += 35 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 19 VT_BOOL properties...
	pSize->QuadPart += 19 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));

	// ...and 2 VT_DISPATCH properties
	pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.font.pFontDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP ExplorerTreeView::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	// NOTE: We neither support legacy streams nor streams with a version < 0x0102.

	HRESULT hr = S_OK;
	LONG signature = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x02020202/*4x AppID*/) {
		return E_FAIL;
	}
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
		return hr;
	}
	if(version < 0x0102 || version > 0x0105) {
		return E_FAIL;
	}

	DWORD atlVersion;
	if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(version > 0x0103) {
		if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
			return hr;
		}
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = ReadVariantFromStream;

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowLabelEditing = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL alwaysShowSelection = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL blendSelectedItemsIcons = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG caretChangedDelayTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG dragExpandTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG dragScrollTimeBase = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR editBackColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR editForeColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG editHoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	IMEModeConstants editIMEMode = static_cast<IMEModeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL favoritesStyle = propertyValue.boolVal;

	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IFontDisp> pFont = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pFont)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR foreColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL fullRowSelect = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR groupBoxColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	HotTrackingConstants hotTracking = static_cast<HotTrackingConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	IMEModeConstants IMEMode = static_cast<IMEModeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS indent = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR insertMarkColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ItemBoundingBoxDefinitionConstants itemBoundingBoxDefinition = static_cast<ItemBoundingBoxDefinitionConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS itemHeight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DWORD itemBorders = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR lineColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LineStyleConstants lineStyle = static_cast<LineStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG maxScrollTime = propertyValue.lVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processContextMenuKeys = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL registerForOLEDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ScrollBarsConstants scrollBars = static_cast<ScrollBarsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showStateImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showToolTips = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	SingleExpandConstants singleExpand = static_cast<SingleExpandConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	SortOrderConstants sortOrder = static_cast<SortOrderConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	TreeViewStyleConstants treeViewStyle = static_cast<TreeViewStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL autoHScroll = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG autoHScrollPixelsPerSecond = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG autoHScrollRedrawInterval = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BuiltInStateImagesConstants builtInStateImages = static_cast<BuiltInStateImagesConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL drawImagesAsynchronously = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL fadeExpandos = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL indentStateImages = propertyValue.boolVal;

	OLEDragImageStyleConstants oleDragImageStyle = odistClassic;
	VARIANT_BOOL richToolTips = VARIANT_FALSE;
	if(version >= 0x0104) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		oleDragImageStyle = static_cast<OLEDragImageStyleConstants>(propertyValue.lVal);
		if(version >= 0x0105) {
			if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
				return hr;
			}
			richToolTips = propertyValue.boolVal;
		}
	}


	hr = put_AllowDragDrop(allowDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AllowLabelEditing(allowLabelEditing);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AlwaysShowSelection(alwaysShowSelection);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoHScroll(autoHScroll);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoHScrollPixelsPerSecond(autoHScrollPixelsPerSecond);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoHScrollRedrawInterval(autoHScrollRedrawInterval);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BlendSelectedItemsIcons(blendSelectedItemsIcons);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BuiltInStateImages(builtInStateImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CaretChangedDelayTime(caretChangedDelayTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DragExpandTime(dragExpandTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DragScrollTimeBase(dragScrollTimeBase);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DrawImagesAsynchronously(drawImagesAsynchronously);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_EditBackColor(editBackColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_EditForeColor(editForeColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_EditHoverTime(editHoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_EditIMEMode(editIMEMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FadeExpandos(fadeExpandos);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FavoritesStyle(favoritesStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ForeColor(foreColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FullRowSelect(fullRowSelect);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_GroupBoxColor(groupBoxColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HotTracking(hotTracking);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IMEMode(IMEMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Indent(indent);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IndentStateImages(indentStateImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_InsertMarkColor(insertMarkColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemBoundingBoxDefinition(itemBoundingBoxDefinition);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemHeight(itemHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemXBorder(LOWORD(itemBorders));
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemYBorder(HIWORD(itemBorders));
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LineColor(lineColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LineStyle(lineStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MaxScrollTime(maxScrollTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_OLEDragImageStyle(oleDragImageStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessContextMenuKeys(processContextMenuKeys);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RichToolTips(richToolTips);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeft(rightToLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ScrollBars(scrollBars);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowStateImages(showStateImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowToolTips(showToolTips);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SingleExpand(singleExpand);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SortOrder(sortOrder);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TreeViewStyle(treeViewStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x02020202/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0105;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowLabelEditing);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.alwaysShowSelection);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.blendSelectedItemsIcons);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.caretChangedDelayTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.dragExpandTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.dragScrollTimeBase;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.editBackColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.editForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.editHoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.editIMEMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.favoritesStyle);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(hr = properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	VARTYPE vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.foreColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.fullRowSelect);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.groupBoxColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hotTracking;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.IMEMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.indent;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.insertMarkColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.itemBoundingBoxDefinition;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.itemHeight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.itemBorders;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.lineColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.lineStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.maxScrollTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.rightToLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.scrollBars;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showStateImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showToolTips);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.singleExpand;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.sortOrder;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.treeViewStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0103 starts here
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.autoHScroll);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.autoHScrollPixelsPerSecond;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.autoHScrollRedrawInterval;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.builtInStateImages;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.drawImagesAsynchronously);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.fadeExpandos);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.indentStateImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0104 starts here
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.oleDragImageStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0105 starts here
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.richToolTips);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND ExplorerTreeView::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_TREEVIEW_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<ExplorerTreeView>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT ExplorerTreeView::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("ExplorerTreeView ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void ExplorerTreeView::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP ExplorerTreeView::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<ExplorerTreeView>::IOleObject_SetClientSite(pClientSite);

	/* Check whether the container has an ambient font. If it does, clone it; otherwise create our own
	   font object when we hook up a client site. */
	if(!properties.font.pFontDisp) {
		FONTDESC defaultFont = properties.font.GetDefaultFont();
		CComPtr<IFontDisp> pFont;
		if(FAILED(GetAmbientFontDisp(&pFont))) {
			// use the default font
			OleCreateFontIndirect(&defaultFont, IID_IFontDisp, reinterpret_cast<LPVOID*>(&pFont));
		}
		put_Font(pFont);
	}

	return hr;
}

STDMETHODIMP ExplorerTreeView::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL ExplorerTreeView::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
{
	if((pMessage->message >= WM_KEYFIRST) && (pMessage->message <= WM_KEYLAST)) {
		LRESULT dialogCode = SendMessage(pMessage->hwnd, WM_GETDLGCODE, 0, 0);
		//ATLASSERT(dialogCode == (DLGC_WANTCHARS | DLGC_WANTARROWS | (pMessage->hwnd == containedEdit.m_hWnd ? DLGC_HASSETSEL : 0)));
		if(dialogCode & DLGC_WANTTAB) {
			if(pMessage->wParam == VK_TAB) {
				hReturnValue = S_FALSE;
				return TRUE;
			}
		}
		switch(pMessage->wParam) {
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
			case VK_HOME:
			case VK_END:
			case VK_NEXT:
			case VK_PRIOR:
				if(dialogCode & DLGC_WANTARROWS) {
					if(!(GetKeyState(VK_SHIFT) & 0x8000) && !(GetKeyState(VK_CONTROL) & 0x8000) && !(GetKeyState(VK_MENU) & 0x8000)) {
						SendMessage(pMessage->hwnd, pMessage->message, pMessage->wParam, pMessage->lParam);
						hReturnValue = S_OK;
						return TRUE;
					}
				}
				break;
		}
	}
	return CComControl<ExplorerTreeView>::PreTranslateAccelerator(pMessage, hReturnValue);
}

HIMAGELIST ExplorerTreeView::CreateLegacyDragImage(HTREEITEM hItem, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle)
{
	/********************************************************************************************************
	 * Known problems:                                                                                      *
	 * - We assume that the selection background goes over icon and label for all themes.                   *
	 * - We use hardcoded margins for themed items.                                                         *
	 ********************************************************************************************************/

	HIMAGELIST hSourceImageList = cachedSettings.hImageList;
	SIZE imageSize = {cachedSettings.iconWidth, cachedSettings.iconHeight};

	// retrieve item details
	TVITEMEX item = {0};
	item.hItem = hItem;
	item.cchTextMax = MAX_ITEMTEXTLENGTH;
	TCHAR pItemTextBuffer[MAX_ITEMTEXTLENGTH + 1];
	item.pszText = pItemTextBuffer;
	item.mask = TVIF_HANDLE | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_EXPANDEDIMAGE | TVIF_STATE | TVIF_STATEEX | TVIF_TEXT;
	item.stateMask = TVIS_SELECTED | TVIS_EXPANDED | TVIS_EXPANDPARTIAL | TVIS_OVERLAYMASK;
	SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item));
	if(!item.pszText) {
		item.pszText = TEXT("");
	}
	BOOL itemIsSelected = ((item.state & TVIS_SELECTED) == TVIS_SELECTED);
	BOOL itemIsFocused = (hItem == reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0)));
	BOOL itemIsExpanded = (((item.state & TVIS_EXPANDED) == TVIS_EXPANDED) || ((item.state & TVIS_EXPANDPARTIAL) == TVIS_EXPANDPARTIAL));
	BOOL itemIsGrayed = ((item.uStateEx & TVIS_EX_DISABLED) == TVIS_EX_DISABLED);
	int iconToDraw = (itemIsFocused ? item.iSelectedImage : (itemIsExpanded ? item.iExpandedImage : item.iImage));

	// retrieve window details
	BOOL hasFocus = (GetFocus() == *this);
	DWORD style = GetExStyle();
	DWORD textDrawStyle = DT_EDITCONTROL | DT_NOPREFIX | DT_SINGLELINE | DT_VCENTER;
	if(style & WS_EX_RTLREADING) {
		textDrawStyle |= DT_RTLREADING;
	}
	BOOL layoutRTL = ((style & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
	BOOL fullRowSelect = ((GetStyle() & TVS_FULLROWSELECT) == TVS_FULLROWSELECT);

	// create the DCs we'll draw into
	HDC hCompatibleDC = GetDC();
	CDC memoryDC;
	memoryDC.CreateCompatibleDC(hCompatibleDC);
	CDC maskMemoryDC;
	maskMemoryDC.CreateCompatibleDC(hCompatibleDC);

	// prepare themed item drawing
	CTheme themingEngine;
	MARGINS selectionMargins = {0};
	BOOL themedTreeItems = FALSE;
	int themeState = TREIS_NORMAL;
	if(itemIsSelected) {
		if(hasFocus) {
			themeState = TREIS_SELECTED;
		} else {
			themeState = TREIS_SELECTEDNOTFOCUS;
		}
	}
	if(flags.usingThemes) {
		themingEngine.OpenThemeData(*this, VSCLASS_TREEVIEW);
		/* We use TREIS_SELECTED here, because it's more likely this one is defined and we don't want a mixture
		   of themed and non-themed items in drag images. What we're doing with TREIS_NORMAL should work
		   regardless whether TREIS_NORMAL is defined. */
		themedTreeItems = themingEngine.IsThemePartDefined(TVP_TREEITEM, TREIS_SELECTED/*themeState*/);
		if(themedTreeItems) {
			/* FIXME: Retrieve the proper margins. Those returned by GetThemeBackgroundContentRect() are wrong
			          (1 pixel in each direction instead of 2 pixels to left and right and 0 to top and bottom). */
			selectionMargins.cxLeftWidth = 2;
			selectionMargins.cxRightWidth = 2;
			selectionMargins.cyTopHeight = 0;
			selectionMargins.cyBottomHeight = 0;
		}
	}

	// calculate the bounding rectangles of the various item parts
	CRect itemBoundingRect;
	CRect selectionBoundingRect;
	CRect labelBoundingRect;
	CRect iconBoundingRect;

	*reinterpret_cast<HTREEITEM*>(&labelBoundingRect) = hItem;
	SendMessage(TVM_GETITEMRECT, TRUE, reinterpret_cast<LPARAM>(&labelBoundingRect));
	if(hSourceImageList) {
		if(IsComctl32Version610OrNewer() && (cachedSettings.itemHeight < cachedSettings.iconHeight)) {
			iconBoundingRect.top = labelBoundingRect.top;
		} else {
			iconBoundingRect.top = labelBoundingRect.top + (cachedSettings.itemHeight - cachedSettings.iconHeight) / 2;
		}
		iconBoundingRect.bottom = iconBoundingRect.top + cachedSettings.iconHeight;
		iconBoundingRect.right = labelBoundingRect.left - TV_TEXTINDENT;
		iconBoundingRect.left = iconBoundingRect.right - imageSize.cx;
	}
	if(fullRowSelect) {
		*reinterpret_cast<HTREEITEM*>(&itemBoundingRect) = hItem;
		SendMessage(TVM_GETITEMRECT, TRUE, reinterpret_cast<LPARAM>(&itemBoundingRect));
		selectionBoundingRect = itemBoundingRect;
	} else {
		itemBoundingRect = labelBoundingRect;
		if(hSourceImageList) {
			itemBoundingRect.left = iconBoundingRect.left;
		}
		if(themedTreeItems) {
			selectionBoundingRect = itemBoundingRect;
		} else {
			selectionBoundingRect = labelBoundingRect;
		}
		selectionBoundingRect.left -= selectionMargins.cxLeftWidth;
		selectionBoundingRect.right += selectionMargins.cxRightWidth;
	}
	itemBoundingRect.UnionRect(&itemBoundingRect, &selectionBoundingRect);
	if(hSourceImageList) {
		itemBoundingRect.UnionRect(&itemBoundingRect, &iconBoundingRect);
	}
	if(pBoundingRectangle) {
		*pBoundingRectangle = itemBoundingRect;
	}

	// calculate drag image size and upper-left corner
	SIZE dragImageSize = {0};
	if(pUpperLeftPoint) {
		pUpperLeftPoint->x = itemBoundingRect.left;
		pUpperLeftPoint->y = itemBoundingRect.top;
	}
	dragImageSize.cx = itemBoundingRect.Width();
	dragImageSize.cy = itemBoundingRect.Height();

	// offset RECTs
	SIZE offset = {0};
	offset.cx = itemBoundingRect.left;
	offset.cy = itemBoundingRect.top;
	labelBoundingRect.OffsetRect(-offset.cx, -offset.cy);
	iconBoundingRect.OffsetRect(-offset.cx, -offset.cy);
	selectionBoundingRect.OffsetRect(-offset.cx, -offset.cy);
	itemBoundingRect.OffsetRect(-offset.cx, -offset.cy);

	// setup the DCs we'll draw into
	if(itemIsSelected) {
		memoryDC.SetBkColor(GetSysColor(hasFocus ? COLOR_HIGHLIGHT : COLOR_BTNFACE));
		memoryDC.SetTextColor(GetSysColor(hasFocus ? COLOR_HIGHLIGHTTEXT : COLOR_BTNTEXT));
	} else {
		memoryDC.SetBkColor(GetSysColor(COLOR_WINDOW));
		memoryDC.SetTextColor(GetSysColor(COLOR_WINDOWTEXT));
		memoryDC.SetBkMode(TRANSPARENT);
	}

	// create drag image bitmap
	/* NOTE: We prefer creating 32bpp drag images, because this improves performance of
	         TreeViewItemContainer::CreateDragImage(). */
	BOOL doAlphaChannelProcessing = RunTimeHelper::IsCommCtrl6();
	BITMAPINFO bitmapInfo = {0};
	if(doAlphaChannelProcessing) {
		bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bitmapInfo.bmiHeader.biWidth = dragImageSize.cx;
		bitmapInfo.bmiHeader.biHeight = -dragImageSize.cy;
		bitmapInfo.bmiHeader.biPlanes = 1;
		bitmapInfo.bmiHeader.biBitCount = 32;
		bitmapInfo.bmiHeader.biCompression = BI_RGB;
	}
	CBitmap dragImage;
	LPRGBQUAD pDragImageBits = NULL;
	if(doAlphaChannelProcessing) {
		dragImage.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
	} else {
		dragImage.CreateCompatibleBitmap(hCompatibleDC, dragImageSize.cx, dragImageSize.cy);
	}
	HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(dragImage);
	CBitmap dragImageMask;
	dragImageMask.CreateBitmap(dragImageSize.cx, dragImageSize.cy, 1, 1, NULL);
	HBITMAP hPreviousBitmapMask = maskMemoryDC.SelectBitmap(dragImageMask);

	// initialize the bitmap
	RECT rc = itemBoundingRect;
	if(doAlphaChannelProcessing && pDragImageBits) {
		// we need a transparent background
		LPRGBQUAD pPixel = pDragImageBits;
		for(int y = 0; y < dragImageSize.cy; ++y) {
			for(int x = 0; x < dragImageSize.cx; ++x, ++pPixel) {
				pPixel->rgbRed = 0xFF;
				pPixel->rgbGreen = 0xFF;
				pPixel->rgbBlue = 0xFF;
				pPixel->rgbReserved = 0x00;
			}
		}
	} else {
		memoryDC.FillRect(&rc, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
	}
	maskMemoryDC.FillRect(&rc, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));

	CFontHandle font = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
	HFONT hPreviousFont = NULL;
	if(!font.IsNull()) {
		hPreviousFont = memoryDC.SelectFont(font);
	}

	// draw the selection area's background
	RECT focusRect = selectionBoundingRect;
	if(itemIsSelected) {
		if(themedTreeItems) {
			themingEngine.DrawThemeBackground(memoryDC, TVP_TREEITEM, themeState, &focusRect, NULL);
		} else {
			memoryDC.FillRect(&focusRect, (hasFocus ? COLOR_HIGHLIGHT : COLOR_BTNFACE));
		}
		maskMemoryDC.FillRect(&focusRect, static_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
	}

	// draw the icon
	if(hSourceImageList) {
		ImageList_DrawEx(hSourceImageList, iconToDraw, memoryDC, iconBoundingRect.left, iconBoundingRect.top, imageSize.cx, imageSize.cy, CLR_NONE, CLR_NONE, (itemIsGrayed ? ILD_SELECTED : ILD_NORMAL) | (item.state & TVIS_OVERLAYMASK));
		if(itemIsSelected && (fullRowSelect || themedTreeItems)) {
			maskMemoryDC.FillRect(&iconBoundingRect, static_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
		} else {
			ImageList_Draw(hSourceImageList, iconToDraw, maskMemoryDC, iconBoundingRect.left, iconBoundingRect.top, ILD_MASK | (item.state & TVIS_OVERLAYMASK));
		}
	}

	// draw the text
	rc = labelBoundingRect;
	InflateRect(&rc, -2, 0);
	if(themedTreeItems) {
		CT2W converter(item.pszText);
		LPWSTR pLabelText = converter;
		themingEngine.DrawThemeText(memoryDC, TVP_TREEITEM, themeState, pLabelText, lstrlenW(pLabelText), textDrawStyle, 0, &rc);
	} else {
		memoryDC.DrawText(item.pszText, lstrlen(item.pszText), &rc, textDrawStyle);
	}
	if(!itemIsSelected) {
		COLORREF bkColor = memoryDC.GetBkColor();
		for(int y = rc.top; y <= rc.bottom; ++y) {
			for(int x = rc.left; x <= rc.right; ++x) {
				if(memoryDC.GetPixel(x, y) != bkColor) {
					maskMemoryDC.SetPixelV(x, y, 0x00000000);
				}
			}
		}
	}

	if(!themedTreeItems) {
		// draw the focus rectangle
		if(hasFocus && itemIsFocused) {
			if((SendMessage(WM_QUERYUISTATE, 0, 0) & UISF_HIDEFOCUS) == 0) {
				memoryDC.DrawFocusRect(&focusRect);
				maskMemoryDC.FrameRect(&focusRect, static_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
			}
		}
	}

	if(doAlphaChannelProcessing && pDragImageBits) {
		// correct the alpha channel
		LPRGBQUAD pPixel = pDragImageBits;
		POINT pt;
		for(pt.y = 0; pt.y < dragImageSize.cy; ++pt.y) {
			for(pt.x = 0; pt.x < dragImageSize.cx; ++pt.x, ++pPixel) {
				if(layoutRTL) {
					// we're working on raw data, so we've to handle WS_EX_LAYOUTRTL ourselves
					POINT pt2 = pt;
					pt2.x = dragImageSize.cx - pt.x - 1;
					if(maskMemoryDC.GetPixel(pt2.x, pt2.y) == 0x00000000) {
						if(themedTreeItems) {
							if(itemIsSelected) {
								if((pPixel->rgbReserved != 0x00) || labelBoundingRect.PtInRect(pt2)) {
									pPixel->rgbReserved = 0xFF;
								}
							} else if(labelBoundingRect.PtInRect(pt2)) {
								if(pPixel->rgbReserved == 0x00) {
									pPixel->rgbReserved = 0xFF;
								}
							}
						} else {
							// items are not themed
							if(itemIsSelected) {
								if((pPixel->rgbReserved == 0x00) || selectionBoundingRect.PtInRect(pt2)) {
									pPixel->rgbReserved = 0xFF;
								}
							} else {
								if(pPixel->rgbReserved == 0x00) {
									pPixel->rgbReserved = 0xFF;
								}
							}
						}
					}
				} else {
					// layout is left to right
					if(maskMemoryDC.GetPixel(pt.x, pt.y) == 0x00000000) {
						if(themedTreeItems) {
							if(itemIsSelected) {
								if((pPixel->rgbReserved != 0x00) || labelBoundingRect.PtInRect(pt)) {
									pPixel->rgbReserved = 0xFF;
								}
							} else if(labelBoundingRect.PtInRect(pt)) {
								if(pPixel->rgbReserved == 0x00) {
									pPixel->rgbReserved = 0xFF;
								}
							}
						} else {
							// items are not themed
							if(itemIsSelected) {
								if((pPixel->rgbReserved == 0x00) || selectionBoundingRect.PtInRect(pt)) {
									pPixel->rgbReserved = 0xFF;
								}
							} else {
								if(pPixel->rgbReserved == 0x00) {
									pPixel->rgbReserved = 0xFF;
								}
							}
						}
					}
				}
			}
		}
	}

	memoryDC.SelectBitmap(hPreviousBitmap);
	maskMemoryDC.SelectBitmap(hPreviousBitmapMask);
	if(hPreviousFont) {
		memoryDC.SelectFont(hPreviousFont);
	}

	// create the imagelist
	HIMAGELIST hDragImageList = ImageList_Create(dragImageSize.cx, dragImageSize.cy, (RunTimeHelper::IsCommCtrl6() ? ILC_COLOR32 : ILC_COLOR24) | ILC_MASK, 1, 0);
	ImageList_SetBkColor(hDragImageList, CLR_NONE);
	ImageList_Add(hDragImageList, dragImage, dragImageMask);

	ReleaseDC(hCompatibleDC);

	return hDragImageList;
}

BOOL ExplorerTreeView::CreateLegacyOLEDragImage(ITreeViewItemContainer* pItems, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(pItems);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;

	// use a normal legacy drag image as base
	OLE_HANDLE h = NULL;
	OLE_XPOS_PIXELS xUpperLeft = 0;
	OLE_YPOS_PIXELS yUpperLeft = 0;
	pItems->CreateDragImage(&xUpperLeft, &yUpperLeft, &h);
	if(h) {
		HIMAGELIST hImageList = static_cast<HIMAGELIST>(LongToHandle(h));

		// retrieve the drag image's size
		int bitmapHeight;
		int bitmapWidth;
		ImageList_GetIconSize(hImageList, &bitmapWidth, &bitmapHeight);
		pDragImage->sizeDragImage.cx = bitmapWidth;
		pDragImage->sizeDragImage.cy = bitmapHeight;

		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		pDragImage->hbmpDragImage = NULL;

		if(RunTimeHelper::IsCommCtrl6()) {
			// handle alpha channel
			IImageList* pImgLst = NULL;
			HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
			if(hMod) {
				typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
				HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
				if(pfnHIMAGELIST_QueryInterface) {
					pfnHIMAGELIST_QueryInterface(hImageList, IID_IImageList, reinterpret_cast<LPVOID*>(&pImgLst));
				}
				FreeLibrary(hMod);
			}
			if(!pImgLst) {
				pImgLst = reinterpret_cast<IImageList*>(hImageList);
				pImgLst->AddRef();
			}
			ATLASSUME(pImgLst);

			DWORD imageFlags = 0;
			pImgLst->GetItemFlags(0, &imageFlags);
			if(imageFlags & ILIF_ALPHA) {
				// the drag image makes use of the alpha channel
				IMAGEINFO imageInfo = {0};
				ImageList_GetImageInfo(hImageList, 0, &imageInfo);

				// fetch raw data
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
				bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;
				LPRGBQUAD pSourceBits = reinterpret_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * sizeof(RGBQUAD)));
				GetDIBits(memoryDC, imageInfo.hbmImage, 0, pDragImage->sizeDragImage.cy, pSourceBits, &bitmapInfo, DIB_RGB_COLORS);
				// create target bitmap
				LPRGBQUAD pDragImageBits = NULL;
				pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
				ATLASSERT(pDragImageBits);
				pDragImage->crColorKey = 0xFFFFFFFF;

				if(pDragImageBits) {
					// transfer raw data
					CopyMemory(pDragImageBits, pSourceBits, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * 4);
				}

				// clean up
				HeapFree(GetProcessHeap(), 0, pSourceBits);
				DeleteObject(imageInfo.hbmImage);
				DeleteObject(imageInfo.hbmMask);
			}
			pImgLst->Release();
		}

		if(!pDragImage->hbmpDragImage) {
			// fallback mode
			memoryDC.SetBkMode(TRANSPARENT);

			// create target bitmap
			HDC hCompatibleDC = ::GetDC(HWND_DESKTOP);
			pDragImage->hbmpDragImage = CreateCompatibleBitmap(hCompatibleDC, bitmapWidth, bitmapHeight);
			::ReleaseDC(HWND_DESKTOP, hCompatibleDC);
			HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(pDragImage->hbmpDragImage);

			// draw target bitmap
			pDragImage->crColorKey = RGB(0xF4, 0x00, 0x00);
			CBrush backroundBrush;
			backroundBrush.CreateSolidBrush(pDragImage->crColorKey);
			memoryDC.FillRect(CRect(0, 0, bitmapWidth, bitmapHeight), backroundBrush);
			ImageList_Draw(hImageList, 0, memoryDC, 0, 0, ILD_NORMAL);

			// clean up
			memoryDC.SelectBitmap(hPreviousBitmap);
		}

		ImageList_Destroy(hImageList);

		if(pDragImage->hbmpDragImage) {
			// retrieve the offset
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			if(GetExStyle() & WS_EX_LAYOUTRTL) {
				pDragImage->ptOffset.x = xUpperLeft + pDragImage->sizeDragImage.cx - mousePosition.x;
			} else {
				pDragImage->ptOffset.x = mousePosition.x - xUpperLeft;
			}
			pDragImage->ptOffset.y = mousePosition.y - yUpperLeft;

			succeeded = TRUE;
		}
	}

	return succeeded;
}

BOOL ExplorerTreeView::CreateVistaOLEDragImage(ITreeViewItemContainer* pItems, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(pItems);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;

	CTheme themingEngine;
	themingEngine.OpenThemeData(NULL, VSCLASS_DRAGDROP);
	if(themingEngine.IsThemeNull()) {
		// FIXME: What should we do here?
		ATLASSERT(FALSE && "Current theme does not define the \"DragDrop\" class.");
	} else {
		// retrieve the drag image's size
		CDC memoryDC;
		memoryDC.CreateCompatibleDC();

		themingEngine.GetThemePartSize(memoryDC, DD_IMAGEBG, 1, NULL, TS_TRUE, &pDragImage->sizeDragImage);
		MARGINS margins = {0};
		themingEngine.GetThemeMargins(memoryDC, DD_IMAGEBG, 1, TMT_CONTENTMARGINS, NULL, &margins);
		pDragImage->sizeDragImage.cx -= margins.cxLeftWidth + margins.cxRightWidth;
		pDragImage->sizeDragImage.cy -= margins.cyTopHeight + margins.cyBottomHeight;
	}

	ATLASSERT(pDragImage->sizeDragImage.cx > 0);
	ATLASSERT(pDragImage->sizeDragImage.cy > 0);

	// create target bitmap
	BITMAPINFO bitmapInfo = {0};
	bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
	bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
	bitmapInfo.bmiHeader.biPlanes = 1;
	bitmapInfo.bmiHeader.biBitCount = 32;
	bitmapInfo.bmiHeader.biCompression = BI_RGB;
	LPRGBQUAD pDragImageBits = NULL;
	pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);

	HIMAGELIST hSourceImageList = (cachedSettings.hHighResImageList ? cachedSettings.hHighResImageList : cachedSettings.hImageList);
	if(!hSourceImageList) {
		// report success, although we've an empty drag image
		return TRUE;
	}

	IImageList2* pImgLst = NULL;
	HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
	if(hMod) {
		typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
		HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
		if(pfnHIMAGELIST_QueryInterface) {
			pfnHIMAGELIST_QueryInterface(hSourceImageList, IID_IImageList2, reinterpret_cast<LPVOID*>(&pImgLst));
		}
		FreeLibrary(hMod);
	}
	if(!pImgLst) {
		IImageList* p = reinterpret_cast<IImageList*>(hSourceImageList);
		p->QueryInterface(IID_IImageList2, reinterpret_cast<LPVOID*>(&pImgLst));
	}
	ATLASSUME(pImgLst);

	if(pImgLst) {
		LONG numberOfItems = 0;
		pItems->Count(&numberOfItems);
		ATLASSERT(numberOfItems > 0);
		// don't display more than 5 (10) thumbnails
		numberOfItems = min(numberOfItems, (hSourceImageList == cachedSettings.hHighResImageList ? 5 : 10));

		CComPtr<IUnknown> pUnknownEnum = NULL;
		pItems->get__NewEnum(&pUnknownEnum);
		CComQIPtr<IEnumVARIANT> pEnum = pUnknownEnum;
		ATLASSUME(pEnum);
		if(pEnum) {
			int cx = 0;
			int cy = 0;
			pImgLst->GetIconSize(&cx, &cy);
			SIZE thumbnailSize;
			thumbnailSize.cy = pDragImage->sizeDragImage.cy - 3 * (numberOfItems - 1);
			if(thumbnailSize.cy < 8) {
				// don't get smaller than 8x8 thumbnails
				numberOfItems = (pDragImage->sizeDragImage.cy - 8) / 3 + 1;
				thumbnailSize.cy = pDragImage->sizeDragImage.cy - 3 * (numberOfItems - 1);
			}
			thumbnailSize.cx = thumbnailSize.cy;
			int thumbnailBufferSize = thumbnailSize.cx * thumbnailSize.cy * sizeof(RGBQUAD);
			LPRGBQUAD pThumbnailBits = reinterpret_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, thumbnailBufferSize));
			ATLASSERT(pThumbnailBits);
			if(pThumbnailBits) {
				// iterate over the dragged items
				VARIANT v;
				int i = 0;
				CComPtr<ITreeViewItem> pItem = NULL;
				while(pEnum->Next(1, &v, NULL) == S_OK) {
					if(v.vt == VT_DISPATCH) {
						v.pdispVal->QueryInterface(IID_ITreeViewItem, reinterpret_cast<LPVOID*>(&pItem));
						ATLASSUME(pItem);
						if(pItem) {
							// get the item's icon
							LONG icon = 0;
							pItem->get_IconIndex(&icon);

							pImgLst->ForceImagePresent(icon, ILFIP_ALWAYS);
							HICON hIcon = NULL;
							pImgLst->GetIcon(icon, ILD_TRANSPARENT, &hIcon);
							ATLASSERT(hIcon);
							if(hIcon) {
								// finally create the thumbnail
								ZeroMemory(pThumbnailBits, thumbnailBufferSize);
								HRESULT hr = CreateThumbnail(hIcon, thumbnailSize, pThumbnailBits, TRUE);
								DestroyIcon(hIcon);
								if(FAILED(hr)) {
									pItem = NULL;
									VariantClear(&v);
									break;
								}

								// add the thumbail to the drag image keeping the alpha channel intact
								if(i == 0) {
									LPRGBQUAD pDragImagePixel = pDragImageBits;
									LPRGBQUAD pThumbnailPixel = pThumbnailBits;
									for(int scanline = 0; scanline < thumbnailSize.cy; ++scanline, pDragImagePixel += pDragImage->sizeDragImage.cx, pThumbnailPixel += thumbnailSize.cx) {
										CopyMemory(pDragImagePixel, pThumbnailPixel, thumbnailSize.cx * sizeof(RGBQUAD));
									}
								} else {
									LPRGBQUAD pDragImagePixel = pDragImageBits;
									LPRGBQUAD pThumbnailPixel = pThumbnailBits;
									pDragImagePixel += 3 * i * pDragImage->sizeDragImage.cx;
									for(int scanline = 0; scanline < thumbnailSize.cy; ++scanline, pDragImagePixel += pDragImage->sizeDragImage.cx) {
										LPRGBQUAD p = pDragImagePixel + 2 * i;
										for(int x = 0; x < thumbnailSize.cx; ++x, ++p, ++pThumbnailPixel) {
											// merge the pixels
											p->rgbRed = pThumbnailPixel->rgbRed * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbRed / 0xFF;
											p->rgbGreen = pThumbnailPixel->rgbGreen * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbGreen / 0xFF;
											p->rgbBlue = pThumbnailPixel->rgbBlue * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbBlue / 0xFF;
											p->rgbReserved = pThumbnailPixel->rgbReserved + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbReserved / 0xFF;
										}
									}
								}
							}

							++i;
							pItem = NULL;
							if(i == numberOfItems) {
								VariantClear(&v);
								break;
							}
						}
					}
					VariantClear(&v);
				}
				HeapFree(GetProcessHeap(), 0, pThumbnailBits);
				succeeded = TRUE;
			}
		}

		pImgLst->Release();
	}

	return succeeded;
}

HRESULT ExplorerTreeView::SortSubItems(HTREEITEM hParentItem, SortByConstants firstCriterion/* = sobShell*/, SortByConstants secondCriterion/* = sobText*/, SortByConstants thirdCriterion/* = sobNone*/, SortByConstants fourthCriterion/* = sobNone*/, SortByConstants fifthCriterion/* = sobNone*/, BOOL recurse/* = FALSE*/, BOOL caseSensitive/* = FALSE*/)
{
	ATLASSERT(IsWindow());

	if(firstCriterion == sobNone) {
		// nothing to do...
		return S_OK;
	}
	#ifndef INCLUDESHELLBROWSERINTERFACE
		if((firstCriterion == sobShell) && (secondCriterion == sobNone)) {
			// nothing to do...
			return S_OK;
		}
	#endif
	if(!HasSubItems(hParentItem)) {
		// nothing to do...
		return S_OK;
	}

	BOOL useCallback = TRUE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(!shellBrowserInterface.pInternalMessageListener) {
	#endif
		if(((firstCriterion == sobShell) && (secondCriterion == sobText) && (thirdCriterion == sobNone)) || ((firstCriterion == sobText) && (secondCriterion == sobNone)) && (caseSensitive == VARIANT_FALSE)) {
			if(properties.sortOrder == soAscending) {
				useCallback = FALSE;
				if(!SendMessage(TVM_SORTCHILDREN, 0, reinterpret_cast<LPARAM>(hParentItem))) {
					return E_FAIL;
				}
			}
		}
	#ifdef INCLUDESHELLBROWSERINTERFACE
		}
	#endif

	if(useCallback) {
		sortingSettings.caseSensitive = caseSensitive;
		sortingSettings.sortingCriteria[0] = firstCriterion;
		sortingSettings.sortingCriteria[1] = secondCriterion;
		sortingSettings.sortingCriteria[2] = thirdCriterion;
		sortingSettings.sortingCriteria[3] = fourthCriterion;
		sortingSettings.sortingCriteria[4] = fifthCriterion;
		sortingSettings.localeID = properties.locale;
		sortingSettings.flagsForVarFromStr = properties.textParsingFlagsForVarFromStr;
		sortingSettings.flagsForCompareString = properties.textParsingFlagsForCompareString;

		TVSORTCB details = {0};
		details.hParent = hParentItem;
		details.lParam = reinterpret_cast<LPARAM>(static_cast<IItemComparator*>(this));
		details.lpfnCompare = ::CompareItems;
		if(!SendMessage(TVM_SORTCHILDRENCB, 0, reinterpret_cast<LPARAM>(&details))) {
			return E_FAIL;
		}
	}

	if(recurse) {
		HTREEITEM hItem = NULL;
		if(hParentItem == TVI_ROOT) {
			hItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_ROOT), 0));
		} else {
			hItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(hParentItem)));
		}
		HRESULT hr = S_OK;
		while(hItem) {
			if(FAILED(hr = SortSubItems(hItem, firstCriterion, secondCriterion, thirdCriterion, fourthCriterion, fifthCriterion, recurse, caseSensitive))) {
				return hr;
			}
			hItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hItem)));
		}
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP ExplorerTreeView::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	BOOL callDropTargetHelper = TRUE;
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);

	if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
		if(dragDropStatus.useItemCountLabelHack) {
			dragDropStatus.pDropTargetHelper->DragLeave();
			dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
			dragDropStatus.useItemCountLabelHack = FALSE;
		}
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::DragLeave(void)
{
	BOOL callDropTargetHelper = TRUE;
	Raise_OLEDragLeave(&callDropTargetHelper);

	if(dragDropStatus.pDropTargetHelper) {
		if(callDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->DragLeave();
		}
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	BOOL callDropTargetHelper = TRUE;
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition, &callDropTargetHelper);

	if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, newDropDescription);
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	BOOL callDropTargetHelper = TRUE;
	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);

	if(dragDropStatus.pDropTargetHelper) {
		if(callDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		}
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSource
STDMETHODIMP ExplorerTreeView::GiveFeedback(DWORD effect)
{
	VARIANT_BOOL useDefaultCursors = VARIANT_TRUE;
	//if(flags.usingThemes && RunTimeHelper::IsVista()) {
		ATLASSUME(dragDropStatus.pSourceDataObject);

		BOOL isShowingLayered = FALSE;
		FORMATETC format = {0};
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("IsShowingLayered")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		STGMEDIUM medium = {0};
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				isShowingLayered = *reinterpret_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}
		BOOL usingDefaultDragImage = FALSE;
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				usingDefaultDragImage = *reinterpret_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}

		if(isShowingLayered && properties.oleDragImageStyle != odistClassic) {
			SetCursor(static_cast<HCURSOR>(LoadImage(NULL, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED)));
			useDefaultCursors = VARIANT_FALSE;
		}
		if(usingDefaultDragImage) {
			// this will make drop descriptions work
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("DragWindow")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
				if(medium.hGlobal) {
					// WM_USER + 1 (with wParam = 0 and lParam = 0) hides the drag image
					#define WM_SETDROPEFFECT				WM_USER + 2     // (wParam = DCID_*, lParam = 0)
					#define DDWM_UPDATEWINDOW				WM_USER + 3     // (wParam = 0, lParam = 0)
					typedef enum DROPEFFECTS
					{
						DCID_NULL = 0,
						DCID_NO = 1,
						DCID_MOVE = 2,
						DCID_COPY = 3,
						DCID_LINK = 4,
						DCID_MAX = 5
					} DROPEFFECTS;

					HWND hWndDragWindow = *reinterpret_cast<HWND*>(GlobalLock(medium.hGlobal));
					GlobalUnlock(medium.hGlobal);

					DROPEFFECTS dropEffect = DCID_NULL;
					switch(effect) {
						case DROPEFFECT_NONE:
							dropEffect = DCID_NO;
							break;
						case DROPEFFECT_COPY:
							dropEffect = DCID_COPY;
							break;
						case DROPEFFECT_MOVE:
							dropEffect = DCID_MOVE;
							break;
						case DROPEFFECT_LINK:
							dropEffect = DCID_LINK;
							break;
					}
					if(::IsWindow(hWndDragWindow)) {
						::PostMessage(hWndDragWindow, WM_SETDROPEFFECT, dropEffect, 0);
					}
				}
				ReleaseStgMedium(&medium);
			}
		}
	//}

	Raise_OLEGiveFeedback(effect, &useDefaultCursors);
	return (useDefaultCursors == VARIANT_FALSE ? S_OK : DRAGDROP_S_USEDEFAULTCURSORS);
}

STDMETHODIMP ExplorerTreeView::QueryContinueDrag(BOOL pressedEscape, DWORD keyState)
{
	HRESULT actionToContinueWith = S_OK;
	if(pressedEscape) {
		actionToContinueWith = DRAGDROP_S_CANCEL;
	} else if(!(keyState & dragDropStatus.draggingMouseButton)) {
		actionToContinueWith = DRAGDROP_S_DROP;
	}
	Raise_OLEQueryContinueDrag(pressedEscape, keyState, &actionToContinueWith);
	return actionToContinueWith;
}
// implementation of IDropSource
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSourceNotify
STDMETHODIMP ExplorerTreeView::DragEnterTarget(HWND hWndTarget)
{
	Raise_OLEDragEnterPotentialTarget(HandleToLong(hWndTarget));
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::DragLeaveTarget(void)
{
	Raise_OLEDragLeavePotentialTarget();
	return S_OK;
}
// implementation of IDropSourceNotify
//////////////////////////////////////////////////////////////////////

#ifdef INCLUDESHELLBROWSERINTERFACE
	//////////////////////////////////////////////////////////////////////
	// implementation of IInternalMessageListener
	HRESULT ExplorerTreeView::ProcessMessage(UINT message, WPARAM wParam, LPARAM lParam)
	{
		switch(message) {
			case EXTVM_ADDITEM:
				return OnInternalAddItem(message, wParam, lParam);
				break;
			case EXTVM_ITEMHANDLETOID:
				return OnInternalItemHandleToID(message, wParam, lParam);
				break;
			case EXTVM_IDTOITEMHANDLE:
				return OnInternalIDToItemHandle(message, wParam, lParam);
				break;
			case EXTVM_GETITEMBYHANDLE:
				return OnInternalGetItemByHandle(message, wParam, lParam);
				break;
			case EXTVM_CREATEITEMCONTAINER:
				return OnInternalCreateItemContainer(message, wParam, lParam);
				break;
			case EXTVM_GETITEMHANDLESFROMVARIANT:
				return OnInternalGetItemHandlesFromVariant(message, wParam, lParam);
				break;
			case EXTVM_SETITEMICONINDEX:
				return OnInternalSetItemIconIndex(message, wParam, lParam);
				break;
			case EXTVM_SETITEMSELECTEDICONINDEX:
				return OnInternalSetItemSelectedIconIndex(message, wParam, lParam);
				break;
			case EXTVM_SETITEMEXPANDEDICONINDEX:
				return OnInternalSetItemExpandedIconIndex(message, wParam, lParam);
				break;
			case EXTVM_HITTEST:
				return OnInternalHitTest(message, wParam, lParam);
				break;
			case EXTVM_SORTITEMS:
				return OnInternalSortItems(message, wParam, lParam);
				break;
			case EXTVM_SETDROPHILITEDITEM:
				return OnInternalSetDropHilitedItem(message, wParam, lParam);
				break;
			case EXTVM_FIREDRAGDROPEVENT:
				return OnInternalFireDragDropEvent(message, wParam, lParam);
				break;
			case EXTVM_OLEDRAG:
				return OnInternalOLEDrag(message, wParam, lParam);
				break;
			case EXTVM_GETIMAGELIST:
				return OnInternalGetImageList(message, wParam, lParam);
				break;
			case EXTVM_SETIMAGELIST:
				return OnInternalSetImageList(message, wParam, lParam);
				break;
		}
		return E_NOTIMPL;
	}
	// implementation of IInternalMessageListener
	//////////////////////////////////////////////////////////////////////
#endif

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP ExplorerTreeView::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
		case PROPCAT_DragDrop:
			*pName = GetResString(IDPC_DRAGDROP).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerTreeView::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_EXTVW_ALWAYSSHOWSELECTION:
		case DISPID_EXTVW_APPEARANCE:
		case DISPID_EXTVW_BLENDSELECTEDITEMSICONS:
		case DISPID_EXTVW_BORDERSTYLE:
		case DISPID_EXTVW_BUILTINSTATEIMAGES:
		case DISPID_EXTVW_FAVORITESSTYLE:
		case DISPID_EXTVW_INDENT:
		case DISPID_EXTVW_INDENTSTATEIMAGES:
		case DISPID_EXTVW_ITEMHEIGHT:
		case DISPID_EXTVW_ITEMXBORDER:
		case DISPID_EXTVW_ITEMYBORDER:
		case DISPID_EXTVW_LINESTYLE:
		case DISPID_EXTVW_MOUSEICON:
		case DISPID_EXTVW_MOUSEPOINTER:
		case DISPID_EXTVW_RICHTOOLTIPS:
		case DISPID_EXTVW_SHOWSTATEIMAGES:
		case DISPID_EXTVW_TREEVIEWSTYLE:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_EXTVW_ALLOWLABELEDITING:
		case DISPID_EXTVW_AUTOHSCROLL:
		case DISPID_EXTVW_AUTOHSCROLLPIXELSPERSECOND:
		case DISPID_EXTVW_AUTOHSCROLLREDRAWINTERVAL:
		case DISPID_EXTVW_CARETCHANGEDDELAYTIME:
		case DISPID_EXTVW_CARETITEM:
		case DISPID_EXTVW_DISABLEDEVENTS:
		case DISPID_EXTVW_DONTREDRAW:
		case DISPID_EXTVW_DRAWIMAGESASYNCHRONOUSLY:
		case DISPID_EXTVW_EDITHOVERTIME:
		case DISPID_EXTVW_EDITIMEMODE:
		case DISPID_EXTVW_FADEEXPANDOS:
		case DISPID_EXTVW_FULLROWSELECT:
		case DISPID_EXTVW_HOTTRACKING:
		case DISPID_EXTVW_HOVERTIME:
		case DISPID_EXTVW_IMEMODE:
		case DISPID_EXTVW_INCREMENTALSEARCHSTRING:
		case DISPID_EXTVW_ITEMBOUNDINGBOXDEFINITION:
		case DISPID_EXTVW_MAXSCROLLTIME:
		case DISPID_EXTVW_PROCESSCONTEXTMENUKEYS:
		case DISPID_EXTVW_RIGHTTOLEFT:
		case DISPID_EXTVW_SCROLLBARS:
		case DISPID_EXTVW_SHOWTOOLTIPS:
		case DISPID_EXTVW_SINGLEEXPAND:
		case DISPID_EXTVW_SORTORDER:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_EXTVW_BACKCOLOR:
		case DISPID_EXTVW_EDITBACKCOLOR:
		case DISPID_EXTVW_EDITFORECOLOR:
		case DISPID_EXTVW_FORECOLOR:
		case DISPID_EXTVW_GROUPBOXCOLOR:
		case DISPID_EXTVW_INSERTMARKCOLOR:
		case DISPID_EXTVW_LINECOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_EXTVW_APPID:
		case DISPID_EXTVW_APPNAME:
		case DISPID_EXTVW_APPSHORTNAME:
		case DISPID_EXTVW_BUILD:
		case DISPID_EXTVW_CHARSET:
		case DISPID_EXTVW_HDRAGIMAGELIST:
		case DISPID_EXTVW_HIMAGELIST:
		case DISPID_EXTVW_HWND:
		case DISPID_EXTVW_HWNDEDIT:
		case DISPID_EXTVW_HWNDTOOLTIP:
		case DISPID_EXTVW_ISRELEASE:
		case DISPID_EXTVW_PROGRAMMER:
		case DISPID_EXTVW_TESTER:
		case DISPID_EXTVW_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_EXTVW_ALLOWDRAGDROP:
		case DISPID_EXTVW_DRAGEXPANDTIME:
		case DISPID_EXTVW_DRAGGEDITEMS:
		case DISPID_EXTVW_DRAGSCROLLTIMEBASE:
		case DISPID_EXTVW_DROPHILITEDITEM:
		case DISPID_EXTVW_OLEDRAGIMAGESTYLE:
		case DISPID_EXTVW_REGISTERFOROLEDRAGDROP:
		case DISPID_EXTVW_SHOWDRAGIMAGE:
		case DISPID_EXTVW_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_EXTVW_FONT:
		case DISPID_EXTVW_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_EXTVW_FIRSTROOTITEM:
		case DISPID_EXTVW_FIRSTVISIBLEITEM:
		case DISPID_EXTVW_LASTVISIBLEITEM:
		case DISPID_EXTVW_ROOTITEMS:
		case DISPID_EXTVW_TREEITEMS:
			*pCategory = PROPCAT_List;
			return S_OK;
			break;
		case DISPID_EXTVW_ENABLED:
		case DISPID_EXTVW_LOCALE:
		case DISPID_EXTVW_TEXTPARSINGFLAGS:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString ExplorerTreeView::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString ExplorerTreeView::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString ExplorerTreeView::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ExplorerTreeView&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString ExplorerTreeView::GetSpecialThanks(void)
{
	return TEXT("Wine Headquarters");
}

CAtlString ExplorerTreeView::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString ExplorerTreeView::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IItemComparator
int ExplorerTreeView::CompareItems(LONG itemID1, LONG itemID2)
{
	// retrieve the items' handles
	HTREEITEM hItem1 = IDToItemHandle(itemID1);
	HTREEITEM hItem2 = IDToItemHandle(itemID2);
	if(!hItem1 || !hItem2) {
		// there's nothing useful we can do
		return 0;
	}

	int result = 0;

	TVITEMEX item1 = {0};
	TCHAR pItem1TextBuffer[MAX_ITEMTEXTLENGTH + 1];
	TVITEMEX item2 = {0};
	TCHAR pItem2TextBuffer[MAX_ITEMTEXTLENGTH + 1];
	BOOL earlyAbort = FALSE;
	for(int i = 0; i < 5 && !earlyAbort; ++i) {
		if(sortingSettings.sortingCriteria[i] == sobNone) {
			// there's nothing useful we can do
			earlyAbort = TRUE;
			continue;
		} else if(sortingSettings.sortingCriteria[i] == sobCustom) {
			CComPtr<ITreeViewItem> pTvwItem1 = ClassFactory::InitTreeItem(hItem1, this);
			CComPtr<ITreeViewItem> pTvwItem2 = ClassFactory::InitTreeItem(hItem2, this);
			if(FAILED(Raise_CompareItems(pTvwItem1, pTvwItem2, reinterpret_cast<CompareResultConstants*>(&result)))) {
				result = 0;
				earlyAbort = TRUE;
			} else {
				earlyAbort = (result != 0);
			}
			continue;
		}

		// we'll need some details about the items
		if((sortingSettings.sortingCriteria[i] != sobShell) && (item1.mask == 0)) {
			item1.hItem = hItem1;
			item1.cchTextMax = MAX_ITEMTEXTLENGTH;
			item1.pszText = pItem1TextBuffer;
			item1.stateMask = TVIS_SELECTED | TVIS_STATEIMAGEMASK;
			item1.mask = TVIF_HANDLE | TVIF_TEXT | TVIF_STATE;
			SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item1));
			if(!item1.pszText) {
				item1.pszText = TEXT("");
			}
			item2.hItem = hItem2;
			item2.cchTextMax = MAX_ITEMTEXTLENGTH;
			item2.pszText = pItem2TextBuffer;
			item2.stateMask = TVIS_SELECTED | TVIS_STATEIMAGEMASK;
			item2.mask = TVIF_HANDLE | TVIF_TEXT | TVIF_STATE;
			SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item2));
			if(!item2.pszText) {
				item2.pszText = TEXT("");
			}
		}
		switch(sortingSettings.sortingCriteria[i]) {
			case sobShell:
				#ifdef INCLUDESHELLBROWSERINTERFACE
					if(shellBrowserInterface.pInternalMessageListener) {
						HTREEITEM items[2] = {hItem1, hItem2};
						BOOL wasHandled = FALSE;
						result = static_cast<short>(HRESULT_CODE(shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_COMPAREITEMS, reinterpret_cast<WPARAM>(items), reinterpret_cast<LPARAM>(&wasHandled))));
						if(!wasHandled) {
							result = 0;
						}
					}
				#else
					result = 0;
				#endif
				break;
			case sobText:
				if(sortingSettings.localeID != static_cast<LCID>(-1) || sortingSettings.flagsForCompareString != 0) {
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					result = CompareString(sortingSettings.localeID, sortingSettings.flagsForCompareString, item1.pszText, lstrlen(item1.pszText), item2.pszText, lstrlen(item2.pszText)) - 2;
				} else {
					if(sortingSettings.caseSensitive) {
						result = lstrcmp(item1.pszText, item2.pszText);
					} else {
						result = lstrcmpi(item1.pszText, item2.pszText);
					}
				}
				break;
			case sobNumericIntText:
				{
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					CComBSTR text1 = item1.pszText;
					CComBSTR text2 = item2.pszText;
					LONG64 int1;
					LONG64 int2;
					if(SUCCEEDED(VarI8FromStr(text1, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &int1))) {
						if(SUCCEEDED(VarI8FromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &int2))) {
							if(int1 == int2) {
								result = 0;
							} else {
								result = ((int1 < int2) ? -1 : 1);
							}
						} else {
							// text2 is illegal - move it to the end
							result = -1;
						}
					} else {
						// text1 is illegal
						if(SUCCEEDED(VarI8FromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &int2))) {
							// text2 is valid - move it to the front
							result = 1;
						} else {
							// both are illegal - sort by text
							result = CompareString(sortingSettings.localeID, sortingSettings.flagsForCompareString, item1.pszText, lstrlen(item1.pszText), item2.pszText, lstrlen(item2.pszText)) - 2;
						}
					}
				}
				break;
			case sobNumericFloatText:
				{
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					CComBSTR text1 = item1.pszText;
					CComBSTR text2 = item2.pszText;
					DOUBLE dbl1;
					DOUBLE dbl2;
					if(SUCCEEDED(VarR8FromStr(text1, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &dbl1))) {
						if(SUCCEEDED(VarR8FromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &dbl2))) {
							if(dbl1 == dbl2) {
								result = 0;
							} else {
								result = ((dbl1 < dbl2) ? -1 : 1);
							}
						} else {
							// text2 is illegal - move it to the end
							result = -1;
						}
					} else {
						// text1 is illegal
						if(SUCCEEDED(VarR8FromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &dbl2))) {
							// text2 is valid - move it to the front
							result = 1;
						} else {
							// both are illegal - sort by text
							result = CompareString(sortingSettings.localeID, sortingSettings.flagsForCompareString, item1.pszText, lstrlen(item1.pszText), item2.pszText, lstrlen(item2.pszText)) - 2;
						}
					}
				}
				break;
			case sobDateTimeText:
				{
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					CComBSTR text1 = item1.pszText;
					CComBSTR text2 = item2.pszText;
					DATE date1;
					DATE date2;
					if(SUCCEEDED(VarDateFromStr(text1, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &date1))) {
						if(SUCCEEDED(VarDateFromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &date2))) {
							if(date1 == date2) {
								result = 0;
							} else {
								result = ((date1 < date2) ? -1 : 1);
							}
						} else {
							// text2 is illegal - move it to the end
							result = -1;
						}
					} else {
						// text1 is illegal
						if(SUCCEEDED(VarDateFromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &date2))) {
							// text2 is valid - move it to the front
							result = 1;
						} else {
							// both are illegal - sort by text
							result = CompareString(sortingSettings.localeID, sortingSettings.flagsForCompareString, item1.pszText, lstrlen(item1.pszText), item2.pszText, lstrlen(item2.pszText)) - 2;
						}
					}
				}
				break;
			case sobCurrencyText:
				{
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					CComBSTR text1 = item1.pszText;
					CComBSTR text2 = item2.pszText;
					CY curr1;
					CY curr2;
					if(SUCCEEDED(VarCyFromStr(text1, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &curr1))) {
						if(SUCCEEDED(VarCyFromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &curr2))) {
							if(curr1.int64 == curr2.int64) {
								result = 0;
							} else {
								result = ((curr1.int64 < curr2.int64) ? -1 : 1);
							}
						} else {
							// text2 is illegal - move it to the end
							result = -1;
						}
					} else {
						// text1 is illegal
						if(SUCCEEDED(VarCyFromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &curr2))) {
							// text2 is valid - move it to the front
							result = 1;
						} else {
							// both are illegal - sort by text
							result = CompareString(sortingSettings.localeID, sortingSettings.flagsForCompareString, item1.pszText, lstrlen(item1.pszText), item2.pszText, lstrlen(item2.pszText)) - 2;
						}
					}
				}
				break;
			case sobSelectionState:
				if((item1.state & TVIS_SELECTED) == (item2.state & TVIS_SELECTED)) {
					result = 0;
				} else if(item1.state & TVIS_SELECTED) {
					result = -1;
				} else {
					result = 1;
				}
				break;
			case sobStateImageIndex:
				if(((item1.state & TVIS_STATEIMAGEMASK) >> 12) == ((item2.state & TVIS_STATEIMAGEMASK) >> 12)) {
					result = 0;
				} else {
					result = (((item1.state & TVIS_STATEIMAGEMASK) >> 12) > ((item2.state & TVIS_STATEIMAGEMASK) >> 12) ? -1 : 1);
				}
				break;
		}

		if(result != 0) {
			earlyAbort = TRUE;
		}
	}

	// ensure we don't return illegal values
	if(result < 0) {
		result = -1;
	} else if(result > 0) {
		result = 1;
	}

	if(properties.sortOrder == soDescending) {
		result = -result;
	}
	return result;
}
// implementation of IItemComparator
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP ExplorerTreeView::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr;
	switch(property) {
		case DISPID_EXTVW_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_EXTVW_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_EXTVW_EDITIMEMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.editIMEMode), description);
			break;
		case DISPID_EXTVW_HOTTRACKING:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.hotTracking), description);
			break;
		case DISPID_EXTVW_IMEMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.IMEMode), description);
			break;
		case DISPID_EXTVW_LINESTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.lineStyle), description);
			break;
		case DISPID_EXTVW_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_EXTVW_OLEDRAGIMAGESTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.oleDragImageStyle), description);
			break;
		case DISPID_EXTVW_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_EXTVW_SCROLLBARS:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.scrollBars), description);
			break;
		case DISPID_EXTVW_SINGLEEXPAND:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.singleExpand), description);
			break;
		case DISPID_EXTVW_SORTORDER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.sortOrder), description);
			break;
		case DISPID_EXTVW_TREEVIEWSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.treeViewStyle), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<ExplorerTreeView>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP ExplorerTreeView::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_EXTVW_BORDERSTYLE:
		case DISPID_EXTVW_LINESTYLE:
		case DISPID_EXTVW_OLEDRAGIMAGESTYLE:
		case DISPID_EXTVW_SORTORDER:
			c = 2;
			break;
		case DISPID_EXTVW_APPEARANCE:
		case DISPID_EXTVW_HOTTRACKING:
		case DISPID_EXTVW_SCROLLBARS:
		case DISPID_EXTVW_SINGLEEXPAND:
			c = 3;
			break;
		case DISPID_EXTVW_RIGHTTOLEFT:
		case DISPID_EXTVW_TREEVIEWSTYLE:
			c = 4;
			break;
		case DISPID_EXTVW_EDITIMEMODE:
		case DISPID_EXTVW_IMEMODE:
			c = 12;
			break;
		case DISPID_EXTVW_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<ExplorerTreeView>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = reinterpret_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = reinterpret_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_EXTVW_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		} else if((property == DISPID_EXTVW_IMEMODE) || (property == DISPID_EXTVW_EDITIMEMODE)) {
			// the enum is -1-based
			--propertyValue;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = reinterpret_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_EXTVW_APPEARANCE:
		case DISPID_EXTVW_BORDERSTYLE:
		case DISPID_EXTVW_EDITIMEMODE:
		case DISPID_EXTVW_HOTTRACKING:
		case DISPID_EXTVW_IMEMODE:
		case DISPID_EXTVW_LINESTYLE:
		case DISPID_EXTVW_MOUSEPOINTER:
		case DISPID_EXTVW_OLEDRAGIMAGESTYLE:
		case DISPID_EXTVW_RIGHTTOLEFT:
		case DISPID_EXTVW_SCROLLBARS:
		case DISPID_EXTVW_SINGLEEXPAND:
		case DISPID_EXTVW_SORTORDER:
		case DISPID_EXTVW_TREEVIEWSTYLE:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<ExplorerTreeView>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	return IPerPropertyBrowsingImpl<ExplorerTreeView>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT ExplorerTreeView::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_EXTVW_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXTVW_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXTVW_HOTTRACKING:
			switch(cookie) {
				case htrNone:
					description = GetResStringWithNumber(IDP_HOTTRACKINGNONE, htrNone);
					break;
				case htrNormal:
					description = GetResStringWithNumber(IDP_HOTTRACKINGNORMAL, htrNormal);
					break;
				case htrWinXPStyle:
					description = GetResStringWithNumber(IDP_HOTTRACKINGWINXPSTYLE, htrWinXPStyle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXTVW_EDITIMEMODE:
		case DISPID_EXTVW_IMEMODE:
			switch(cookie) {
				case imeInherit:
					description = GetResStringWithNumber(IDP_IMEMODEINHERIT, imeInherit);
					break;
				case imeNoControl:
					description = GetResStringWithNumber(IDP_IMEMODENOCONTROL, imeNoControl);
					break;
				case imeOn:
					description = GetResStringWithNumber(IDP_IMEMODEON, imeOn);
					break;
				case imeOff:
					description = GetResStringWithNumber(IDP_IMEMODEOFF, imeOff);
					break;
				case imeDisable:
					description = GetResStringWithNumber(IDP_IMEMODEDISABLE, imeDisable);
					break;
				case imeHiragana:
					description = GetResStringWithNumber(IDP_IMEMODEHIRAGANA, imeHiragana);
					break;
				case imeKatakana:
					description = GetResStringWithNumber(IDP_IMEMODEKATAKANA, imeKatakana);
					break;
				case imeKatakanaHalf:
					description = GetResStringWithNumber(IDP_IMEMODEKATAKANAHALF, imeKatakanaHalf);
					break;
				case imeAlphaFull:
					description = GetResStringWithNumber(IDP_IMEMODEALPHAFULL, imeAlphaFull);
					break;
				case imeAlpha:
					description = GetResStringWithNumber(IDP_IMEMODEALPHA, imeAlpha);
					break;
				case imeHangulFull:
					description = GetResStringWithNumber(IDP_IMEMODEHANGULFULL, imeHangulFull);
					break;
				case imeHangul:
					description = GetResStringWithNumber(IDP_IMEMODEHANGUL, imeHangul);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXTVW_LINESTYLE:
			switch(cookie) {
				case lsLinesAtItem:
					description = GetResStringWithNumber(IDP_LINESTYLEATITEM, lsLinesAtItem);
					break;
				case lsLinesAtRoot:
					description = GetResStringWithNumber(IDP_LINESTYLEATROOT, lsLinesAtRoot);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXTVW_MOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXTVW_OLEDRAGIMAGESTYLE:
			switch(cookie) {
				case odistClassic:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLECLASSIC, odistClassic);
					break;
				case odistAeroIfAvailable:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLEAERO, odistAeroIfAvailable);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXTVW_RIGHTTOLEFT:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTNONE, 0);
					break;
				case rtlText:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXT, rtlText);
					break;
				case rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTLAYOUT, rtlLayout);
					break;
				case rtlText | rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXTLAYOUT, rtlText | rtlLayout);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXTVW_SCROLLBARS:
			switch(cookie) {
				case sbNone:
					description = GetResStringWithNumber(IDP_SCROLLBARSNONE, sbNone);
					break;
				case sbVerticalOnly:
					description = GetResStringWithNumber(IDP_SCROLLBARSVERTICALONLY, sbVerticalOnly);
					break;
				case sbNormal:
					description = GetResStringWithNumber(IDP_SCROLLBARSNORMAL, sbNormal);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXTVW_SINGLEEXPAND:
			switch(cookie) {
				case seNone:
					description = GetResStringWithNumber(IDP_SINGLEEXPANDNONE, seNone);
					break;
				case seNormal:
					description = GetResStringWithNumber(IDP_SINGLEEXPANDNORMAL, seNormal);
					break;
				case seWinXPStyle:
					description = GetResStringWithNumber(IDP_SINGLEEXPANDWINXPSTYLE, seWinXPStyle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXTVW_SORTORDER:
			switch(cookie) {
				case soAscending:
					description = GetResStringWithNumber(IDP_SORTORDERASCENDING, soAscending);
					break;
				case soDescending:
					description = GetResStringWithNumber(IDP_SORTORDERDESCENDING, soDescending);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXTVW_TREEVIEWSTYLE:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_TREEVIEWSTYLETEXTONLY, 0);
					break;
				case tvsExpandos:
					description = GetResStringWithNumber(IDP_TREEVIEWSTYLEEXPANDOS, tvsExpandos);
					break;
				case tvsLines:
					description = GetResStringWithNumber(IDP_TREEVIEWSTYLELINES, tvsLines);
					break;
				case tvsExpandos | tvsLines:
					description = GetResStringWithNumber(IDP_TREEVIEWSTYLEEXPANDOSLINES, tvsExpandos | tvsLines);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP ExplorerTreeView::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 4;
	pPropertyPages->pElems = reinterpret_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StockColorPage;
		pPropertyPages->pElems[2] = CLSID_StockFontPage;
		pPropertyPages->pElems[3] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP ExplorerTreeView::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<ExplorerTreeView>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP ExplorerTreeView::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP ExplorerTreeView::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = NULL;
	pControlInfo->cAccel = 0;
	pControlInfo->dwFlags = 0;
	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT ExplorerTreeView::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT ExplorerTreeView::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_EXTVW_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT ExplorerTreeView::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP ExplorerTreeView::get_AllowDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowDragDrop = !(GetStyle() & TVS_DISABLEDRAGDROP);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowDragDrop);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_AllowDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_ALLOWDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowDragDrop != b) {
		properties.allowDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.allowDragDrop) {
				ModifyStyle(TVS_DISABLEDRAGDROP, 0);
			} else {
				ModifyStyle(0, TVS_DISABLEDRAGDROP);
			}
		}
		FireOnChanged(DISPID_EXTVW_ALLOWDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_AllowLabelEditing(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowLabelEditing = ((GetStyle() & TVS_EDITLABELS) == TVS_EDITLABELS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowLabelEditing);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_AllowLabelEditing(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_ALLOWLABELEDITING);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowLabelEditing != b) {
		properties.allowLabelEditing = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.allowLabelEditing) {
				ModifyStyle(0, TVS_EDITLABELS);
			} else {
				ModifyStyle(TVS_EDITLABELS, 0);
			}
		}
		FireOnChanged(DISPID_EXTVW_ALLOWLABELEDITING);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_AlwaysShowSelection(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.alwaysShowSelection = ((GetStyle() & TVS_SHOWSELALWAYS) == TVS_SHOWSELALWAYS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.alwaysShowSelection);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_AlwaysShowSelection(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_ALWAYSSHOWSELECTION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.alwaysShowSelection != b) {
		properties.alwaysShowSelection = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.alwaysShowSelection) {
				ModifyStyle(0, TVS_SHOWSELALWAYS, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			} else {
				ModifyStyle(TVS_SHOWSELALWAYS, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_ALWAYSSHOWSELECTION);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_APPEARANCE);
	if(properties.appearance != newValue) {
		BOOL requiresRecreation = FALSE;
		if(RunTimeHelper::IsCommCtrl6()) {
			if((newValue == a3D) || (newValue == a3DLight)) {
				requiresRecreation = TRUE;
			}
		}
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(requiresRecreation) {
				RecreateControlWindow();
			} else {
				switch(properties.appearance) {
					case a2D:
						ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						break;
					case a3D:
						ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						break;
					case a3DLight:
						ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						break;
				}
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_EXTVW_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 2;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ExplorerTreeView");
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ExTVw");
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_AutoHScroll(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.autoHScroll = ((SendMessage(TVM_GETEXTENDEDSTYLE, 0, 0) & TVS_EX_AUTOHSCROLL) == TVS_EX_AUTOHSCROLL);
	}

	*pValue = BOOL2VARIANTBOOL(properties.autoHScroll);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_AutoHScroll(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_AUTOHSCROLL);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.autoHScroll != b) {
		properties.autoHScroll = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.autoHScroll) {
				SendMessage(TVM_SETEXTENDEDSTYLE, TVS_EX_AUTOHSCROLL, TVS_EX_AUTOHSCROLL);
			} else {
				SendMessage(TVM_SETEXTENDEDSTYLE, TVS_EX_AUTOHSCROLL, 0);
			}
		}
		FireOnChanged(DISPID_EXTVW_AUTOHSCROLL);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_AutoHScrollPixelsPerSecond(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.autoHScrollPixelsPerSecond;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_AutoHScrollPixelsPerSecond(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_AUTOHSCROLLPIXELSPERSECOND);
	if(properties.autoHScrollPixelsPerSecond != newValue) {
		properties.autoHScrollPixelsPerSecond = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			SendMessage(TVM_SETAUTOSCROLLINFO, properties.autoHScrollPixelsPerSecond, properties.autoHScrollRedrawInterval);
		}
		FireOnChanged(DISPID_EXTVW_AUTOHSCROLLPIXELSPERSECOND);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_AutoHScrollRedrawInterval(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.autoHScrollRedrawInterval;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_AutoHScrollRedrawInterval(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_AUTOHSCROLLREDRAWINTERVAL);
	if(properties.autoHScrollRedrawInterval != newValue) {
		properties.autoHScrollRedrawInterval = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			SendMessage(TVM_SETAUTOSCROLLINFO, properties.autoHScrollPixelsPerSecond, properties.autoHScrollRedrawInterval);
		}
		FireOnChanged(DISPID_EXTVW_AUTOHSCROLLREDRAWINTERVAL);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(TVM_GETBKCOLOR, 0, 0));
		if(color == CLR_NONE) {
			properties.backColor = 0x80000000 | COLOR_WINDOW;
		} else if(color != OLECOLOR2COLORREF(properties.backColor)) {
			properties.backColor = color;
		}
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TVM_SETBKCOLOR, 0, OLECOLOR2COLORREF(properties.backColor));
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_BlendSelectedItemsIcons(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.blendSelectedItemsIcons);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_BlendSelectedItemsIcons(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_BLENDSELECTEDITEMSICONS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.blendSelectedItemsIcons != b) {
		properties.blendSelectedItemsIcons = b;
		SetDirty(TRUE);
		FireViewChange();
		FireOnChanged(DISPID_EXTVW_BLENDSELECTEDITEMSICONS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_BORDERSTYLE);
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_BuiltInStateImages(BuiltInStateImagesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BuiltInStateImagesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(IsComctl32Version610OrNewer()) {
			properties.showStateImages = ((GetStyle() & TVS_CHECKBOXES) == TVS_CHECKBOXES);
			DWORD extendedStyle = static_cast<DWORD>(SendMessage(TVM_GETEXTENDEDSTYLE, 0, 0));
			if(!properties.showStateImages) {
				properties.showStateImages = ((extendedStyle & (TVS_EX_PARTIALCHECKBOXES | TVS_EX_DIMMEDCHECKBOXES | TVS_EX_EXCLUSIONCHECKBOXES)) != 0);
			}
			if(properties.showStateImages) {
				properties.builtInStateImages = bisiSelectedAndUnselected;
				if(extendedStyle & TVS_EX_PARTIALCHECKBOXES) {
					properties.builtInStateImages = static_cast<BuiltInStateImagesConstants>(static_cast<long>(properties.builtInStateImages) | bisiIndeterminate);
				}
				if(extendedStyle & TVS_EX_DIMMEDCHECKBOXES) {
					properties.builtInStateImages = static_cast<BuiltInStateImagesConstants>(static_cast<long>(properties.builtInStateImages) | bisiSelectedDimmed);
				}
				if(extendedStyle & TVS_EX_EXCLUSIONCHECKBOXES) {
					properties.builtInStateImages = static_cast<BuiltInStateImagesConstants>(static_cast<long>(properties.builtInStateImages) | bisiExcluded);
				}
			}
		} else {
			properties.builtInStateImages = bisiSelectedAndUnselected;
		}
	}

	*pValue = properties.builtInStateImages;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_BuiltInStateImages(BuiltInStateImagesConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_BUILTINSTATEIMAGES);
	if(properties.builtInStateImages != newValue) {
		properties.builtInStateImages = newValue;
		SetDirty(TRUE);

		if(properties.showStateImages && IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.builtInStateImages == bisiSelectedAndUnselected) {
				DWORD extendedStyle = static_cast<DWORD>(SendMessage(TVM_GETEXTENDEDSTYLE, 0, 0));
				if(extendedStyle & (TVS_EX_PARTIALCHECKBOXES | TVS_EX_DIMMEDCHECKBOXES | TVS_EX_EXCLUSIONCHECKBOXES)) {
					// we can't remove state images dynamically
					RecreateControlWindow();
				} else {
					// no TVS_EX_*CHECKBOXES style set yet, so no need for recreation
					ModifyStyle(0, TVS_CHECKBOXES);
				}
			} else {
				DWORD extendedStyle = static_cast<DWORD>(SendMessage(TVM_GETEXTENDEDSTYLE, 0, 0));
				if(extendedStyle & (TVS_EX_PARTIALCHECKBOXES | TVS_EX_DIMMEDCHECKBOXES | TVS_EX_EXCLUSIONCHECKBOXES)) {
					// we can't call TVM_SETEXTENDEDSTYLE twice, so recreate the control
					RecreateControlWindow();
				} else if(GetStyle() & TVS_CHECKBOXES) {
					// we can't remove TVS_CHECKBOXES dynamically
					RecreateControlWindow();
				} else {
					// TVM_SETEXTENDEDSTYLE was not yet called, call it now
					extendedStyle = 0;
					if(properties.builtInStateImages & bisiIndeterminate) {
						extendedStyle |= TVS_EX_PARTIALCHECKBOXES;
					}
					if(properties.builtInStateImages & bisiSelectedDimmed) {
						extendedStyle |= TVS_EX_DIMMEDCHECKBOXES;
					}
					if(properties.builtInStateImages & bisiExcluded) {
						extendedStyle |= TVS_EX_EXCLUSIONCHECKBOXES;
					}
					SendMessage(TVM_SETEXTENDEDSTYLE, (TVS_EX_PARTIALCHECKBOXES | TVS_EX_EXCLUSIONCHECKBOXES | TVS_EX_DIMMEDCHECKBOXES), extendedStyle);
				}
			}
		}
		FireOnChanged(DISPID_EXTVW_BUILTINSTATEIMAGES);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_CaretChangedDelayTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.caretChangedDelayTime;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_CaretChangedDelayTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_CARETCHANGEDDELAYTIME);
	if((newValue < 0) || (newValue > 5000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.caretChangedDelayTime != newValue) {
		properties.caretChangedDelayTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXTVW_CARETCHANGEDDELAYTIME);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_CaretItem(VARIANT_BOOL/* noSingleExpand = VARIANT_TRUE*/, ITreeViewItem** ppCaretItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppCaretItem, ITreeViewItem*);
	if(!ppCaretItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		HTREEITEM hCaretItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0));
		ClassFactory::InitTreeItem(hCaretItem, this, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppCaretItem));
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP ExplorerTreeView::putref_CaretItem(VARIANT_BOOL noSingleExpand/* = VARIANT_TRUE*/, ITreeViewItem* pNewCaretItem/* = NULL*/)
{
	PUTPROPPROLOG(DISPID_EXTVW_CARETITEM);
	HRESULT hr = E_FAIL;

	HTREEITEM hNewCaretItem = NULL;
	if(pNewCaretItem) {
		OLE_HANDLE h = NULL;
		pNewCaretItem->get_Handle(&h);
		hNewCaretItem = static_cast<HTREEITEM>(LongToHandle(h));
		// TODO: Shouldn't we AddRef' pNewCaretItem?
	}

	if(IsWindow()) {
		HTREEITEM hPreviousCaretItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0));
		if(hNewCaretItem != hPreviousCaretItem) {
			// prepare the control for selecting the new caret
			AddInternalCaretChange(hNewCaretItem, TVC_INTERNAL);
			// select the new caret
			if(noSingleExpand && (GetStyle() & TVS_SINGLEEXPAND)) {
				EnterNoSingleExpandSection();
				if(SUCCEEDED(SendMessage(TVM_SELECTITEM, static_cast<WPARAM>(TVGN_CARET), reinterpret_cast<LPARAM>(hNewCaretItem)))) {
					hr = S_OK;
				}
				LeaveNoSingleExpandSection();
			} else {
				if(SUCCEEDED(SendMessage(TVM_SELECTITEM, static_cast<WPARAM>(TVGN_CARET), reinterpret_cast<LPARAM>(hNewCaretItem)))) {
					hr = S_OK;
				}
			}
			// clean up
			RemoveInternalCaretChange(hNewCaretItem);
		} else {
			hr = S_OK;
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXTVW_CARETITEM);
	return hr;
}

STDMETHODIMP ExplorerTreeView::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		if(GetStyle() & TVS_INFOTIP) {
			properties.disabledEvents = static_cast<DisabledEventsConstants>(properties.disabledEvents & ~deItemGetInfoTipText);
		} else {
			properties.disabledEvents = static_cast<DisabledEventsConstants>(properties.disabledEvents | deItemGetInfoTipText);
		}
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if((properties.disabledEvents & deItemGetInfoTipText) != (newValue & deItemGetInfoTipText)) {
			if(IsWindow()) {
				if(newValue & deItemGetInfoTipText) {
					ModifyStyle(TVS_INFOTIP, 0);
				} else {
					ModifyStyle(0, TVS_INFOTIP);
				}
			}
		}

		if((properties.disabledEvents & deTreeMouseEvents) != (newValue & deTreeMouseEvents)) {
			if(IsWindow()) {
				if(newValue & deTreeMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = *this;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
				}
				hItemUnderMouse = NULL;
			}
		}

		if((properties.disabledEvents & deEditMouseEvents) != (newValue & deEditMouseEvents)) {
			if(containedEdit.IsWindow()) {
				if(newValue & deEditMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = containedEdit.m_hWnd;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
				}
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXTVW_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_EXTVW_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_DragExpandTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dragExpandTime;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_DragExpandTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_DRAGEXPANDTIME);
	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.dragExpandTime != newValue) {
		properties.dragExpandTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXTVW_DRAGEXPANDTIME);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_DraggedItems(ITreeViewItemContainer** ppItems)
{
	ATLASSERT_POINTER(ppItems, ITreeViewItemContainer*);
	if(!ppItems) {
		return E_POINTER;
	}

	*ppItems = NULL;
	if(dragDropStatus.pDraggedItems) {
		return dragDropStatus.pDraggedItems->Clone(ppItems);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_DragScrollTimeBase(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dragScrollTimeBase;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_DragScrollTimeBase(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_DRAGSCROLLTIMEBASE);
	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.dragScrollTimeBase != newValue) {
		properties.dragScrollTimeBase = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXTVW_DRAGSCROLLTIMEBASE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_DrawImagesAsynchronously(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.drawImagesAsynchronously = ((SendMessage(TVM_GETEXTENDEDSTYLE, 0, 0) & TVS_EX_DRAWIMAGEASYNC) == TVS_EX_DRAWIMAGEASYNC);
	}

	*pValue = BOOL2VARIANTBOOL(properties.drawImagesAsynchronously);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_DrawImagesAsynchronously(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_DRAWIMAGESASYNCHRONOUSLY);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.drawImagesAsynchronously != b) {
		properties.drawImagesAsynchronously = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.drawImagesAsynchronously) {
				SendMessage(TVM_SETEXTENDEDSTYLE, TVS_EX_DRAWIMAGEASYNC, TVS_EX_DRAWIMAGEASYNC);
			} else {
				SendMessage(TVM_SETEXTENDEDSTYLE, TVS_EX_DRAWIMAGEASYNC, 0);
			}
		}
		FireOnChanged(DISPID_EXTVW_DRAWIMAGESASYNCHRONOUSLY);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_DropHilitedItem(ITreeViewItem** ppDropHilitedItem)
{
	ATLASSERT_POINTER(ppDropHilitedItem, ITreeViewItem*);
	if(!ppDropHilitedItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		HTREEITEM hDropHilitedItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_DROPHILITE), 0));
		ClassFactory::InitTreeItem(hDropHilitedItem, this, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppDropHilitedItem));
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::putref_DropHilitedItem(ITreeViewItem* pNewDropHilitedItem)
{
	PUTPROPPROLOG(DISPID_EXTVW_DROPHILITEDITEM);
	HRESULT hr = E_FAIL;

	HTREEITEM hNewDropHilitedItem = NULL;
	if(pNewDropHilitedItem) {
		OLE_HANDLE h = NULL;
		pNewDropHilitedItem->get_Handle(&h);
		hNewDropHilitedItem = static_cast<HTREEITEM>(LongToHandle(h));
		// TODO: Shouldn't we AddRef' pNewDropHilitedItem?
	}

	if(IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);
		if(SUCCEEDED(SendMessage(TVM_SELECTITEM, static_cast<WPARAM>(TVGN_DROPHILITE), reinterpret_cast<LPARAM>(hNewDropHilitedItem)))) {
			hr = S_OK;
		}
		dragDropStatus.ShowDragImage(TRUE);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXTVW_DROPHILITEDITEM);
	return hr;
}

STDMETHODIMP ExplorerTreeView::get_EditBackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.editBackColor;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_EditBackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_EDITBACKCOLOR);
	if(properties.editBackColor != newValue) {
		properties.editBackColor = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			DeleteObject(hEditBackColorBrush);
			hEditBackColorBrush = CreateSolidBrush(OLECOLOR2COLORREF(properties.editBackColor));
			containedEdit.Invalidate();
		}
		FireOnChanged(DISPID_EXTVW_EDITBACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_EditForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.editForeColor;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_EditForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_EDITFORECOLOR);
	if(properties.editForeColor != newValue) {
		properties.editForeColor = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			containedEdit.Invalidate();
		}
		FireOnChanged(DISPID_EXTVW_EDITFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_EditHoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.editHoverTime;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_EditHoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_EDITHOVERTIME);
	if((newValue < 0) && (newValue != -1))
	{
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.editHoverTime != newValue) {
		properties.editHoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXTVW_EDITHOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_EditIMEMode(IMEModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, IMEModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsInDesignMode()) {
		*pValue = properties.editIMEMode;
	} else {
		if((GetFocus() == containedEdit.m_hWnd) && (GetEffectiveIMEMode_Edit() != imeNoControl)) {
			// we have control over the IME, so retrieve its current config
			IMEModeConstants ime = GetCurrentIMEContextMode(containedEdit.m_hWnd);
			if((ime != imeInherit) && (properties.editIMEMode != imeInherit)) {
				properties.editIMEMode = ime;
			}
		}
		*pValue = GetEffectiveIMEMode_Edit();
	}

	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_EditIMEMode(IMEModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_EDITIMEMODE);
	if(properties.editIMEMode != newValue) {
		properties.editIMEMode = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode()) {
			if(GetFocus() == containedEdit.m_hWnd) {
				// we have control over the IME, so update its config
				IMEModeConstants ime = GetEffectiveIMEMode_Edit();
				if(ime != imeNoControl) {
					SetCurrentIMEContextMode(containedEdit.m_hWnd, ime);
				}
			}
		}
		FireOnChanged(DISPID_EXTVW_EDITIMEMODE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_EditText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = NULL;
	if(containedEdit.IsWindow()) {
		containedEdit.GetWindowText(pValue);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerTreeView::put_EditText(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_EDITTEXT);
	if(!containedEdit.IsWindow()) {
		return E_FAIL;
	}

	containedEdit.SetWindowText(COLE2CT(newValue));

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXTVW_EDITTEXT);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_ENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
			FireViewChange();
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_EXTVW_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_FadeExpandos(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.fadeExpandos = ((SendMessage(TVM_GETEXTENDEDSTYLE, 0, 0) & TVS_EX_FADEINOUTEXPANDOS) == TVS_EX_FADEINOUTEXPANDOS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.fadeExpandos);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_FadeExpandos(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_FADEEXPANDOS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.fadeExpandos != b) {
		properties.fadeExpandos = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.fadeExpandos) {
				SendMessage(TVM_SETEXTENDEDSTYLE, TVS_EX_FADEINOUTEXPANDOS, TVS_EX_FADEINOUTEXPANDOS);
			} else {
				SendMessage(TVM_SETEXTENDEDSTYLE, TVS_EX_FADEINOUTEXPANDOS, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_FADEEXPANDOS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_FavoritesStyle(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.favoritesStyle);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_FavoritesStyle(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_FAVORITESSTYLE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.favoritesStyle != b) {
		properties.favoritesStyle = b;
		SetDirty(TRUE);
		FireViewChange();
		FireOnChanged(DISPID_EXTVW_FAVORITESSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_FirstRootItem(ITreeViewItem** ppFirstItem)
{
	ATLASSERT_POINTER(ppFirstItem, ITreeViewItem*);
	if(!ppFirstItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		HTREEITEM hFirstItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_ROOT), 0));
		ClassFactory::InitTreeItem(hFirstItem, this, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppFirstItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerTreeView::get_FirstVisibleItem(ITreeViewItem** ppFirstItem)
{
	ATLASSERT_POINTER(ppFirstItem, ITreeViewItem*);
	if(!ppFirstItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		HTREEITEM hFirstItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_FIRSTVISIBLE), 0));
		ClassFactory::InitTreeItem(hFirstItem, this, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppFirstItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerTreeView::putref_FirstVisibleItem(ITreeViewItem* pNewFirstItem)
{
	PUTPROPPROLOG(DISPID_EXTVW_FIRSTVISIBLEITEM);
	HRESULT hr = E_FAIL;

	HTREEITEM hFirstItem = NULL;
	if(pNewFirstItem) {
		OLE_HANDLE h = NULL;
		pNewFirstItem->get_Handle(&h);
		hFirstItem = static_cast<HTREEITEM>(LongToHandle(h));
		// TODO: Shouldn't we AddRef' pNewFirstItem?
	}

	if(IsWindow()) {
		if(SUCCEEDED(SendMessage(TVM_SELECTITEM, static_cast<WPARAM>(TVGN_FIRSTVISIBLE), reinterpret_cast<LPARAM>(hFirstItem)))) {
			hr = S_OK;
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXTVW_FIRSTVISIBLEITEM);
	FireViewChange();
	return hr;
}

STDMETHODIMP ExplorerTreeView::get_Font(IFontDisp** ppFont)
{
	ATLASSERT_POINTER(ppFont, IFontDisp*);
	if(!ppFont) {
		return E_POINTER;
	}

	if(*ppFont) {
		(*ppFont)->Release();
		*ppFont = NULL;
	}
	if(properties.font.pFontDisp) {
		properties.font.pFontDisp->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(ppFont));
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_EXTVW_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			CComQIPtr<IFont, &IID_IFont> pFont(pNewFont);
			if(pFont) {
				CComPtr<IFont> pClonedFont = NULL;
				pFont->Clone(&pClonedFont);
				if(pClonedFont) {
					pClonedFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
				}
			}
		}
		properties.font.StartWatching();
	}
	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXTVW_FONT);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_EXTVW_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			pNewFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
		}
		properties.font.StartWatching();
	} else if(pNewFont) {
		pNewFont->AddRef();
	}

	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXTVW_FONT);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(TVM_GETTEXTCOLOR, 0, 0));
		if(color == CLR_NONE) {
			properties.foreColor = 0x80000000 | COLOR_WINDOWTEXT;
		} else if(color != OLECOLOR2COLORREF(properties.foreColor)) {
			properties.foreColor = color;
		}
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_ForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_FORECOLOR);
	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TVM_SETTEXTCOLOR, 0, OLECOLOR2COLORREF(properties.foreColor));
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_FORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_FullRowSelect(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.fullRowSelect = ((GetStyle() & TVS_FULLROWSELECT) == TVS_FULLROWSELECT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.fullRowSelect);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_FullRowSelect(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_FULLROWSELECT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.fullRowSelect != b) {
		properties.fullRowSelect = b;
		SetDirty(TRUE);

		if(properties.fullRowSelect) {
			// lines and full row select are incompatible
			TreeViewStyleConstants style = static_cast<TreeViewStyleConstants>(0);
			get_TreeViewStyle(&style);
			style = static_cast<TreeViewStyleConstants>(static_cast<long>(style) & ~tvsLines);
			put_TreeViewStyle(style);
		}
		if(IsWindow()) {
			if(properties.fullRowSelect) {
				ModifyStyle(0, TVS_FULLROWSELECT, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			} else {
				ModifyStyle(TVS_FULLROWSELECT, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_FULLROWSELECT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_GroupBoxColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.groupBoxColor;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_GroupBoxColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_GROUPBOXCOLOR);
	if(properties.groupBoxColor != newValue) {
		properties.groupBoxColor = newValue;
		SetDirty(TRUE);
		FireViewChange();
		FireOnChanged(DISPID_EXTVW_GROUPBOXCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_hDragImageList(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(dragDropStatus.IsDragging()) {
		*pValue = HandleToLong(dragDropStatus.hDragImageList);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerTreeView::get_hImageList(ImageListConstants imageList, OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = NULL;
	switch(imageList) {
		case ilItems:
			if(IsWindow()) {
				*pValue = HandleToLong(cachedSettings.hImageList);
			}
			break;
		case ilHighResolution:
			*pValue = HandleToLong(cachedSettings.hHighResImageList);
			break;
		case ilState:
			if(IsWindow()) {
				*pValue = HandleToLong(cachedSettings.hStateImageList);
			}
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_hImageList(ImageListConstants imageList, OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_HIMAGELIST);
	BOOL fireViewChange = TRUE;
	switch(imageList) {
		case ilItems:
			if(IsWindow()) {
				SendMessage(TVM_SETIMAGELIST, TVSIL_NORMAL, newValue);
			}
			break;
		case ilHighResolution:
			cachedSettings.hHighResImageList = reinterpret_cast<HIMAGELIST>(LongToHandle(newValue));
			fireViewChange = FALSE;
			break;
		case ilState:
			if(IsWindow()) {
				SendMessage(TVM_SETIMAGELIST, TVSIL_STATE, newValue);
			}
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}

	FireOnChanged(DISPID_EXTVW_HIMAGELIST);
	if(fireViewChange) {
		FireViewChange();
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_HotTracking(HotTrackingConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HotTrackingConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetStyle() & TVS_TRACKSELECT) {
			if((properties.hotTracking != htrNormal) && (properties.hotTracking != htrWinXPStyle)) {
				properties.hotTracking = htrNormal;
			}
		} else {
			properties.hotTracking = htrNone;
		}
	}

	*pValue = properties.hotTracking;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_HotTracking(HotTrackingConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_HOTTRACKING);
	if(properties.hotTracking != newValue) {
		properties.hotTracking = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.hotTracking) {
				case htrNone:
					ModifyStyle(TVS_TRACKSELECT, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case htrNormal:
				case htrWinXPStyle:
					ModifyStyle(0, TVS_TRACKSELECT, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_HOTTRACKING);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXTVW_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_hWndEdit(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		//*pValue = HandleToLong(SendMessage(TVM_GETEDITCONTROL, 0, 0));
		*pValue = HandleToLong(containedEdit.m_hWnd);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_hWndToolTip(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(reinterpret_cast<HWND>(SendMessage(TVM_GETTOOLTIPS, 0, 0)));
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_hWndToolTip(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_HWNDTOOLTIP);
	if(IsWindow()) {
		SendMessage(TVM_SETTOOLTIPS, reinterpret_cast<WPARAM>(LongToHandle(newValue)), 0);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXTVW_HWNDTOOLTIP);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_IMEMode(IMEModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, IMEModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsInDesignMode()) {
		*pValue = properties.IMEMode;
	} else {
		if((GetFocus() == *this) && (GetEffectiveIMEMode() != imeNoControl)) {
			// we have control over the IME, so retrieve its current config
			IMEModeConstants ime = GetCurrentIMEContextMode(*this);
			if((ime != imeInherit) && (properties.IMEMode != imeInherit)) {
				properties.IMEMode = ime;
			}
		}
		*pValue = GetEffectiveIMEMode();
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_IMEMode(IMEModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_IMEMODE);
	if(properties.IMEMode != newValue) {
		properties.IMEMode = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode()) {
			if(GetFocus() == *this) {
				// we have control over the IME, so update its config
				IMEModeConstants ime = GetEffectiveIMEMode();
				if(ime != imeNoControl) {
					SetCurrentIMEContextMode(*this, ime);
				}
			} else if((GetFocus() == containedEdit.m_hWnd) && (properties.editIMEMode == imeInherit)) {
				// the label-edit control inherits from us and has the focus, so update the IME's config
				SetCurrentIMEContextMode(containedEdit.m_hWnd, GetEffectiveIMEMode_Edit());
			}
		}
		FireOnChanged(DISPID_EXTVW_IMEMODE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_IncrementalSearchString(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsWindow()) {
		return E_FAIL;
	}

	*pValue = SysAllocString(L"");
	UINT chars = static_cast<UINT>(SendMessage(TVM_GETISEARCHSTRING, 0, 0));
	if(chars > 0) {
		LPTSTR pBuffer = new TCHAR[chars + 1];
		SendMessage(TVM_GETISEARCHSTRING, 0, reinterpret_cast<LPARAM>(pBuffer));
		*pValue = _bstr_t(pBuffer).Detach();
		delete[] pBuffer;
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_Indent(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.indent = static_cast<OLE_XSIZE_PIXELS>(SendMessage(TVM_GETINDENT, 0, 0));
	}

	*pValue = properties.indent;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_Indent(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_INDENT);
	if(properties.indent != newValue) {
		properties.indent = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TVM_SETINDENT, properties.indent, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_INDENT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_IndentStateImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.indentStateImages = ((SendMessage(TVM_GETEXTENDEDSTYLE, 0, 0) & TVS_EX_NOINDENTSTATE) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.indentStateImages);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_IndentStateImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_INDENTSTATEIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.indentStateImages != b) {
		properties.indentStateImages = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.indentStateImages) {
				SendMessage(TVM_SETEXTENDEDSTYLE, TVS_EX_NOINDENTSTATE, 0);
			} else {
				SendMessage(TVM_SETEXTENDEDSTYLE, TVS_EX_NOINDENTSTATE, TVS_EX_NOINDENTSTATE);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_INDENTSTATEIMAGES);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_InsertMarkColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(TVM_GETINSERTMARKCOLOR, 0, 0));
		if(color == CLR_NONE) {
			properties.insertMarkColor = 0;
		} else if(color != OLECOLOR2COLORREF(properties.insertMarkColor)) {
			properties.insertMarkColor = color;
		}
	}

	*pValue = properties.insertMarkColor;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_InsertMarkColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_INSERTMARKCOLOR);
	if(properties.insertMarkColor != newValue) {
		properties.insertMarkColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TVM_SETINSERTMARKCOLOR, 0, OLECOLOR2COLORREF(properties.insertMarkColor));
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_INSERTMARKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_ItemBoundingBoxDefinition(ItemBoundingBoxDefinitionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemBoundingBoxDefinitionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemBoundingBoxDefinition;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_ItemBoundingBoxDefinition(ItemBoundingBoxDefinitionConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_ITEMBOUNDINGBOXDEFINITION);
	if(properties.itemBoundingBoxDefinition != newValue) {
		properties.itemBoundingBoxDefinition = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXTVW_ITEMBOUNDINGBOXDEFINITION);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_ItemHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.itemHeight = static_cast<OLE_YSIZE_PIXELS>(cachedSettings.itemHeight);
	}

	*pValue = properties.itemHeight;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_ItemHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_ITEMHEIGHT);
	if(properties.itemHeight != newValue) {
		properties.itemHeight = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TVM_SETITEMHEIGHT, properties.itemHeight, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_ITEMHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_ItemXBorder(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.itemBorders = static_cast<DWORD>(SendMessage(TVM_GETBORDER, 0, 0));
	}

	*pValue = LOWORD(properties.itemBorders);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_ItemXBorder(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_ITEMXBORDER);
	if(LOWORD(properties.itemBorders) != newValue) {
		properties.itemBorders = MAKELONG(newValue, HIWORD(properties.itemBorders));
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TVM_SETBORDER, TVSBF_XBORDER, properties.itemBorders);
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_ITEMXBORDER);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_ItemYBorder(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.itemBorders = static_cast<DWORD>(SendMessage(TVM_GETBORDER, 0, 0));
	}

	*pValue = HIWORD(properties.itemBorders);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_ItemYBorder(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_ITEMYBORDER);
	if(HIWORD(properties.itemBorders) != newValue) {
		properties.itemBorders = MAKELONG(LOWORD(properties.itemBorders), newValue);
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TVM_SETBORDER, TVSBF_YBORDER, properties.itemBorders);
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_ITEMYBORDER);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_LastRootItem(ITreeViewItem** ppLastItem)
{
	ATLASSERT_POINTER(ppLastItem, ITreeViewItem*);
	if(!ppLastItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		HTREEITEM hLastItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_ROOT), 0));
		HTREEITEM hDummy = hLastItem;
		while(hDummy) {
			hLastItem = hDummy;
			hDummy = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hLastItem)));
		}
		ClassFactory::InitTreeItem(hLastItem, this, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppLastItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerTreeView::get_LastVisibleItem(ITreeViewItem** ppLastItem)
{
	ATLASSERT_POINTER(ppLastItem, ITreeViewItem*);
	if(!ppLastItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		HTREEITEM hLastItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_LASTVISIBLE), 0));
		ClassFactory::InitTreeItem(hLastItem, this, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppLastItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerTreeView::get_LineColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(TVM_GETLINECOLOR, 0, 0));
		if(color == CLR_NONE) {
			properties.lineColor = 0x80000000 | COLOR_BTNSHADOW;
		} else if(color != OLECOLOR2COLORREF(properties.lineColor)) {
			properties.lineColor = color;
		}
	}

	*pValue = properties.lineColor;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_LineColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_LINECOLOR);
	if(properties.lineColor != newValue) {
		properties.lineColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TVM_SETLINECOLOR, 0, OLECOLOR2COLORREF(properties.lineColor));
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_LINECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_LineStyle(LineStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, LineStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.lineStyle = ((GetStyle() & TVS_LINESATROOT) == TVS_LINESATROOT ? lsLinesAtRoot : lsLinesAtItem);
	}

	*pValue = properties.lineStyle;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_LineStyle(LineStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_LINESTYLE);
	if(properties.lineStyle != newValue) {
		properties.lineStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.lineStyle) {
				case lsLinesAtItem:
					ModifyStyle(TVS_LINESATROOT, 0);
					break;
				case lsLinesAtRoot:
					ModifyStyle(0, TVS_LINESATROOT);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXTVW_LINESTYLE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_Locale(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.locale;
	if(*pValue == -1) {
		*pValue = GetThreadLocale();
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_Locale(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_LOCALE);
	properties.locale = newValue;

	FireOnChanged(DISPID_EXTVW_LOCALE);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_MaxScrollTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.maxScrollTime = static_cast<LONG>(SendMessage(TVM_GETSCROLLTIME, 0, 0));
	}

	*pValue = properties.maxScrollTime;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_MaxScrollTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_MAXSCROLLTIME);
	if((newValue < 100) || (newValue > 5000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.maxScrollTime != newValue) {
		properties.maxScrollTime = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TVM_SETSCROLLTIME, properties.maxScrollTime, 0);
		}
		FireOnChanged(DISPID_EXTVW_MAXSCROLLTIME);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_EXTVW_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXTVW_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_EXTVW_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXTVW_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXTVW_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OLEDragImageStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.oleDragImageStyle;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_OLEDragImageStyle(OLEDragImageStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_OLEDRAGIMAGESTYLE);
	if(properties.oleDragImageStyle != newValue) {
		properties.oleDragImageStyle = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXTVW_OLEDRAGIMAGESTYLE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_PROCESSCONTEXTMENUKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXTVW_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_REGISTERFOROLEDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				//CoLockObjectExternal(static_cast<IDropTarget*>(this), TRUE, FALSE);
				ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
			} else {
				ATLVERIFY(RevokeDragDrop(*this) == S_OK);
				//CoLockObjectExternal(static_cast<IDropTarget*>(this), FALSE, TRUE);
			}
		}
		FireOnChanged(DISPID_EXTVW_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_RichToolTips(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.richToolTips = ((SendMessage(TVM_GETEXTENDEDSTYLE, 0, 0) & TVS_EX_RICHTOOLTIP) == TVS_EX_RICHTOOLTIP);
	}

	*pValue = BOOL2VARIANTBOOL(properties.richToolTips);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_RichToolTips(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_RICHTOOLTIPS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.richToolTips != b) {
		properties.richToolTips = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.richToolTips) {
				SendMessage(TVM_SETEXTENDEDSTYLE, TVS_EX_RICHTOOLTIP, TVS_EX_RICHTOOLTIP);
			} else {
				SendMessage(TVM_SETEXTENDEDSTYLE, TVS_EX_RICHTOOLTIP, 0);
			}
		}
		FireOnChanged(DISPID_EXTVW_RICHTOOLTIPS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_RightToLeft(RightToLeftConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RightToLeftConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.rightToLeft = static_cast<RightToLeftConstants>(0);
		if(GetExStyle() & WS_EX_LAYOUTRTL) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlLayout);
		}
		if(GetStyle() & TVS_RTLREADING) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlText);
		}
	}

	*pValue = properties.rightToLeft;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_RIGHTTOLEFT);
	if(properties.rightToLeft != newValue) {
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			BOOL requiresRecreation = FALSE;
			if(properties.rightToLeft & rtlLayout) {
				if(!(GetExStyle() & WS_EX_LAYOUTRTL)) {
					requiresRecreation = TRUE;
				}
			} else {
				if(GetExStyle() & WS_EX_LAYOUTRTL) {
					requiresRecreation = TRUE;
				}
			}
			if(requiresRecreation) {
				RecreateControlWindow();
			} else {
				if(properties.rightToLeft & rtlText) {
					ModifyStyle(0, TVS_RTLREADING);
					ModifyStyleEx(0, WS_EX_RTLREADING);
				} else {
					ModifyStyle(TVS_RTLREADING, 0);
					ModifyStyleEx(WS_EX_RTLREADING, 0);
				}
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_EXTVW_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_RootItems(ITreeViewItems** ppItems/* = NULL*/)
{
	ATLASSERT_POINTER(ppItems, ITreeViewItems*);
	if(!ppItems) {
		return E_POINTER;
	}

	ClassFactory::InitTreeItems(this, IID_ITreeViewItems, reinterpret_cast<LPUNKNOWN*>(ppItems), TRUE, NULL);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_ScrollBars(ScrollBarsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ScrollBarsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetStyle() & TVS_NOSCROLL) {
			properties.scrollBars = sbNone;
		} else if(GetStyle() & TVS_NOHSCROLL) {
			properties.scrollBars = sbVerticalOnly;
		} else {
			properties.scrollBars = sbNormal;
		}
	}

	*pValue = properties.scrollBars;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_ScrollBars(ScrollBarsConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_SCROLLBARS);
	if(properties.scrollBars != newValue) {
		properties.scrollBars = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			BOOL requiresRecreation = FALSE;
			switch(properties.scrollBars) {
				case sbNone:
					requiresRecreation = TRUE;
					//ModifyStyle(TVS_NOHSCROLL, TVS_NOSCROLL, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case sbVerticalOnly:
					if((GetStyle() & TVS_NOSCROLL) == 0) {
						requiresRecreation = TRUE;
					} else {
						ModifyStyle(TVS_NOSCROLL, TVS_NOHSCROLL, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					}
					break;
				case sbNormal:
					ModifyStyle(TVS_NOSCROLL | TVS_NOHSCROLL, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			if(requiresRecreation) {
				RecreateControlWindow();
			} else {
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_EXTVW_SCROLLBARS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_ShowDragImage(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!dragDropStatus.IsDragging()) {
		return E_FAIL;
	}

	*pValue = BOOL2VARIANTBOOL(dragDropStatus.IsDragImageVisible());
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_ShowDragImage(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_SHOWDRAGIMAGE);
	if(!dragDropStatus.hDragImageList && !dragDropStatus.pDropTargetHelper) {
		return E_FAIL;
	}

	if(newValue == VARIANT_FALSE) {
		dragDropStatus.HideDragImage(FALSE);
	} else {
		dragDropStatus.ShowDragImage(FALSE);
	}

	FireOnChanged(DISPID_EXTVW_SHOWDRAGIMAGE);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_ShowStateImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.showStateImages = ((GetStyle() & TVS_CHECKBOXES) == TVS_CHECKBOXES);
		if(!properties.showStateImages && IsComctl32Version610OrNewer()) {
			DWORD extendedStyle = static_cast<DWORD>(SendMessage(TVM_GETEXTENDEDSTYLE, 0, 0));
			properties.showStateImages = ((extendedStyle & (TVS_EX_PARTIALCHECKBOXES | TVS_EX_DIMMEDCHECKBOXES | TVS_EX_EXCLUSIONCHECKBOXES)) != 0);
		}
	}

	*pValue = BOOL2VARIANTBOOL(properties.showStateImages);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_ShowStateImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_SHOWSTATEIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showStateImages != b) {
		properties.showStateImages = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(IsComctl32Version610OrNewer()) {
				if(properties.showStateImages) {
					if(properties.builtInStateImages == bisiSelectedAndUnselected) {
						DWORD extendedStyle = static_cast<DWORD>(SendMessage(TVM_GETEXTENDEDSTYLE, 0, 0));
						if(extendedStyle & (TVS_EX_PARTIALCHECKBOXES | TVS_EX_DIMMEDCHECKBOXES | TVS_EX_EXCLUSIONCHECKBOXES)) {
							// we can't remove state images dynamically
							RecreateControlWindow();
						} else {
							ModifyStyle(0, TVS_CHECKBOXES);
						}
					} else {
						if(GetStyle() & TVS_CHECKBOXES) {
							// we can't remove TVS_CHECKBOXES dynamically
							RecreateControlWindow();
						} else {
							DWORD newExtendedStyle = 0;
							if(properties.builtInStateImages & bisiIndeterminate) {
								newExtendedStyle |= TVS_EX_PARTIALCHECKBOXES;
							}
							if(properties.builtInStateImages & bisiSelectedDimmed) {
								newExtendedStyle |= TVS_EX_DIMMEDCHECKBOXES;
							}
							if(properties.builtInStateImages & bisiExcluded) {
								newExtendedStyle |= TVS_EX_EXCLUSIONCHECKBOXES;
							}
							DWORD oldExtendedStyle = static_cast<DWORD>(SendMessage(TVM_GETEXTENDEDSTYLE, 0, 0)) & (TVS_EX_PARTIALCHECKBOXES | TVS_EX_DIMMEDCHECKBOXES | TVS_EX_EXCLUSIONCHECKBOXES);
							if(newExtendedStyle != oldExtendedStyle) {
								if(oldExtendedStyle) {
									// we can't call TVM_SETEXTENDEDSTYLE twice, so recreate the control
									RecreateControlWindow();
								} else {
									// TVM_SETEXTENDEDSTYLE was not yet called, call it now
									SendMessage(TVM_SETEXTENDEDSTYLE, (TVS_EX_PARTIALCHECKBOXES | TVS_EX_EXCLUSIONCHECKBOXES | TVS_EX_DIMMEDCHECKBOXES), newExtendedStyle);
								}
							}
						}
					}
				} else {     // properties.showStateImages
					if(GetStyle() & TVS_CHECKBOXES) {
						// we can't remove TVS_CHECKBOXES dynamically
						RecreateControlWindow();
					} else {
						DWORD extendedStyle = static_cast<DWORD>(SendMessage(TVM_GETEXTENDEDSTYLE, 0, 0));
						if(extendedStyle & (TVS_EX_PARTIALCHECKBOXES | TVS_EX_DIMMEDCHECKBOXES | TVS_EX_EXCLUSIONCHECKBOXES)) {
							// we can't remove state images dynamically
							RecreateControlWindow();
						}
					}
				}
			} else {     // IsComctl32Version610OrNewer
				if(properties.showStateImages) {
					ModifyStyle(0, TVS_CHECKBOXES);
				} else {
					if(GetStyle() & TVS_CHECKBOXES) {
						// we can't remove TVS_CHECKBOXES dynamically
						RecreateControlWindow();
					}
				}
			}
		}
		FireOnChanged(DISPID_EXTVW_SHOWSTATEIMAGES);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_ShowToolTips(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.showToolTips = ((GetStyle() & TVS_NOTOOLTIPS) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showToolTips);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_ShowToolTips(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_SHOWTOOLTIPS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showToolTips != b) {
		properties.showToolTips = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.showToolTips) {
				ModifyStyle(TVS_NOTOOLTIPS, 0);
			} else {
				ModifyStyle(0, TVS_NOTOOLTIPS);
			}
		}
		FireOnChanged(DISPID_EXTVW_SHOWTOOLTIPS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_SingleExpand(SingleExpandConstants* pValue)
{
	ATLASSERT_POINTER(pValue, SingleExpandConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetStyle() & TVS_SINGLEEXPAND) {
			if(properties.singleExpand != seWinXPStyle)
				properties.singleExpand = seNormal;
		} else {
			properties.singleExpand = seNone;
		}
	}

	*pValue = properties.singleExpand;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_SingleExpand(SingleExpandConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_SINGLEEXPAND);
	if(properties.singleExpand != newValue) {
		properties.singleExpand = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.singleExpand) {
				case seNone:
					ModifyStyle(TVS_SINGLEEXPAND, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case seNormal:
				case seWinXPStyle:
					ModifyStyle(0, TVS_SINGLEEXPAND, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
		}
		FireOnChanged(DISPID_EXTVW_SINGLEEXPAND);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_SortOrder(SortOrderConstants* pValue)
{
	ATLASSERT_POINTER(pValue, SortOrderConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.sortOrder;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_SortOrder(SortOrderConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_SORTORDER);
	if(properties.sortOrder != newValue) {
		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_ChangingSortOrder(properties.sortOrder, newValue, &cancel);
		if(cancel == VARIANT_FALSE) {
			SortOrderConstants buffer = properties.sortOrder;
			properties.sortOrder = newValue;
			SetDirty(TRUE);
			FireOnChanged(DISPID_EXTVW_SORTORDER);
			Raise_ChangedSortOrder(buffer, properties.sortOrder);
		}
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXTVW_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_TextParsingFlags(TextParsingFunctionConstants parsingFunction, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	switch(parsingFunction) {
		case tpfCompareString:
			*pValue = properties.textParsingFlagsForCompareString;
			break;
		case tpfVarFromStr:
			*pValue = properties.textParsingFlagsForVarFromStr;
			break;
		default:
			return E_INVALIDARG;
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_TextParsingFlags(TextParsingFunctionConstants parsingFunction, LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_TEXTPARSINGFLAGS);
	switch(parsingFunction) {
		case tpfCompareString:
			properties.textParsingFlagsForCompareString = newValue;
			break;
		case tpfVarFromStr:
			properties.textParsingFlagsForVarFromStr = newValue;
			break;
		default:
			return E_INVALIDARG;
	}

	FireOnChanged(DISPID_EXTVW_TEXTPARSINGFLAGS);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_TreeItems(ITreeViewItems** ppItems)
{
	ATLASSERT_POINTER(ppItems, ITreeViewItems*);
	if(!ppItems) {
		return E_POINTER;
	}

	ClassFactory::InitTreeItems(this, IID_ITreeViewItems, reinterpret_cast<LPUNKNOWN*>(ppItems), FALSE, reinterpret_cast<HTREEITEM>(-1));
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_TreeViewStyle(TreeViewStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, TreeViewStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.treeViewStyle = static_cast<TreeViewStyleConstants>(0);
		if(GetStyle() & TVS_HASBUTTONS) {
			properties.treeViewStyle = static_cast<TreeViewStyleConstants>(static_cast<long>(properties.treeViewStyle) | tvsExpandos);
		}
		if(GetStyle() & TVS_HASLINES) {
			properties.treeViewStyle = static_cast<TreeViewStyleConstants>(static_cast<long>(properties.treeViewStyle) | tvsLines);
		}
	}

	*pValue = properties.treeViewStyle;
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_TreeViewStyle(TreeViewStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_TREEVIEWSTYLE);
	if(properties.treeViewStyle != newValue) {
		properties.treeViewStyle = newValue;
		SetDirty(TRUE);

		if(properties.treeViewStyle & tvsLines) {
			// lines and full row select are incompatible
			put_FullRowSelect(VARIANT_FALSE);
		}

		if(IsWindow()) {
			if(properties.treeViewStyle & tvsExpandos) {
				ModifyStyle(0, TVS_HASBUTTONS);
			} else {
				ModifyStyle(TVS_HASBUTTONS, 0);
			}
			if(properties.treeViewStyle & tvsLines) {
				ModifyStyle(0, TVS_HASLINES);
			} else {
				ModifyStyle(TVS_HASLINES, 0);
			}
		}
		FireOnChanged(DISPID_EXTVW_TREEVIEWSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXTVW_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_EXTVW_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::BeginDrag(ITreeViewItemContainer* pDraggedItems, OLE_HANDLE hDragImageList/* = NULL*/, OLE_XPOS_PIXELS* pXHotSpot/* = NULL*/, OLE_YPOS_PIXELS* pYHotSpot/* = NULL*/)
{
	ATLASSUME(pDraggedItems);
	if(!pDraggedItems) {
		return E_POINTER;
	}

	int xHotSpot = 0;
	if(pXHotSpot) {
		xHotSpot = *pXHotSpot;
	}
	int yHotSpot = 0;
	if(pYHotSpot) {
		yHotSpot = *pYHotSpot;
	}
	HRESULT hr = dragDropStatus.BeginDrag(*this, pDraggedItems, static_cast<HIMAGELIST>(LongToHandle(hDragImageList)), &xHotSpot, &yHotSpot);
	SetCapture();
	if(pXHotSpot) {
		*pXHotSpot = xHotSpot;
	}
	if(pYHotSpot) {
		*pYHotSpot = yHotSpot;
	}

	if(dragDropStatus.hDragImageList) {
		ImageList_BeginDrag(dragDropStatus.hDragImageList, 0, xHotSpot, yHotSpot);
		dragDropStatus.dragImageIsHidden = 0;
		ImageList_DragEnter(0, 0, 0);

		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ImageList_DragMove(mousePosition.x, mousePosition.y);
	}
	return hr;
}

STDMETHODIMP ExplorerTreeView::CountVisible(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = static_cast<LONG>(SendMessage(TVM_GETVISIBLECOUNT, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerTreeView::CreateItemContainer(VARIANT items/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, ITreeViewItemContainer** ppContainer/* = NULL*/)
{
	ATLASSERT_POINTER(ppContainer, ITreeViewItemContainer*);
	if(!ppContainer) {
		return E_POINTER;
	}

	*ppContainer = NULL;
	CComObject<TreeViewItemContainer>* pTvwItemContainerObj = NULL;
	CComObject<TreeViewItemContainer>::CreateInstance(&pTvwItemContainerObj);
	pTvwItemContainerObj->AddRef();

	// clone all settings
	pTvwItemContainerObj->SetOwner(this);

	pTvwItemContainerObj->QueryInterface(__uuidof(ITreeViewItemContainer), reinterpret_cast<LPVOID*>(ppContainer));
	pTvwItemContainerObj->Release();

	if(*ppContainer) {
		(*ppContainer)->Add(items);
		RegisterItemContainer(static_cast<IItemContainer*>(pTvwItemContainerObj));
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::EndDrag(VARIANT_BOOL abort)
{
	if(!dragDropStatus.IsDragging()) {
		return E_FAIL;
	}

	KillTimer(timers.ID_DRAGSCROLL);
	KillTimer(timers.ID_DRAGEXPAND);
	ReleaseCapture();
	if(dragDropStatus.hDragImageList) {
		dragDropStatus.HideDragImage(TRUE);
		ImageList_EndDrag();
	}

	HRESULT hr = S_OK;
	if(abort) {
		hr = Raise_AbortedDrag();
	} else {
		hr = Raise_Drop();
	}

	dragDropStatus.EndDrag();
	Invalidate();

	return hr;
}

STDMETHODIMP ExplorerTreeView::EndLabelEdit(VARIANT_BOOL discard)
{
	if(IsWindow()) {
		if(SendMessage(TVM_ENDEDITLABELNOW, VARIANTBOOL2BOOL(discard), 0)) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerTreeView::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP ExplorerTreeView::GetInsertMarkPosition(InsertMarkPositionConstants* pRelativePosition, ITreeViewItem** ppTreeItem)
{
	ATLASSERT_POINTER(pRelativePosition, InsertMarkPositionConstants);
	if(!pRelativePosition) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(ppTreeItem, ITreeViewItem*);
	if(!ppTreeItem) {
		return E_POINTER;
	}

	if(insertMark.hItem) {
		if(insertMark.afterItem) {
			*pRelativePosition = impAfter;
		} else {
			*pRelativePosition = impBefore;
		}
		ClassFactory::InitTreeItem(insertMark.hItem, this, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppTreeItem));
	} else {
		*ppTreeItem = NULL;
		*pRelativePosition = impNowhere;
	}

	return S_OK;
}

STDMETHODIMP ExplorerTreeView::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, ITreeViewItem** ppHitItem)
{
	ATLASSERT_POINTER(ppHitItem, ITreeViewItem*);
	if(!ppHitItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		UINT hitTestFlags = static_cast<UINT>(*pHitTestDetails);
		HTREEITEM hItem = HitTest(x, y, &hitTestFlags, TRUE);

		if(pHitTestDetails) {
			*pHitTestDetails = static_cast<HitTestConstants>(hitTestFlags);
		}
		ClassFactory::InitTreeItem(hItem, this, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppHitItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerTreeView::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::OLEDrag(LONG* pDataObject/* = NULL*/, OLEDropEffectConstants supportedEffects/* = odeCopyOrMove*/, OLE_HANDLE hWndToAskForDragImage/* = -1*/, ITreeViewItemContainer* pDraggedItems/* = NULL*/, LONG itemCountToDisplay/* = -1*/, OLEDropEffectConstants* pPerformedEffects/* = NULL*/)
{
	if(supportedEffects == odeNone) {
		// don't waste time
		return S_OK;
	}
	if(hWndToAskForDragImage == -1) {
		ATLASSUME(pDraggedItems);
		if(!pDraggedItems) {
			return E_INVALIDARG;
		}
	}
	ATLASSERT_POINTER(pDataObject, LONG);
	LONG dummy = NULL;
	if(!pDataObject) {
		pDataObject = &dummy;
	}
	ATLASSERT_POINTER(pPerformedEffects, OLEDropEffectConstants);
	OLEDropEffectConstants performedEffect = odeNone;
	if(!pPerformedEffects) {
		pPerformedEffects = &performedEffect;
	}

	HWND hWnd = NULL;
	if(hWndToAskForDragImage == -1) {
		hWnd = *this;
	} else {
		hWnd = static_cast<HWND>(LongToHandle(hWndToAskForDragImage));
	}

	CComPtr<IOLEDataObject> pOLEDataObject = NULL;
	CComPtr<IDataObject> pDataObjectToUse = NULL;
	if(*pDataObject) {
		pDataObjectToUse = *reinterpret_cast<LPDATAOBJECT*>(pDataObject);

		CComObject<TargetOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->Attach(pDataObjectToUse);
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->Release();
	} else {
		CComObject<SourceOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<SourceOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->SetOwner(this);
		if(itemCountToDisplay == -1) {
			if(pDraggedItems) {
				if(flags.usingThemes && RunTimeHelper::IsVista()) {
					pDraggedItems->Count(&pOLEDataObjectObj->properties.numberOfItemsToDisplay);
				}
			}
		} else if(itemCountToDisplay >= 0) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				pOLEDataObjectObj->properties.numberOfItemsToDisplay = itemCountToDisplay;
			}
		}
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&pDataObjectToUse));
		pOLEDataObjectObj->Release();
	}
	ATLASSUME(pDataObjectToUse);
	pDataObjectToUse->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&dragDropStatus.pSourceDataObject));
	CComQIPtr<IDropSource, &IID_IDropSource> pDragSource(this);

	if(pDraggedItems) {
		pDraggedItems->Clone(&dragDropStatus.pDraggedItems);
	}
	POINT mousePosition = {0};
	GetCursorPos(&mousePosition);
	ScreenToClient(&mousePosition);

	if(properties.supportOLEDragImages) {
		IDragSourceHelper* pDragSourceHelper = NULL;
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pDragSourceHelper));
		if(pDragSourceHelper) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				IDragSourceHelper2* pDragSourceHelper2 = NULL;
				pDragSourceHelper->QueryInterface(IID_IDragSourceHelper2, reinterpret_cast<LPVOID*>(&pDragSourceHelper2));
				if(pDragSourceHelper2) {
					pDragSourceHelper2->SetFlags(DSH_ALLOWDROPDESCRIPTIONTEXT);
					// this was the only place we actually use IDragSourceHelper2
					pDragSourceHelper->Release();
					pDragSourceHelper = static_cast<IDragSourceHelper*>(pDragSourceHelper2);
				}
			}

			HRESULT hr = pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
			if(FAILED(hr)) {
				/* This happens if full window dragging is deactivated. Actually, InitializeFromWindow() contains a
				   fallback mechanism for this case. This mechanism retrieves the passed window's class name and
				   builds the drag image using TVM_CREATEDRAGIMAGE if it's SysTreeView32, LVM_CREATEDRAGIMAGE if
				   it's SysListView32 and so on. Our class name is ExplorerTreeView[U|A], so we're doomed.
				   So how can we have drag images anyway? Well, we use a very ugly hack: We temporarily activate
				   full window dragging. */
				BOOL fullWindowDragging;
				SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0, &fullWindowDragging, 0);
				if(!fullWindowDragging) {
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, TRUE, NULL, 0);
					pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, FALSE, NULL, 0);
				}
			}

			if(pDragSourceHelper) {
				pDragSourceHelper->Release();
			}
		}
	}

	if(IsLeftMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_LBUTTON;
	} else if(IsRightMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_RBUTTON;
	}
	if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
		dragDropStatus.useItemCountLabelHack = TRUE;
	}

	if(pOLEDataObject) {
		Raise_OLEStartDrag(pOLEDataObject);
	}
	HRESULT hr = DoDragDrop(pDataObjectToUse, pDragSource, supportedEffects, reinterpret_cast<LPDWORD>(pPerformedEffects));
	KillTimer(timers.ID_DRAGSCROLL);
	KillTimer(timers.ID_DRAGEXPAND);
	if((hr == DRAGDROP_S_DROP) && pOLEDataObject) {
		Raise_OLECompleteDrag(pOLEDataObject, *pPerformedEffects);
	}

	dragDropStatus.Reset();
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::Refresh(void)
{
	if(IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);
		Invalidate();
		UpdateWindow();
		dragDropStatus.ShowDragImage(TRUE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP ExplorerTreeView::SetInsertMarkPosition(InsertMarkPositionConstants relativePosition, ITreeViewItem* pTreeItem)
{
	HTREEITEM hItem = NULL;
	if(pTreeItem) {
		OLE_HANDLE h = NULL;
		pTreeItem->get_Handle(&h);
		hItem = static_cast<HTREEITEM>(LongToHandle(h));
	}

	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);
		switch(relativePosition) {
			case impNowhere:
				if(SendMessage(TVM_SETINSERTMARK, FALSE, 0)) {
					hr = S_OK;
				}
				break;
			case impBefore:
				if(SendMessage(TVM_SETINSERTMARK, FALSE, reinterpret_cast<LPARAM>(hItem))) {
					hr = S_OK;
				}
				break;
			case impAfter:
				if(SendMessage(TVM_SETINSERTMARK, TRUE, reinterpret_cast<LPARAM>(hItem))) {
					hr = S_OK;
				}
				break;
			case impDontChange:
				hr = S_OK;
				break;
		}
		dragDropStatus.ShowDragImage(TRUE);
	}

	return hr;
}

STDMETHODIMP ExplorerTreeView::SortItems(SortByConstants firstCriterion/* = sobShell*/, SortByConstants secondCriterion/* = sobText*/, SortByConstants thirdCriterion/* = sobNone*/, SortByConstants fourthCriterion/* = sobNone*/, SortByConstants fifthCriterion/* = sobNone*/, VARIANT_BOOL recurse/* = VARIANT_FALSE*/, VARIANT_BOOL caseSensitive/* = VARIANT_FALSE*/)
{
	return SortSubItems(TVI_ROOT, firstCriterion, secondCriterion, thirdCriterion, fourthCriterion, fifthCriterion, VARIANTBOOL2BOOL(recurse), VARIANTBOOL2BOOL(caseSensitive));
}


LRESULT ExplorerTreeView::OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!(properties.disabledEvents & deTreeKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_KeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT ExplorerTreeView::OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y);

	return 0;
}

LRESULT ExplorerTreeView::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		/* SysTreeView32 already sent a WM_NOTIFYFORMAT/NF_QUERY message to the parent window, but
		   unfortunately our reflection object did not yet know where to reflect this message to. So we ask
		   SysTreeView32 to send the message again. */
		SendMessage(WM_NOTIFYFORMAT, reinterpret_cast<WPARAM>(reinterpret_cast<LPCREATESTRUCT>(lParam)->hwndParent), NF_REQUERY);

		Raise_RecreatedControlWindow(HandleToLong(static_cast<HWND>(*this)));
	}
	return lr;
}

LRESULT ExplorerTreeView::OnCtlColorEdit(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	SetBkColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.editBackColor));
	SetTextColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.editForeColor));
	return reinterpret_cast<LRESULT>(hEditBackColorBrush);
}

LRESULT ExplorerTreeView::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(HandleToLong(static_cast<HWND>(*this)));

	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerTreeView::OnInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	IMEModeConstants ime = GetEffectiveIMEMode();
	if((ime != imeNoControl) && (GetFocus() == *this)) {
		// we've the focus, so configure the IME
		SetCurrentIMEContextMode(*this, ime);
	} else {
		ime = GetEffectiveIMEMode_Edit();
		if((ime != imeNoControl) && (GetFocus() == containedEdit)) {
			// we've the focus, so configure the IME
			SetCurrentIMEContextMode(containedEdit, ime);
		}
	}
	return lr;
}

LRESULT ExplorerTreeView::OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = 0;
	// required by OnKeyDownNotification
	lastKeyDownLParam = lParam;

	if(!(properties.disabledEvents & deTreeKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}

	wasHandled = FALSE;
	if((cachedSettings.hStateImageList != NULL) && ((GetAsyncKeyState(VK_SPACE) & 0x8000) == 0x8000) && !(GetAsyncKeyState(VK_CONTROL) & 0x8000)) {
		stateImageChangeFlags.reasonForPotentialStateImageChange = siccbKeyboard;
		if(!IsComctl32Version610OrNewer()) {
			// toggle the caret's state image
			HTREEITEM hCaretItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0));
			if(hCaretItem) {
				// retrieve the old state image's index
				TVITEMEX item = {0};
				item.hItem = hCaretItem;
				item.mask = TVIF_HANDLE | TVIF_STATE;
				if(SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
					int previousStateImage = (item.state & TVIS_STATEIMAGEMASK) >> 12;
					// How many state images do we have?
					int c = ImageList_GetImageCount(cachedSettings.hStateImageList);
					// calc the new state image
					int newStateImage = (previousStateImage + 1) % c;
					if(newStateImage == 0) {
						// state images are one-based
						newStateImage = 1;
					}
					int i = newStateImage;

					CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hCaretItem, this);
					VARIANT_BOOL cancel = VARIANT_FALSE;
					Raise_ItemStateImageChanging(pTvwItem, previousStateImage, reinterpret_cast<LONG*>(&newStateImage), siccbKeyboard, &cancel);
					wasHandled = TRUE;
					if(cancel != VARIANT_FALSE) {
						return 0;
					}

					lr = DefWindowProc(message, wParam, lParam);
					if(i != newStateImage) {
						// explicitly set the new state image
						item.state = INDEXTOSTATEIMAGEMASK(newStateImage);
						item.stateMask = TVIS_STATEIMAGEMASK;
						if(SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
							Raise_ItemStateImageChanged(pTvwItem, previousStateImage, newStateImage, siccbKeyboard);
						}
					} else {
						// simply raise the event
						Raise_ItemStateImageChanged(pTvwItem, previousStateImage, newStateImage, siccbKeyboard);
					}
				}
			}
		}
	}

	if(!wasHandled) {
		wasHandled = TRUE;
		lr = DefWindowProc(message, wParam, lParam);
	}

	return lr;
}

LRESULT ExplorerTreeView::OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deTreeKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerTreeView::OnLButtonDblClk(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(IsComctl32Version610OrNewer()) {
		/* Starting with version 6.10 of comctl32.dll, double-clicking a state image toggles it twice.
		   Because the state image change reason is reset by Raise_ItemStateImageChanged, we have to set
		   the reason to siccbMouse again for the 2nd series of events. */
		UINT hitTestFlags = 0;
		HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &hitTestFlags, TRUE);
		if(hitTestFlags & TVHT_ONITEMSTATEICON) {
			stateImageChangeFlags.reasonForPotentialStateImageChange = siccbMouse;
		}
	}

	if(!(properties.disabledEvents & deTreeClickEvents)) {
		/* SysTreeView32 has a bug: It sends the NM_DBLCLK notification only if an item was hit. Thus we
		   use WM_LBUTTONDBLCLK for the other situations (using it for all situations causes problems with
		   label-editing - see revision 35). WM_LBUTTONDBLCLK is sent before NM_DBLCLK. To not raise the
		   event 2 times, we simply set a flag to raise it in OnLButtonUp. It'll be cleared by
		   Raise_DblClick. */
		mouseStatus.raiseDblClkEvent = TRUE;
	}

	return 0;
}

LRESULT ExplorerTreeView::OnLButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = 0;

	HTREEITEM hItemStateImageClicked = NULL;
	BOOL allowingDragDrop = ((GetStyle() & TVS_DISABLEDRAGDROP) == 0);

	if(!SendMessage(TVM_GETEDITCONTROL, 0, 0)) {
		UINT hitTestFlags = 0;
		// retrieve the item whose row the click was within
		HTREEITEM hClickedRow = HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &hitTestFlags, TRUE);
		if(hitTestFlags & TVHT_ONITEMSTATEICON) {
			stateImageChangeFlags.hClickedItem = NULL;
		}

		if(hitTestFlags & TVHT_ONITEMSTATEICON) {
			// the item's state image was clicked
			if(allowingDragDrop) {
				// will be handled in OnLButtonUp
				stateImageChangeFlags.hClickedItem = hClickedRow;
			} else {
				hItemStateImageClicked = hClickedRow;
			}
		}
	}

	if(!(properties.disabledEvents & deTreeMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	int previousStateImage = 0;
	int newStateImage = 0;
	int newStateImageToUse = newStateImage;
	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(!allowingDragDrop && (hItemStateImageClicked != NULL) && (cachedSettings.hStateImageList != NULL)) {
		stateImageChangeFlags.reasonForPotentialStateImageChange = siccbMouse;
		if(!IsComctl32Version610OrNewer()) {
			// the item's state image will be toggled - retrieve the previous state image's index
			TVITEMEX item = {0};
			item.hItem = hItemStateImageClicked;
			item.mask = TVIF_HANDLE | TVIF_STATE;
			if(SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
				previousStateImage = (item.state & TVIS_STATEIMAGEMASK) >> 12;
				// How many state images do we have?
				int c = ImageList_GetImageCount(cachedSettings.hStateImageList);
				// calculate the new state image
				newStateImage = (previousStateImage + 1) % c;
				if(newStateImage == 0) {
					// state images are one-based
					newStateImage = 1;
				}
				newStateImageToUse = newStateImage;

				CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItemStateImageClicked, this);
				Raise_ItemStateImageChanging(pTvwItem, previousStateImage, reinterpret_cast<LONG*>(&newStateImageToUse), siccbMouse, &cancel);
				if(cancel != VARIANT_FALSE) {
					newStateImageToUse = previousStateImage;
				}
			}
		}
	}

	lr = DefWindowProc(message, wParam, lParam);

	if(!allowingDragDrop && (hItemStateImageClicked != NULL) && (cachedSettings.hStateImageList != NULL)) {
		stateImageChangeFlags.reasonForPotentialStateImageChange = siccbMouse;
		if(!IsComctl32Version610OrNewer()) {
			// finish the state image change
			CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItemStateImageClicked, this);
			if(newStateImageToUse != newStateImage) {
				TVITEMEX item = {0};
				item.hItem = hItemStateImageClicked;
				item.mask = TVIF_HANDLE | TVIF_STATE;
				// explicitly set the new state image
				item.state = INDEXTOSTATEIMAGEMASK(newStateImageToUse);
				item.stateMask = TVIS_STATEIMAGEMASK;
				if(SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
					if(cancel == VARIANT_FALSE) {
						Raise_ItemStateImageChanged(pTvwItem, previousStateImage, newStateImageToUse, siccbMouse);
					}
				}
			} else {
				// simply raise the event
				if(cancel == VARIANT_FALSE) {
					Raise_ItemStateImageChanged(pTvwItem, previousStateImage, newStateImageToUse, siccbMouse);
				}
			}
		}
	}

	return lr;
}

LRESULT ExplorerTreeView::OnLButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = 0;

	HTREEITEM hItemStateImageClicked = NULL;
	BOOL allowingDragDrop = ((GetStyle() & TVS_DISABLEDRAGDROP) == 0);
	// Did we call this method ourselves?
	BOOL eatMessage = (message == 0);

	if(!SendMessage(TVM_GETEDITCONTROL, 0, 0)) {
		UINT hitTestFlags = 0;
		// retrieve the item whose row the click was within
		HTREEITEM hClickedRow = HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &hitTestFlags, TRUE);
		if(hitTestFlags & TVHT_ONITEMSTATEICON) {
			if(stateImageChangeFlags.hClickedItem == hClickedRow) {
				hItemStateImageClicked = hClickedRow;
			}
		}
	}

	if(!(properties.disabledEvents & deTreeClickEvents) && mouseStatus.raiseDblClkEvent) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button |= 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_DblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}
	if(!(properties.disabledEvents & deTreeMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	int previousStateImage = 0;
	int newStateImage = 0;
	int newStateImageToUse = newStateImage;
	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(allowingDragDrop && (hItemStateImageClicked != NULL) && (cachedSettings.hStateImageList != NULL)) {
		stateImageChangeFlags.reasonForPotentialStateImageChange = siccbMouse;
		if(!IsComctl32Version610OrNewer()) {
			// the item's state image will be toggled - retrieve the previous state image's index
			TVITEMEX item = {0};
			item.hItem = hItemStateImageClicked;
			item.mask = TVIF_HANDLE | TVIF_STATE;
			if(SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
				previousStateImage = (item.state & TVIS_STATEIMAGEMASK) >> 12;
				// How many state images do we have?
				int c = ImageList_GetImageCount(cachedSettings.hStateImageList);
				// calculate the new state image
				newStateImage = (previousStateImage + 1) % c;
				if(newStateImage == 0) {
					// state images are one-based
					newStateImage = 1;
				}
				newStateImageToUse = newStateImage;

				CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItemStateImageClicked, this);
				Raise_ItemStateImageChanging(pTvwItem, previousStateImage, reinterpret_cast<LONG*>(&newStateImageToUse), siccbMouse, &cancel);
				if(cancel != VARIANT_FALSE) {
					newStateImageToUse = previousStateImage;
				}
			}
		}
	}

	if(!eatMessage) {
		lr = DefWindowProc(message, wParam, lParam);
	}

	if(allowingDragDrop && (hItemStateImageClicked != NULL) && (cachedSettings.hStateImageList != NULL)) {
		stateImageChangeFlags.reasonForPotentialStateImageChange = siccbMouse;
		if(!IsComctl32Version610OrNewer()) {
			// finish the state image change
			CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItemStateImageClicked, this);
			if(newStateImageToUse != newStateImage) {
				if(newStateImageToUse == 0) {
					// for some reason the change must be delayed
					stateImageChangeFlags.hItemToResetStateImageFor = hItemStateImageClicked;
					stateImageChangeFlags.previousStateImage = previousStateImage;
					stateImageChangeFlags.changeHasBeenCanceled = VARIANTBOOL2BOOL(cancel);
					SetTimer(timers.ID_STATEIMAGECHANGEHELPER, timers.INT_STATEIMAGECHANGEHELPER);
				} else {
					TVITEMEX item = {0};
					item.hItem = hItemStateImageClicked;
					item.mask = TVIF_HANDLE | TVIF_STATE;
					// explicitly set the new state image
					if(eatMessage && (newStateImageToUse >= 1)) {
						/* We use a trick here: Since SysTreeView32 will change the state image *after* we've corrected
						   it to 'newStateImageToUse', we correct it to 'newStateImageToUse - 1' instead. */
						--newStateImageToUse;
					}
					item.state = INDEXTOSTATEIMAGEMASK(newStateImageToUse);
					item.stateMask = TVIS_STATEIMAGEMASK;
					if(SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
						if(cancel == VARIANT_FALSE) {
							Raise_ItemStateImageChanged(pTvwItem, previousStateImage, newStateImageToUse, siccbMouse);
						}
					}
				}
			} else {
				// simply raise the event
				if(cancel == VARIANT_FALSE) {
					Raise_ItemStateImageChanged(pTvwItem, previousStateImage, newStateImageToUse, siccbMouse);
				}
			}
		}
	}

	stateImageChangeFlags.hClickedItem = NULL;
	return lr;
}

LRESULT ExplorerTreeView::OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT ExplorerTreeView::OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnMouseActivate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	/* Hide the message from CComControl, so that the control doesn't go into label-editing mode if the
	   caret item was clicked. */
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerTreeView::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);
	}

	return 0;
}

LRESULT ExplorerTreeView::OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	if(dragDropStatus.IsDragging()) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_DragMouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(message == WM_MOUSEHWHEEL) {
			// wParam and lParam seem to be 0
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			GetCursorPos(&mousePosition);
		} else {
			WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		}
		ScreenToClient(&mousePosition);
		Raise_MouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT ExplorerTreeView::OnNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	switch(reinterpret_cast<LPNMHDR>(lParam)->code) {
		case TTN_GETDISPINFOA:
			return OnToolTipGetDispInfoNotificationA(message, wParam, lParam, wasHandled);
			break;

		case TTN_GETDISPINFOW:
			return OnToolTipGetDispInfoNotificationW(message, wParam, lParam, wasHandled);
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT ExplorerTreeView::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerTreeView::OnParentNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = S_OK;
	wasHandled = FALSE;

	switch(wParam) {
		case WM_CREATE: {
			lr = DefWindowProc(message, wParam, lParam);
			wasHandled = TRUE;

			LPTSTR pBuffer = new TCHAR[MAX_PATH + 1];
			GetClassName(reinterpret_cast<HWND>(lParam), pBuffer, MAX_PATH);
			if(lstrcmp(pBuffer, reinterpret_cast<LPCTSTR>(&WC_EDIT)) == 0) {
				Raise_CreatedEditControlWindow(static_cast<LONG>(lParam));
			}
			delete[] pBuffer;
			break;
		}
	}
	return lr;
}

LRESULT ExplorerTreeView::OnRButtonDblClk(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeClickEvents)) {
		/* SysTreeView32 has a bug: It sends the NM_RDBLCLK notification only if an item was hit. Thus we
		   use WM_RBUTTONDBLCLK for the other situations (using it for all situations causes problems with
		   label-editing - see revision 35). WM_RBUTTONDBLCLK is sent before NM_RDBLCLK. To not raise the
		   event 2 times, we simply set a flag to raise it in OnRButtonUp. It'll be removed by
		   Raise_RDblClick. */
		mouseStatus.raiseDblClkEvent = TRUE;
	}

	return 0;
}

LRESULT ExplorerTreeView::OnRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT ExplorerTreeView::OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeClickEvents) && mouseStatus.raiseDblClkEvent) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button |= 2/*MouseButtonConstants.vbRightButton*/;
		Raise_RDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}
	if(!(properties.disabledEvents & deTreeMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_EXTVW_FONT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!properties.font.dontGetFontObject) {
		// This message wasn't sent by ourselves, so we have to retrieve the new font.pFontDisp object.
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(reinterpret_cast<HFONT>(wParam));
		properties.font.owningFontResource = FALSE;
		properties.useSystemFont = FALSE;
		properties.font.StopWatching();

		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(!properties.font.currentFont.IsNull()) {
			LOGFONT logFont = {0};
			int bytes = properties.font.currentFont.GetLogFont(&logFont);
			if(bytes) {
				FONTDESC font = {0};
				CT2OLE converter(logFont.lfFaceName);

				HDC hDC = GetDC();
				if(hDC) {
					LONG fontHeight = logFont.lfHeight;
					if(fontHeight < 0) {
						fontHeight = -fontHeight;
					}

					int pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
					ReleaseDC(hDC);
					font.cySize.Lo = fontHeight * 720000 / pixelsPerInch;
					font.cySize.Hi = 0;

					font.lpstrName = converter;
					font.sWeight = static_cast<SHORT>(logFont.lfWeight);
					font.sCharset = logFont.lfCharSet;
					font.fItalic = logFont.lfItalic;
					font.fUnderline = logFont.lfUnderline;
					font.fStrikethrough = logFont.lfStrikeOut;
				}
				font.cbSizeofstruct = sizeof(FONTDESC);

				OleCreateFontIndirect(&font, IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
			}
		}
		properties.font.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_EXTVW_FONT);
	}

	return lr;
}

LRESULT ExplorerTreeView::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == 71216) {
		// the message was sent by ourselves
		lParam = 0;
		if(wParam) {
			// We're gonna activate redrawing - does the client allow this?
			if(properties.dontRedraw) {
				// no, so eat this message
				return 0;
			}
		}
	} else {
		// TODO: Should we really do this?
		properties.dontRedraw = !wParam;
	}

	return DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerTreeView::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam == SPI_SETICONTITLELOGFONT) {
		if(properties.useSystemFont) {
			ApplyFont();
			//Invalidate();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerTreeView::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerTreeView::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_DELAYEDCARETCHANGE:
			// Is it time to raise our delayed CaretChanged event?
			if(caretChangedDelayStatus.DelayTimePassed(properties.caretChangedDelayTime)) {
				// yes, it is
				caretChangedDelayStatus.ResetTimeOfCaretChange();
				KillTimer(timers.ID_DELAYEDCARETCHANGE);
				if(caretChangedDelayStatus.hNewCaretItem != caretChangedDelayStatus.hPreviousCaretItem) {
					Raise_CaretChanged(caretChangedDelayStatus.hPreviousCaretItem, caretChangedDelayStatus.hNewCaretItem, cccbKeyboard, FALSE);
				}
			}
			break;

		case timers.ID_REDRAW:
			KillTimer(timers.ID_REDRAW);
			hBuiltInStateImageList = reinterpret_cast<HIMAGELIST>(SendMessage(TVM_GETIMAGELIST, TVSIL_STATE, 0));
			cachedSettings.hStateImageList = hBuiltInStateImageList;
			RedrawWindow();
			break;

		case timers.ID_DRAGSCROLL:
			AutoScroll();
			break;

		case timers.ID_DRAGEXPAND:
			KillTimer(timers.ID_DRAGEXPAND);
			if(dragDropStatus.hLastDropTarget && !IsExpanded(dragDropStatus.hLastDropTarget, FALSE)) {
				SendMessage(TVM_EXPAND, TVE_EXPAND, reinterpret_cast<LPARAM>(dragDropStatus.hLastDropTarget));
			}
			break;

		case timers.ID_EDITLABEL:
			KillTimer(timers.ID_EDITLABEL);
			if(labelEditStatus.watchingForEdit) {
				// label-edit the caret
				HTREEITEM hItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0));
				if(hItem) {
					SendMessage(TVM_EDITLABEL, 0, reinterpret_cast<LPARAM>(hItem));
				}
				labelEditStatus.watchingForEdit = FALSE;
			}
			break;

		case timers.ID_STATEIMAGECHANGEHELPER: {
			KillTimer(timers.ID_STATEIMAGECHANGEHELPER);

			TVITEMEX item = {0};
			item.hItem = stateImageChangeFlags.hItemToResetStateImageFor;
			item.mask = TVIF_HANDLE | TVIF_STATE;
			// explicitly set the new state image
			item.state = INDEXTOSTATEIMAGEMASK(0);
			item.stateMask = TVIS_STATEIMAGEMASK;
			if(SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
				if(!stateImageChangeFlags.changeHasBeenCanceled) {
					CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(stateImageChangeFlags.hItemToResetStateImageFor, this);
					Raise_ItemStateImageChanged(pTvwItem, stateImageChangeFlags.previousStateImage, 0, siccbMouse);
				}
			}
			stateImageChangeFlags.hItemToResetStateImageFor = NULL;
			break;
		}

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT ExplorerTreeView::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	CRect windowRectangle = m_rcPos;
	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	              not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	              even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull() && pDetails->x != 0 && pDetails->y != 0)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerTreeView::OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_XDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT ExplorerTreeView::OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnReflectedNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	switch(reinterpret_cast<LPNMHDR>(lParam)->code) {
		case NM_CUSTOMDRAW:
			return OnCustomDrawNotification(message, wParam, lParam, wasHandled);
			break;
		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT ExplorerTreeView::OnReflectedNotifyFormat(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == NF_QUERY) {
		#ifdef UNICODE
			return NFR_UNICODE;
		#else
			return NFR_ANSI;
		#endif
	}
	return 0;
}

#ifdef INCLUDESHELLBROWSERINTERFACE
	LRESULT ExplorerTreeView::OnHandshake(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
	{
		LPSHBHANDSHAKE pDetails = reinterpret_cast<LPSHBHANDSHAKE>(lParam);
		ATLASSERT_POINTER(pDetails, SHBHANDSHAKE);
		if(!pDetails) {
			return FALSE;
		}

		pDetails->processed = TRUE;
		if(pDetails->structSize < CCSIZEOF_STRUCT(SHBHANDSHAKE, pCtlInterface)) {
			return FALSE;
		}

		if(pDetails->pMessageHook && shellBrowserInterface.pMessageListener && shellBrowserInterface.pInternalMessageListener) {
			// there's already a ShellBrowser attached - fail
			pDetails->errorCode = E_FAIL;
			return FALSE;
		}

		if(pDetails->pCtlBuildNumber) {
			*(pDetails->pCtlBuildNumber) = VERSION_BUILD;
		}
		if(pDetails->pCtlSupportsUnicode) {
			#ifdef UNICODE
				*(pDetails->pCtlSupportsUnicode) = TRUE;
			#else
				*(pDetails->pCtlSupportsUnicode) = FALSE;
			#endif
		}
		if(pDetails->pCtlInterface) {
			pDetails->errorCode = QueryInterface(IID_IUnknown, pDetails->pCtlInterface);
		} else {
			pDetails->errorCode = S_OK;
		}
		if(SUCCEEDED(pDetails->errorCode)) {
			shellBrowserInterface.pMessageListener = pDetails->pMessageHook;
			shellBrowserInterface.pInternalMessageListener = pDetails->pShBInternalMessageHook;
			if(pDetails->ppCtlInternalMessageHook) {
				*(pDetails->ppCtlInternalMessageHook) = this;
			}
		}
		return SUCCEEDED(pDetails->errorCode);
	}
#endif

LRESULT ExplorerTreeView::OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	BOOL succeeded = FALSE;
	BOOL useVistaDragImage = FALSE;
	if(dragDropStatus.pDraggedItems) {
		if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
			succeeded = CreateVistaOLEDragImage(dragDropStatus.pDraggedItems, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
			useVistaDragImage = succeeded;
		}
		if(!succeeded) {
			// use a legacy drag image as fallback
			succeeded = CreateLegacyOLEDragImage(dragDropStatus.pDraggedItems, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
		}

		if(succeeded && RunTimeHelper::IsVista()) {
			FORMATETC format = {0};
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			STGMEDIUM medium = {0};
			medium.tymed = TYMED_HGLOBAL;
			medium.hGlobal = GlobalAlloc(GPTR, sizeof(BOOL));
			if(medium.hGlobal) {
				LPBOOL pUseVistaDragImage = reinterpret_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				*pUseVistaDragImage = useVistaDragImage;
				GlobalUnlock(medium.hGlobal);

				dragDropStatus.pSourceDataObject->SetData(&format, &medium, TRUE);
			}
		}
	}

	wasHandled = succeeded;
	// TODO: Why do we have to return FALSE to have the set offset not ignored if a Vista drag image is used?
	return succeeded && !useVistaDragImage;
}

LRESULT ExplorerTreeView::OnCreateDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return reinterpret_cast<LRESULT>(CreateLegacyDragImage(reinterpret_cast<HTREEITEM>(lParam), NULL, NULL));
}

LRESULT ExplorerTreeView::OnDeleteItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = FALSE;

	if(!lParam || (reinterpret_cast<HTREEITEM>(lParam) == TVI_ROOT)) {
		// all items are removed
		#ifdef USE_STL
			std::vector<HTREEITEM>::iterator iter = std::find(itemDeletionsToIgnore.begin(), itemDeletionsToIgnore.end(), TVI_ROOT);
			BOOL fireEvents = (iter == itemDeletionsToIgnore.end());
		#else
			BOOL fireEvents = FALSE;
			for(size_t i = 0; i < itemDeletionsToIgnore.GetCount(); ++i) {
				if(itemDeletionsToIgnore[i] == TVI_ROOT) {
					fireEvents = TRUE;
					break;
				}
			}
		#endif

		VARIANT_BOOL cancel = VARIANT_FALSE;
		if(!(properties.disabledEvents & deItemDeletionEvents)) {
			if(fireEvents) {
				Raise_RemovingItem(NULL, &cancel);
			}
		}
		if(cancel == VARIANT_FALSE) {
			if(!(properties.disabledEvents & deTreeMouseEvents)) {
				if(hItemUnderMouse) {
					// we're removing the item below the mouse cursor
					CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItemUnderMouse, this);
					DWORD position = GetMessagePos();
					POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
					ScreenToClient(&mousePosition);
					SHORT button = 0;
					SHORT shift = 0;
					WPARAM2BUTTONANDSHIFT(-1, button, shift);
					UINT hitTestDetails = 0;
					HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
					Raise_ItemMouseLeave(pTvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					hItemUnderMouse = NULL;
				}
			}

			// finally pass the message to the treeview
			flags.removingAllItems = TRUE;
			lr = DefWindowProc(message, wParam, lParam);
			flags.removingAllItems = FALSE;
			if(lr) {
				if(fireEvents) {
					if(!(properties.disabledEvents & deItemDeletionEvents)) {
						Raise_RemovedItem(NULL);
					}
				}
			}

			RemoveItemFromItemContainers(0);
			#ifdef USE_STL
				itemHandles.clear();
				itemParams.clear();
			#else
				itemHandles.RemoveAll();
				itemParams.RemoveAll();
			#endif
		}

	} else if(ValidateItemHandle(reinterpret_cast<HTREEITEM>(lParam))) {
		// a single item is being removed
		#ifdef USE_STL
			std::vector<HTREEITEM>::iterator iter = std::find(itemDeletionsToIgnore.begin(), itemDeletionsToIgnore.end(), reinterpret_cast<HTREEITEM>(lParam));
			BOOL fireEvents = (iter == itemDeletionsToIgnore.end());
		#else
			BOOL fireEvents = FALSE;
			for(size_t i = 0; i < itemDeletionsToIgnore.GetCount(); ++i) {
				if(itemDeletionsToIgnore[i] == reinterpret_cast<HTREEITEM>(lParam)) {
					fireEvents = TRUE;
					break;
				}
			}
		#endif

		VARIANT_BOOL cancel = VARIANT_FALSE;
		CComPtr<ITreeViewItem> pTvwItemItf = NULL;
		CComObject<TreeViewItem>* pTvwItemObj = NULL;
		if(!(properties.disabledEvents & deItemDeletionEvents) && fireEvents) {
			if(lParam) {
				CComObject<TreeViewItem>::CreateInstance(&pTvwItemObj);
				pTvwItemObj->AddRef();
				pTvwItemObj->SetOwner(this);
				pTvwItemObj->Attach(reinterpret_cast<HTREEITEM>(lParam));
				pTvwItemObj->QueryInterface(IID_ITreeViewItem, reinterpret_cast<LPVOID*>(&pTvwItemItf));
				pTvwItemObj->Release();
			}
			Raise_RemovingItem(pTvwItemItf, &cancel);
		}

		if(cancel == VARIANT_FALSE) {
			CComPtr<IVirtualTreeViewItem> pVTvwItemItf = NULL;
			if(!(properties.disabledEvents & deItemDeletionEvents) && fireEvents) {
				CComObject<VirtualTreeViewItem>* pVTvwItemObj = NULL;
				CComObject<VirtualTreeViewItem>::CreateInstance(&pVTvwItemObj);
				pVTvwItemObj->AddRef();
				pVTvwItemObj->SetOwner(this);
				if(pTvwItemObj) {
					pTvwItemObj->SaveState(pVTvwItemObj);
				}

				pVTvwItemObj->QueryInterface(IID_IVirtualTreeViewItem, reinterpret_cast<LPVOID*>(&pVTvwItemItf));
				pVTvwItemObj->Release();
			}

			if(!(properties.disabledEvents & deTreeMouseEvents)) {
				if(reinterpret_cast<HTREEITEM>(lParam) == hItemUnderMouse) {
					// we're removing the item below the mouse cursor
					DWORD position = GetMessagePos();
					POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
					ScreenToClient(&mousePosition);
					SHORT button = 0;
					SHORT shift = 0;
					WPARAM2BUTTONANDSHIFT(-1, button, shift);
					UINT hitTestDetails = 0;
					HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
					Raise_ItemMouseLeave(pTvwItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					hItemUnderMouse = NULL;
				}
			}

			// finally pass the message to the treeview
			lr = DefWindowProc(message, wParam, lParam);
			if(lr) {
				if(fireEvents) {
					if(!(properties.disabledEvents & deItemDeletionEvents)) {
						Raise_RemovedItem(pVTvwItemItf);
					}
				}
			}

			if(!(properties.disabledEvents & deTreeMouseEvents)) {
				if(lr) {
					// maybe we have a new item below the mouse cursor now
					DWORD position = GetMessagePos();
					POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
					ScreenToClient(&mousePosition);
					UINT hitTestDetails = 0;
					HTREEITEM hNewItemUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
					if(hNewItemUnderMouse != hItemUnderMouse) {
						SHORT button = 0;
						SHORT shift = 0;
						WPARAM2BUTTONANDSHIFT(-1, button, shift);
						if(hItemUnderMouse) {
							pTvwItemItf = ClassFactory::InitTreeItem(hItemUnderMouse, this);
							Raise_ItemMouseLeave(pTvwItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
						}
						hItemUnderMouse = hNewItemUnderMouse;
						if(hItemUnderMouse) {
							pTvwItemItf = ClassFactory::InitTreeItem(hItemUnderMouse, this);
							Raise_ItemMouseEnter(pTvwItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
						}
					}
				}
			}
		}
	} else {
		// invalid item handle
		lr = DefWindowProc(message, wParam, lParam);
	}

	return lr;
}

LRESULT ExplorerTreeView::OnEditLabel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = S_OK;

	labelEditStatus.watchingForEdit = TRUE;
	lr = DefWindowProc(message, wParam, lParam);
	labelEditStatus.watchingForEdit = FALSE;

	return lr;
}

LRESULT ExplorerTreeView::OnExpandItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = FALSE;

	// TVN_ITEMEXPAND* notifications are not sent if the item has the TVIS_EXPANDEDONCE state
	TVITEMEX item = {0};
	item.hItem = reinterpret_cast<HTREEITEM>(lParam);
	item.mask = TVIF_HANDLE | TVIF_STATE;
	SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item));
	BOOL raiseEvents = ((item.state & TVIS_EXPANDEDONCE) == TVIS_EXPANDEDONCE);

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(reinterpret_cast<HTREEITEM>(lParam), this);
	VARIANT_BOOL cancel = VARIANT_FALSE;

	if((wParam & TVE_TOGGLE) == TVE_TOGGLE) {
		if(raiseEvents) {
			// Are we actually expanding or collapsing?
			if(item.state & TVIS_EXPANDED) {
				// we're currently expanded, so we'll get collapsed
				Raise_CollapsingItem(pTvwItem, VARIANT_FALSE, &cancel);
				if(cancel == VARIANT_FALSE) {
					lr = DefWindowProc(message, wParam, lParam);
					Raise_CollapsedItem(pTvwItem, VARIANT_FALSE);
				}
			} else {
				// we're currently collapsed or partially expanded, so we'll get expanded
				Raise_ExpandingItem(pTvwItem, VARIANT_FALSE, &cancel);
				if(cancel == VARIANT_FALSE) {
					lr = DefWindowProc(message, wParam, lParam);
					Raise_ExpandedItem(pTvwItem, VARIANT_FALSE);
				}
			}
		} else {
			// simply forward the message
			lr = DefWindowProc(message, wParam, lParam);
		}
	} else if(wParam & TVE_COLLAPSE) {
		if(raiseEvents) {
			Raise_CollapsingItem(pTvwItem, BOOL2VARIANTBOOL(wParam & TVE_COLLAPSERESET), &cancel);
		}

		if(cancel == VARIANT_FALSE) {
			lr = DefWindowProc(message, wParam, lParam);
			if(raiseEvents) {
				Raise_CollapsedItem(pTvwItem, BOOL2VARIANTBOOL(wParam & TVE_COLLAPSERESET));
			}
		}

	} else if(wParam & TVE_EXPAND) {
		if(raiseEvents) {
			Raise_ExpandingItem(pTvwItem, BOOL2VARIANTBOOL(wParam & TVE_EXPANDPARTIAL), &cancel);
		}

		if(cancel == VARIANT_FALSE) {
			lr = DefWindowProc(message, wParam, lParam);
			if(raiseEvents) {
				Raise_ExpandedItem(pTvwItem, BOOL2VARIANTBOOL(wParam & TVE_EXPANDPARTIAL));
			}
		}
	}

	return lr;
}

LRESULT ExplorerTreeView::OnGetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	BOOL overwriteLParam = FALSE;

	LPTVITEMEX pItemData = reinterpret_cast<LPTVITEMEX>(lParam);
	#ifdef DEBUG
		if(message == TVM_GETITEMA) {
			ATLASSERT_POINTER(pItemData, TVITEMEXA);
		} else {
			ATLASSERT_POINTER(pItemData, TVITEMEXW);
		}
	#endif
	if(!pItemData) {
		return DefWindowProc(message, wParam, lParam);
	}

	if(pItemData->mask & TVIF_NOINTERCEPTION) {
		// SysTreeView32 shouldn't see this flag
		pItemData->mask &= ~TVIF_NOINTERCEPTION;
	} else {
		if(pItemData->mask == TVIF_PARAM) {
			#ifdef USE_STL
				// simply look up in the 'itemParams' hash map
				std::unordered_map<HTREEITEM, ItemData>::iterator iter = itemParams.find(pItemData->hItem);
				ATLASSERT(iter != itemParams.end());
				if(iter != itemParams.end()) {
					pItemData->lParam = iter->second.lParam;
					return TRUE;
				} else {
					// item not found
					return FALSE;
				}
			#else
				// simply look up in the 'itemParams' map
				CAtlMap<HTREEITEM, ItemData>::CPair* pEntry = itemParams.Lookup(pItemData->hItem);
				ATLASSERT(pEntry != NULL);
				if(pEntry) {
					pItemData->lParam = pEntry->m_value.lParam;
					return TRUE;
				} else {
					// item not found
					return FALSE;
				}
			#endif
		} else if(pItemData->mask & TVIF_PARAM) {
			overwriteLParam = TRUE;
		}
	}

	// forward the message
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(overwriteLParam) {
		#ifdef USE_STL
			std::unordered_map<HTREEITEM, ItemData>::iterator iter = itemParams.find(pItemData->hItem);
			ATLASSERT(iter != itemParams.end());
			if(iter != itemParams.end()) {
				pItemData->lParam = iter->second.lParam;
			}
		#else
			CAtlMap<HTREEITEM, ItemData>::CPair* pEntry = itemParams.Lookup(pItemData->hItem);
			ATLASSERT(pEntry != NULL);
			if(pEntry) {
				pItemData->lParam = pEntry->m_value.lParam;
			}
		#endif
	}
	return lr;
}

LRESULT ExplorerTreeView::OnInsertItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	HTREEITEM hInsertedItem = NULL;

	if(!(properties.disabledEvents & deItemInsertionEvents)) {
		CComObject<VirtualTreeViewItem>* pVTvwItemObj = NULL;
		CComPtr<IVirtualTreeViewItem> pVTvwItemItf = NULL;
		VARIANT_BOOL cancel = VARIANT_FALSE;
		if(flags.silentItemInsertions == 0) {
			CComObject<VirtualTreeViewItem>::CreateInstance(&pVTvwItemObj);
			pVTvwItemObj->AddRef();
			pVTvwItemObj->SetOwner(this);

			#ifdef UNICODE
				BOOL mustConvert = (message == TVM_INSERTITEMA);
			#else
				BOOL mustConvert = (message == TVM_INSERTITEMW);
			#endif
			if(mustConvert) {
				#ifdef UNICODE
					LPTVINSERTSTRUCTA pItemData = reinterpret_cast<LPTVINSERTSTRUCTA>(lParam);
					TVINSERTSTRUCTW convertedItemData = {0};
					LPSTR p = NULL;
					if(pItemData->itemex.pszText != LPSTR_TEXTCALLBACKA) {
						p = pItemData->itemex.pszText;
					}
					CA2W converter(p);
					if(pItemData->itemex.pszText == LPSTR_TEXTCALLBACKA) {
						convertedItemData.itemex.pszText = LPSTR_TEXTCALLBACKW;
					} else {
						convertedItemData.itemex.pszText = converter;
					}
				#else
					LPTVINSERTSTRUCTW pItemData = reinterpret_cast<LPTVINSERTSTRUCTW>(lParam);
					TVINSERTSTRUCTA convertedItemData = {0};
					LPWSTR p = NULL;
					if(pItemData->itemex.pszText != LPSTR_TEXTCALLBACKW) {
						p = pItemData->itemex.pszText;
					}
					CW2A converter(p);
					if(pItemData->itemex.pszText == LPSTR_TEXTCALLBACKW) {
						convertedItemData.itemex.pszText = LPSTR_TEXTCALLBACKA;
					} else {
						convertedItemData.itemex.pszText = converter;
					}
				#endif
				convertedItemData.itemex.cchTextMax = pItemData->itemex.cchTextMax;
				convertedItemData.hParent = pItemData->hParent;
				convertedItemData.hInsertAfter = pItemData->hInsertAfter;
				convertedItemData.itemex.mask = pItemData->itemex.mask;
				convertedItemData.itemex.hItem = pItemData->itemex.hItem;
				convertedItemData.itemex.state = pItemData->itemex.state;
				convertedItemData.itemex.stateMask = pItemData->itemex.stateMask;
				convertedItemData.itemex.iImage = pItemData->itemex.iImage;
				convertedItemData.itemex.iSelectedImage = pItemData->itemex.iSelectedImage;
				convertedItemData.itemex.cChildren = pItemData->itemex.cChildren;
				convertedItemData.itemex.lParam = pItemData->itemex.lParam;
				convertedItemData.itemex.iIntegral = pItemData->itemex.iIntegral;
				// the TVINSERTSTRUCT struct may end here, so be more careful with the remaining members
				if(convertedItemData.itemex.mask & TVIF_STATEEX) {
					convertedItemData.itemex.uStateEx = pItemData->itemex.uStateEx;
					convertedItemData.itemex.hwnd = pItemData->itemex.hwnd;
				}
				if(convertedItemData.itemex.mask & TVIF_EXPANDEDIMAGE) {
					convertedItemData.itemex.iExpandedImage = pItemData->itemex.iExpandedImage;
				}
				pVTvwItemObj->LoadState(&convertedItemData);
			} else {
				pVTvwItemObj->LoadState(reinterpret_cast<LPTVINSERTSTRUCT>(lParam));
			}

			pVTvwItemObj->QueryInterface(IID_IVirtualTreeViewItem, reinterpret_cast<LPVOID*>(&pVTvwItemItf));
			pVTvwItemObj->Release();

			Raise_InsertingItem(pVTvwItemItf, &cancel);
			pVTvwItemObj = NULL;
		}
		if(cancel == VARIANT_FALSE) {
			// finally pass the message to the treeview
			LPTVINSERTSTRUCT pDetails = reinterpret_cast<LPTVINSERTSTRUCT>(lParam);
			ItemData itemData = {0};
			pDetails->itemex.mask &= ~TVIF_NOINTERCEPTION;
			if(flags.silentItemInsertions == 0) {
				itemData.lParam = pDetails->itemex.lParam;
				pDetails->itemex.lParam = GetNewItemID();
				pDetails->itemex.mask |= TVIF_PARAM;
			}

			#ifdef INCLUDESHELLBROWSERINTERFACE
				BOOL isShellItem = ((pDetails->itemex.mask & TVIF_ISASHELLITEM) == TVIF_ISASHELLITEM);
				pDetails->itemex.mask &= ~TVIF_ISASHELLITEM;
				if(shellBrowserInterface.pInternalMessageListener) {
					shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_INSERTINGITEM, pDetails->itemex.lParam, 0);
				}
			#endif

			flags.noCustomDraw = TRUE;
			hInsertedItem = reinterpret_cast<HTREEITEM>(DefWindowProc(message, wParam, lParam));
			flags.noCustomDraw = FALSE;
			if(hInsertedItem) {
				if(flags.silentItemInsertions == 0) {
					/* NOTE: Activating silent insertions implies that the 'itemHandles' and 'itemParams' hash
					         maps must be updated explicitly. */
					itemHandles[static_cast<LONG>(pDetails->itemex.lParam)] = hInsertedItem;
					itemParams[hInsertedItem] = itemData;

					#ifdef INCLUDESHELLBROWSERINTERFACE
						if(shellBrowserInterface.pInternalMessageListener) {
							if(isShellItem) {
								pDetails->itemex.mask |= TVIF_ISASHELLITEM;
							}
							shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_INSERTEDITEM, reinterpret_cast<WPARAM>(hInsertedItem), lParam);
						}
					#endif

					CComPtr<ITreeViewItem> pTvwItemItf = ClassFactory::InitTreeItem(hInsertedItem, this);
					Raise_InsertedItem(pTvwItemItf);
				}
			}
		}
	} else {
		// finally pass the message to the treeview
		LPTVINSERTSTRUCT pDetails = reinterpret_cast<LPTVINSERTSTRUCT>(lParam);
		ItemData itemData = {0};
		pDetails->itemex.mask &= ~TVIF_NOINTERCEPTION;
		if(flags.silentItemInsertions == 0) {
			itemData.lParam = pDetails->itemex.lParam;
			pDetails->itemex.lParam = GetNewItemID();
			pDetails->itemex.mask |= TVIF_PARAM;
		}

		#ifdef INCLUDESHELLBROWSERINTERFACE
			BOOL isShellItem = ((pDetails->itemex.mask & TVIF_ISASHELLITEM) == TVIF_ISASHELLITEM);
			pDetails->itemex.mask &= ~TVIF_ISASHELLITEM;
			if(shellBrowserInterface.pInternalMessageListener) {
				shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_INSERTINGITEM, pDetails->itemex.lParam, 0);
			}
		#endif

		flags.noCustomDraw = TRUE;
		hInsertedItem = reinterpret_cast<HTREEITEM>(DefWindowProc(message, wParam, lParam));
		flags.noCustomDraw = FALSE;
		if(hInsertedItem) {
			if(flags.silentItemInsertions == 0) {
				/* NOTE: Activating silent insertions implies that the 'itemHandles' and 'itemParams' hash
				         maps must be updated explicitly. */
				itemHandles[static_cast<LONG>(pDetails->itemex.lParam)] = hInsertedItem;
				itemParams[hInsertedItem] = itemData;

				#ifdef INCLUDESHELLBROWSERINTERFACE
					if(shellBrowserInterface.pInternalMessageListener) {
						if(isShellItem) {
							pDetails->itemex.mask |= TVIF_ISASHELLITEM;
						}
						shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_INSERTEDITEM, reinterpret_cast<WPARAM>(hInsertedItem), lParam);
					}
				#endif
			}
		}
	}

	if(!(properties.disabledEvents & deTreeMouseEvents)) {
		if(hInsertedItem) {
			// maybe we have a new item below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			UINT hitTestDetails = 0;
			HTREEITEM hNewItemUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
			if(hNewItemUnderMouse != hItemUnderMouse) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(hItemUnderMouse) {
					CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItemUnderMouse, this);
					Raise_ItemMouseLeave(pTvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				hItemUnderMouse = hNewItemUnderMouse;
				if(hItemUnderMouse) {
					CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItemUnderMouse, this);
					Raise_ItemMouseEnter(pTvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}
			}
		}
	}

	return reinterpret_cast<LRESULT>(hInsertedItem);
}

LRESULT ExplorerTreeView::OnSelectItem(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(LOWORD(wParam) == TVGN_CARET) {
		labelEditStatus.watchingForEdit = FALSE;
	}

	wasHandled = FALSE;
	return TRUE;
}

LRESULT ExplorerTreeView::OnSetAutoScrollInfo(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		properties.autoHScrollPixelsPerSecond = static_cast<LONG>(wParam);
		properties.autoHScrollRedrawInterval = static_cast<LONG>(lParam);
	}
	return lr;
}

LRESULT ExplorerTreeView::OnSetBorder(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	/* We must store the new settings before the call to DefWindowProc, because we may need them to
	   handle the WM_PAINT message this message probably will cause. */
	if(wParam & TVSBF_XBORDER) {
		cachedSettings.itemXBorder = LOWORD(lParam);
	}
	if(wParam & TVSBF_YBORDER) {
		cachedSettings.itemYBorder = HIWORD(lParam);
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	// TODO: Isn't there a better way for correcting the values if an error occured?
	DWORD itemBorders = static_cast<DWORD>(SendMessage(TVM_GETBORDER, 0, 0));
	cachedSettings.itemXBorder = LOWORD(itemBorders);
	cachedSettings.itemYBorder = HIWORD(itemBorders);

	return lr;
}

LRESULT ExplorerTreeView::OnSetImageList(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	/* We must store the new settings before the call to DefWindowProc, because we may need them to
	   handle the WM_PAINT message this message probably will cause. */
	switch(wParam) {
		case TVSIL_NORMAL:
			cachedSettings.hImageList = reinterpret_cast<HIMAGELIST>(lParam);
			if(!ImageList_GetIconSize(cachedSettings.hImageList, &cachedSettings.iconWidth, &cachedSettings.iconHeight)) {
				cachedSettings.iconHeight = 0;
				cachedSettings.iconWidth = 0;
			}
			break;
		case TVSIL_STATE:
			cachedSettings.hStateImageList = reinterpret_cast<HIMAGELIST>(lParam);
			break;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	// TODO: Isn't there a better way for correcting the values if an error occured?
	switch(wParam) {
		case TVSIL_NORMAL:
			cachedSettings.hImageList = reinterpret_cast<HIMAGELIST>(SendMessage(TVM_GETIMAGELIST, TVSIL_NORMAL, 0));
			if(!ImageList_GetIconSize(cachedSettings.hImageList, &cachedSettings.iconWidth, &cachedSettings.iconHeight)) {
				cachedSettings.iconHeight = 0;
				cachedSettings.iconWidth = 0;
			}
			break;
		case TVSIL_STATE:
			cachedSettings.hStateImageList = reinterpret_cast<HIMAGELIST>(SendMessage(TVM_GETIMAGELIST, TVSIL_STATE, 0));
			break;
	}

	return lr;
}

LRESULT ExplorerTreeView::OnSetInsertMark(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	insertMark.ProcessSetInsertMark(wParam, lParam);
	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerTreeView::OnSetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPTVITEMEX pItemData = reinterpret_cast<LPTVITEMEX>(lParam);
	#ifdef DEBUG
		if(message == TVM_SETITEMA) {
			ATLASSERT_POINTER(pItemData, TVITEMEXA);
		} else {
			ATLASSERT_POINTER(pItemData, TVITEMEXW);
		}
	#endif
	if(!pItemData) {
		return DefWindowProc(message, wParam, lParam);
	}

	if(pItemData->mask & TVIF_NOINTERCEPTION) {
		// SysTreeView32 shouldn't see this flag
		pItemData->mask &= ~TVIF_NOINTERCEPTION;
	} else {
		if(pItemData->mask == TVIF_PARAM) {
			#ifdef USE_STL
				// simply update the 'itemParams' hash map
				std::unordered_map<HTREEITEM, ItemData>::iterator iter = itemParams.find(pItemData->hItem);
				ATLASSERT(iter != itemParams.end());
				if(iter != itemParams.end()) {
					iter->second.lParam = pItemData->lParam;
					return TRUE;
				} else {
					// item not found
					return FALSE;
				}
			#else
				// simply update the 'itemParams' map
				CAtlMap<HTREEITEM, ItemData>::CPair* pEntry = itemParams.Lookup(pItemData->hItem);
				ATLASSERT(pEntry != NULL);
				if(pEntry) {
					itemParams[pItemData->hItem].lParam = pItemData->lParam;
					return TRUE;
				} else {
					// item not found
					return FALSE;
				}
			#endif
		} else if(pItemData->mask & TVIF_PARAM) {
			#ifdef USE_STL
				std::unordered_map<HTREEITEM, ItemData>::iterator iter = itemParams.find(pItemData->hItem);
				ATLASSERT(iter != itemParams.end());
				if(iter != itemParams.end()) {
					iter->second.lParam = pItemData->lParam;
					pItemData->mask &= ~TVIF_PARAM;
				}
			#else
				CAtlMap<HTREEITEM, ItemData>::CPair* pEntry = itemParams.Lookup(pItemData->hItem);
				ATLASSERT(pEntry != NULL);
				if(pEntry) {
					itemParams[pItemData->hItem].lParam = pItemData->lParam;
					pItemData->mask &= ~TVIF_PARAM;
				}
			#endif
		}
	}

	if((stateImageChangeFlags.silentStateImageChanges == 0) && IsComctl32Version610OrNewer()) {
		if(pItemData->mask & TVIF_STATE) {
			if(pItemData->stateMask & TVIS_STATEIMAGEMASK) {
				stateImageChangeFlags.hItemCurrentInternalStateImageChange = pItemData->hItem;
				LRESULT lr = DefWindowProc(message, wParam, lParam);
				stateImageChangeFlags.hItemCurrentInternalStateImageChange = NULL;
				return lr;
			}
		}
	}

	return DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerTreeView::OnSetItemHeight(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	/* We must store the new settings before the call to DefWindowProc, because we may need them to
	   handle the WM_PAINT message this message probably will cause. */
	cachedSettings.itemHeight = static_cast<SHORT>(wParam);

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	// TODO: Isn't there a better way for correcting the values if an error occured?
	cachedSettings.itemHeight = static_cast<SHORT>(SendMessage(TVM_GETITEMHEIGHT, 0, 0));

	return lr;
}


LRESULT ExplorerTreeView::OnEditChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_EditKeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT ExplorerTreeView::OnEditContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		containedEdit.ScreenToClient(&mousePosition);
	}
	VARIANT_BOOL showDefaultMenu = VARIANT_TRUE;
	Raise_EditContextMenu(button, shift, mousePosition.x, mousePosition.y, &showDefaultMenu);
	if(showDefaultMenu != VARIANT_FALSE) {
		wasHandled = FALSE;
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_DestroyedEditControlWindow(HandleToLong(properties.hWndEdit));
	return 0;
}

LRESULT ExplorerTreeView::OnEditKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deEditKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_EditKeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return containedEdit.DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerTreeView::OnEditKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deEditKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_EditKeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return containedEdit.DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerTreeView::OnEditKillFocus(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_EditLostFocus();
	return 0;
}

LRESULT ExplorerTreeView::OnEditLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_EditDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_EditMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_EditMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_EditMDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_EditMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_EditMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_EditMouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		Raise_EditMouseLeave(button, shift, mouseStatus_Edit.lastPosition.x, mouseStatus_Edit.lastPosition.y);
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_EditMouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(message == WM_MOUSEHWHEEL) {
			// wParam and lParam seem to be 0
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			GetCursorPos(&mousePosition);
		} else {
			WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		}
		containedEdit.ScreenToClient(&mousePosition);
		Raise_EditMouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_EditRDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_EditMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_EditMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	LRESULT lr = CComControl<ExplorerTreeView>::OnSetFocus(message, wParam, lParam, wasHandled);

	if(!IsInDesignMode()) {
		// now that we've the focus, we should configure the IME
		IMEModeConstants ime = GetCurrentIMEContextMode(containedEdit.m_hWnd);
		if(ime != imeInherit) {
			ime = GetEffectiveIMEMode_Edit();
			if(ime != imeNoControl) {
				SetCurrentIMEContextMode(containedEdit.m_hWnd, ime);
			}
		}
	}

	Raise_EditGotFocus();
	return lr;
}

LRESULT ExplorerTreeView::OnEditXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_EditXDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_EditMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerTreeView::OnEditXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_EditMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

#ifdef INCLUDESHELLBROWSERINTERFACE
	HRESULT ExplorerTreeView::OnInternalAddItem(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		LPEXTVMADDITEMDATA pDetails = reinterpret_cast<LPEXTVMADDITEMDATA>(lParam);
		ATLASSERT_POINTER(pDetails, EXTVMADDITEMDATA);
		if(!pDetails) {
			return E_POINTER;
		}

		if(wParam) {
			TVINSERTSTRUCT insertionData = {0};
			insertionData.hParent = static_cast<HTREEITEM>(LongToHandle(pDetails->parentItem));
			insertionData.hInsertAfter = static_cast<HTREEITEM>(LongToHandle(pDetails->insertAfter));
			insertionData.itemex.mask = TVIF_CHILDREN | TVIF_INTEGRAL | TVIF_ISASHELLITEM | TVIF_PARAM | TVIF_STATE | TVIF_STATEEX | TVIF_TEXT;
			insertionData.itemex.pszText = LPSTR_TEXTCALLBACK;
			if(pDetails->iconIndex != -2) {
				insertionData.itemex.iImage = pDetails->iconIndex;
				insertionData.itemex.mask |= TVIF_IMAGE;
			}
			if(pDetails->selectedIconIndex != -2) {
				insertionData.itemex.iSelectedImage = pDetails->selectedIconIndex;
				insertionData.itemex.mask |= TVIF_SELECTEDIMAGE;
			}
			if(pDetails->expandedIconIndex != -2) {
				insertionData.itemex.iExpandedImage = pDetails->expandedIconIndex;
				insertionData.itemex.mask |= TVIF_EXPANDEDIMAGE;
			}
			if(pDetails->overlayIndex > 0) {
				insertionData.itemex.state |= INDEXTOOVERLAYMASK(pDetails->overlayIndex);
				insertionData.itemex.stateMask |= TVIS_OVERLAYMASK;
			}
			insertionData.itemex.cChildren = pDetails->hasExpando;
			insertionData.itemex.iIntegral = pDetails->heightIncrement;
			insertionData.itemex.lParam = pDetails->itemData;
			if(pDetails->isGhosted != VARIANT_FALSE) {
				insertionData.itemex.state |= TVIS_CUT;
				insertionData.itemex.stateMask |= TVIS_CUT;
			}
			if(pDetails->isVirtual != VARIANT_FALSE) {
				insertionData.itemex.uStateEx |= TVIS_EX_FLAT;
			}
			pDetails->hInsertedItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData)));
		} else {
			CComObject<TreeViewItems>* pTvwItemsObj = NULL;
			CComObject<TreeViewItems>::CreateInstance(&pTvwItemsObj);
			pTvwItemsObj->AddRef();
			pTvwItemsObj->SetOwner(this);
			pDetails->hInsertedItem = pTvwItemsObj->Add(pDetails->pItemText, static_cast<HTREEITEM>(LongToHandle(pDetails->parentItem)), static_cast<HTREEITEM>(LongToHandle(pDetails->insertAfter)), pDetails->hasExpando, pDetails->iconIndex, pDetails->selectedIconIndex, pDetails->expandedIconIndex, pDetails->overlayIndex, pDetails->itemData, pDetails->isGhosted, pDetails->isVirtual, pDetails->heightIncrement, TRUE);
			pTvwItemsObj->Release();
		}
		return (pDetails->hInsertedItem ? S_OK : E_FAIL);
	}

	HRESULT ExplorerTreeView::OnInternalItemHandleToID(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		PLONG pItemID = reinterpret_cast<PLONG>(lParam);
		ATLASSERT_POINTER(pItemID, LONG);
		if(!pItemID) {
			return E_POINTER;
		}

		*pItemID = ItemHandleToID(reinterpret_cast<HTREEITEM>(wParam));
		return (*pItemID != 0 ? S_OK : E_FAIL);
	}

	HRESULT ExplorerTreeView::OnInternalIDToItemHandle(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		HTREEITEM* pItemHandle = reinterpret_cast<HTREEITEM*>(lParam);
		ATLASSERT_POINTER(pItemHandle, HTREEITEM);
		if(!pItemHandle) {
			return E_POINTER;
		}

		*pItemHandle = IDToItemHandle(static_cast<LONG>(wParam));
		return (*pItemHandle ? S_OK : E_FAIL);
	}

	HRESULT ExplorerTreeView::OnInternalGetItemByHandle(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		HTREEITEM hItem = reinterpret_cast<HTREEITEM>(wParam);
		if(hItem) {
			ClassFactory::InitTreeItem(hItem, this, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(lParam));
			return S_OK;
		}
		return E_INVALIDARG;
	}

	HRESULT ExplorerTreeView::OnInternalCreateItemContainer(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		LPEXTVMCREATEITEMCONTAINERDATA pDetails = reinterpret_cast<LPEXTVMCREATEITEMCONTAINERDATA>(lParam);
		ATLASSERT_POINTER(pDetails, EXTVMCREATEITEMCONTAINERDATA);
		if(!pDetails) {
			return E_POINTER;
		}
		ATLASSERT(pDetails->pItems);
		ATLASSERT(!pDetails->pContainer);

		CComObject<TreeViewItemContainer>* pTvwItemContainerObj = NULL;
		CComObject<TreeViewItemContainer>::CreateInstance(&pTvwItemContainerObj);
		pTvwItemContainerObj->AddRef();

		// clone all settings
		pTvwItemContainerObj->SetOwner(this);
		PLONG pItems = new LONG[pDetails->numberOfItems];
		for(UINT i = 0; i < pDetails->numberOfItems; i++) {
			pItems[i] = ItemHandleToID(pDetails->pItems[i]);
		}
		// fast-add the items
		static_cast<IItemContainer*>(pTvwItemContainerObj)->AddItems(pDetails->numberOfItems, pItems);
		delete[] pItems;

		HRESULT hr = pTvwItemContainerObj->QueryInterface(IID_PPV_ARGS(&pDetails->pContainer));
		RegisterItemContainer(static_cast<IItemContainer*>(pTvwItemContainerObj));
		pTvwItemContainerObj->Release();
		return hr;
	}

	HRESULT ExplorerTreeView::OnInternalGetItemHandlesFromVariant(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		LPEXTVMGETITEMHANDLESDATA pDetails = reinterpret_cast<LPEXTVMGETITEMHANDLESDATA>(lParam);
		ATLASSERT_POINTER(pDetails, EXTVMGETITEMHANDLESDATA);
		if(!pDetails) {
			return E_POINTER;
		}
		ATLASSERT(pDetails->pVariant);
		ATLASSERT(!pDetails->pItemHandles);
		if(!pDetails->pVariant || pDetails->pItemHandles) {
			return E_INVALIDARG;
		}

		if(pDetails->pVariant->vt != VT_DISPATCH) {
			return DISP_E_TYPEMISMATCH;
		}

		// invalid arg - VB runtime error 380
		HRESULT hr = MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		if(pDetails->pVariant->pdispVal) {
			OLE_HANDLE h = NULL;
			CComQIPtr<ITreeViewItem> pTvwItem = pDetails->pVariant->pdispVal;
			if(pTvwItem) {
				// a single TreeViewItem object
				ATLASSUME(shellBrowserInterface.pInternalMessageListener);
				pDetails->pItemHandles = reinterpret_cast<HTREEITEM*>(shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_ALLOCATEMEMORY, sizeof(HTREEITEM), NULL));
				ATLASSERT_ARRAYPOINTER(pDetails->pItemHandles, HTREEITEM, 1);
				pDetails->pItemHandles[0] = NULL;
				hr = pTvwItem->get_Handle(&h);
				pDetails->pItemHandles[0] = static_cast<HTREEITEM>(LongToHandle(h));
				pDetails->numberOfItems = 1;
			} else {
				CComQIPtr<ITreeViewItemContainer> pTvwItemContainer = pDetails->pVariant->pdispVal;
				if(pTvwItemContainer) {
					// a TreeViewItemContainer collection
					CComQIPtr<IEnumVARIANT> pEnumerator = pTvwItemContainer;
					LONG c = 0;
					pTvwItemContainer->Count(&c);
					pDetails->numberOfItems = c;
					if(pDetails->numberOfItems > 0 && pEnumerator) {
						ATLASSUME(shellBrowserInterface.pInternalMessageListener);
						pDetails->pItemHandles = reinterpret_cast<HTREEITEM*>(shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_ALLOCATEMEMORY, pDetails->numberOfItems * sizeof(HTREEITEM), NULL));
						ATLASSERT_ARRAYPOINTER(pDetails->pItemHandles, HTREEITEM, pDetails->numberOfItems);
						VARIANT item;
						VariantInit(&item);
						ULONG dummy = 0;
						for(UINT i = 0; i < pDetails->numberOfItems && pEnumerator->Next(1, &item, &dummy) == S_OK; ++i) {
							pDetails->pItemHandles[i] = NULL;
							if((item.vt == VT_DISPATCH) && item.pdispVal) {
								pTvwItem = item.pdispVal;
								if(pTvwItem) {
									h = NULL;
									pTvwItem->get_Handle(&h);
									pDetails->pItemHandles[i] = static_cast<HTREEITEM>(LongToHandle(h));
								}
							}
							VariantClear(&item);
						}
						hr = S_OK;
					}
				} else {
					CComQIPtr<ITreeViewItems> pTvwItems = pDetails->pVariant->pdispVal;
					if(pTvwItems) {
						// a TreeViewItems collection
						CComQIPtr<IEnumVARIANT> pEnumerator = pTvwItems;
						LONG c = 0;
						pTvwItems->Count(&c);
						pDetails->numberOfItems = c;
						if(pDetails->numberOfItems > 0 && pEnumerator) {
							ATLASSUME(shellBrowserInterface.pInternalMessageListener);
							pDetails->pItemHandles = reinterpret_cast<HTREEITEM*>(shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_ALLOCATEMEMORY, pDetails->numberOfItems * sizeof(HTREEITEM), NULL));
							ATLASSERT_ARRAYPOINTER(pDetails->pItemHandles, HTREEITEM, pDetails->numberOfItems);
							VARIANT item;
							VariantInit(&item);
							ULONG dummy = 0;
							for(UINT i = 0; i < pDetails->numberOfItems && pEnumerator->Next(1, &item, &dummy) == S_OK; ++i) {
								pDetails->pItemHandles[i] = NULL;
								if((item.vt == VT_DISPATCH) && item.pdispVal) {
									pTvwItem = item.pdispVal;
									if(pTvwItem) {
										h = NULL;
										pTvwItem->get_Handle(&h);
										pDetails->pItemHandles[i] = static_cast<HTREEITEM>(LongToHandle(h));
									}
								}
								VariantClear(&item);
							}
							hr = S_OK;
						}
					}
				}
			}
		}
		return hr;
	}

	HRESULT ExplorerTreeView::OnInternalSetItemIconIndex(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		TVITEMEX item = {0};
		item.hItem = reinterpret_cast<HTREEITEM>(wParam);
		item.iImage = static_cast<int>(lParam);
		item.mask = TVIF_HANDLE | TVIF_IMAGE;
		/* NOTE: Don't use SendMessage here, because this leads to the message being processed by ShellListView
		 *       which won't know that item.iImage is an ID.
		 */
		if(DefWindowProc(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
		return DISP_E_BADINDEX;
	}

	HRESULT ExplorerTreeView::OnInternalSetItemSelectedIconIndex(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		TVITEMEX item = {0};
		item.hItem = reinterpret_cast<HTREEITEM>(wParam);
		item.iSelectedImage = static_cast<int>(lParam);
		item.mask = TVIF_HANDLE | TVIF_SELECTEDIMAGE;
		/* NOTE: Don't use SendMessage here, because this leads to the message being processed by ShellListView
		 *       which won't know that item.iImage is an ID.
		 */
		if(DefWindowProc(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
		return DISP_E_BADINDEX;
	}

	HRESULT ExplorerTreeView::OnInternalSetItemExpandedIconIndex(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		TVITEMEX item = {0};
		item.hItem = reinterpret_cast<HTREEITEM>(wParam);
		item.iExpandedImage = static_cast<int>(lParam);
		item.mask = TVIF_HANDLE | TVIF_EXPANDEDIMAGE;
		/* NOTE: Don't use SendMessage here, because this leads to the message being processed by ShellListView
		 *       which won't know that item.iImage is an ID.
		 */
		if(DefWindowProc(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
		return DISP_E_BADINDEX;
	}

	HRESULT ExplorerTreeView::OnInternalHitTest(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		LPEXTVMHITTESTDATA pDetails = reinterpret_cast<LPEXTVMHITTESTDATA>(lParam);
		ATLASSERT_POINTER(pDetails, EXTVMHITTESTDATA);
		if(!pDetails) {
			return E_POINTER;
		}

		pDetails->hHitItem = HitTest(pDetails->pointToTest.x, pDetails->pointToTest.y, &pDetails->hitTestDetails);

		return S_OK;
	}

	HRESULT ExplorerTreeView::OnInternalSortItems(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		if(lParam) {
			CComPtr<ITreeViewItem> pItem = ClassFactory::InitTreeItem(reinterpret_cast<HTREEITEM>(lParam), this);
			return pItem->SortSubItems(sobShell, sobText, sobNone, sobNone, sobNone, BOOL2VARIANTBOOL(wParam));
		} else {
			return SortItems(sobShell, sobText, sobNone, sobNone, sobNone, BOOL2VARIANTBOOL(wParam));
		}
	}

	HRESULT ExplorerTreeView::OnInternalSetDropHilitedItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		return putref_DropHilitedItem(reinterpret_cast<ITreeViewItem*>(lParam));
	}

	HRESULT ExplorerTreeView::OnInternalFireDragDropEvent(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		LPSHTVMDRAGDROPEVENTDATA pDetails = reinterpret_cast<LPSHTVMDRAGDROPEVENTDATA>(lParam);
		ATLASSERT(pDetails);
		if(!pDetails) {
			return E_POINTER;
		}
		if(pDetails->structSize < sizeof(SHTVMDRAGDROPEVENTDATA)) {
			return E_INVALIDARG;
		}

		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(pDetails->keyState, button, shift);
		POINT cursorPosition = {pDetails->cursorPosition.x, pDetails->cursorPosition.y};
		ScreenToClient(&cursorPosition);
		if(pDetails->pDataObject) {
			switch(wParam) {
				case OLEDRAGEVENT_DRAGENTER:
					return Fire_OLEDragEnter(reinterpret_cast<IOLEDataObject*>(pDetails->pDataObject), reinterpret_cast<OLEDropEffectConstants*>(&pDetails->effect), reinterpret_cast<ITreeViewItem**>(&pDetails->pDropTarget), button, shift, cursorPosition.x, cursorPosition.y, pDetails->yToItemTop, static_cast<HitTestConstants>(pDetails->hitTestDetails), &pDetails->autoExpandItem, &pDetails->autoHScrollVelocity, &pDetails->autoVScrollVelocity);
					break;
				case OLEDRAGEVENT_DRAGOVER:
					return Fire_OLEDragMouseMove(reinterpret_cast<IOLEDataObject*>(pDetails->pDataObject), reinterpret_cast<OLEDropEffectConstants*>(&pDetails->effect), reinterpret_cast<ITreeViewItem**>(&pDetails->pDropTarget), button, shift, cursorPosition.x, cursorPosition.y, pDetails->yToItemTop, static_cast<HitTestConstants>(pDetails->hitTestDetails), &pDetails->autoExpandItem, &pDetails->autoHScrollVelocity, &pDetails->autoVScrollVelocity);
					break;
				case OLEDRAGEVENT_DROP:
					return Fire_OLEDragDrop(reinterpret_cast<IOLEDataObject*>(pDetails->pDataObject), reinterpret_cast<OLEDropEffectConstants*>(&pDetails->effect), reinterpret_cast<ITreeViewItem**>(&pDetails->pDropTarget), button, shift, cursorPosition.x, cursorPosition.y, pDetails->yToItemTop, static_cast<HitTestConstants>(pDetails->hitTestDetails));
					break;
				case OLEDRAGEVENT_DRAGLEAVE:
					return Fire_OLEDragLeave(reinterpret_cast<IOLEDataObject*>(pDetails->pDataObject), reinterpret_cast<ITreeViewItem*>(pDetails->pDropTarget), button, shift, cursorPosition.x, cursorPosition.y, pDetails->yToItemTop, static_cast<HitTestConstants>(pDetails->hitTestDetails));
					break;
			}
		}
		return E_INVALIDARG;
	}

	HRESULT ExplorerTreeView::OnInternalOLEDrag(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		LPEXTVMOLEDRAGDATA pDetails = reinterpret_cast<LPEXTVMOLEDRAGDATA>(lParam);
		ATLASSERT_POINTER(pDetails, EXTVMOLEDRAGDATA);
		if(!pDetails) {
			return E_POINTER;
		}
		ATLASSERT(pDetails->pDataObject);
		ATLASSERT(pDetails->pDraggedItems);

		if(pDetails->useSHDoDragDrop) {
			if(!properties.supportOLEDragImages) {
				// SHDoDragDrop always is with drag images
				return E_FAIL;
			}
			pDetails->performedEffects = odeNone;

			CComPtr<IOLEDataObject> pOLEDataObject = NULL;
			CComObject<TargetOLEDataObject>* pOLEDataObjectObj = NULL;
			CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
			pOLEDataObjectObj->AddRef();
			pOLEDataObjectObj->Attach(pDetails->pDataObject);
			pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
			pOLEDataObjectObj->Release();
			pDetails->pDataObject->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&dragDropStatus.pSourceDataObject));

			if(pDetails->pDraggedItems) {
				CComQIPtr<ITreeViewItemContainer> pContainer = pDetails->pDraggedItems;
				if(pContainer) {
					pContainer->Clone(&dragDropStatus.pDraggedItems);
				}
			}
			CComPtr<IDropSource> pDragSource = NULL;
			if(!pDetails->useShellDropSource) {
				QueryInterface(IID_PPV_ARGS(&pDragSource));
			}

			if(IsLeftMouseButtonDown()) {
				dragDropStatus.draggingMouseButton = MK_LBUTTON;
			} else if(IsRightMouseButtonDown()) {
				dragDropStatus.draggingMouseButton = MK_RBUTTON;
			}

			if(pOLEDataObject) {
				Raise_OLEStartDrag(pOLEDataObject);
			}
			HRESULT hr = SHDoDragDrop((RunTimeHelper::IsVista() ? NULL : static_cast<HWND>(*this)), pDetails->pDataObject, pDragSource, pDetails->supportedEffects, &pDetails->performedEffects);
			if((hr == DRAGDROP_S_DROP) && pOLEDataObject) {
				Raise_OLECompleteDrag(pOLEDataObject, static_cast<OLEDropEffectConstants>(pDetails->performedEffects));
			}

			dragDropStatus.Reset();
			return (hr != DRAGDROP_S_DROP && hr != DRAGDROP_S_CANCEL ? hr : S_OK);
		} else {
			pDetails->performedEffects = odeNone;
			CComQIPtr<ITreeViewItemContainer> pDraggedItems = pDetails->pDraggedItems;
			return OLEDrag(reinterpret_cast<LONG*>(&pDetails->pDataObject), static_cast<OLEDropEffectConstants>(pDetails->supportedEffects), static_cast<OLE_HANDLE>(-1), pDraggedItems, -1, reinterpret_cast<OLEDropEffectConstants*>(&pDetails->performedEffects));
		}
	}

	HRESULT ExplorerTreeView::OnInternalGetImageList(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		OLE_HANDLE h;
		HRESULT hr = get_hImageList(static_cast<ImageListConstants>(wParam), &h);
		*reinterpret_cast<HIMAGELIST*>(lParam) = reinterpret_cast<HIMAGELIST>(LongToHandle(h));
		return hr;
	}

	HRESULT ExplorerTreeView::OnInternalSetImageList(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		return put_hImageList(static_cast<ImageListConstants>(wParam), HandleToLong(reinterpret_cast<HIMAGELIST>(lParam)));
	}
#endif


LRESULT ExplorerTreeView::OnClickNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	ScreenToClient(&mousePosition);
	HTREEITEM hPreviousCaretItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0));

	UINT hitTestFlags = 0;
	// retrieve the item whose row the click was within
	HTREEITEM hClickedRow = HitTest(mousePosition.x, mousePosition.y, &hitTestFlags, TRUE);

	SingleExpandConstants singleExpand = seNone;
	get_SingleExpand(&singleExpand);
	if(hitTestFlags & TVHT_ONITEMLABEL) {
		if(GetFocus() == *this) {
			if(hClickedRow == hPreviousCaretItem) {
				labelEditStatus.watchingForEdit = TRUE;

				if(singleExpand != seNone) {
					SetTimer(timers.ID_EDITLABEL, GetDoubleClickTime());
				}
			}
		}
	}

	if(!(properties.disabledEvents & deTreeClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		button |= 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_Click(button, shift, mousePosition.x, mousePosition.y);
	}

	LRESULT lr = FALSE;
	if(singleExpand == seWinXPStyle) {
		if(((hitTestFlags & TVHT_ONITEMBUTTON) == 0) && (hClickedRow == reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0))) && IsExpanded(hClickedRow, FALSE)) {
			lr = TRUE;
		}
	}

	if(pNotificationDetails->code == 0) {
		// we called this method ourselves
	} else {
		if(!(GetStyle() & TVS_DISABLEDRAGDROP)) {
			if(hitTestFlags & (TVHT_ONITEMSTATEICON | TVHT_ONITEMICON | TVHT_ONITEMLABEL)) {
				// the treeview seems to eat the WM_LBUTTONUP message
				BOOL dummy = TRUE;
				OnLButtonUp(NULL, GetButtonStateBitField(), MAKELPARAM(mousePosition.x, mousePosition.y), dummy);
			}
		}
	}

	return lr;
}

LRESULT ExplorerTreeView::OnDblClkNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	labelEditStatus.watchingForEdit = FALSE;

	if(!(properties.disabledEvents & deTreeClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		button |= 1/*MouseButtonConstants.vbLeftButton*/;
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ScreenToClient(&mousePosition);
		Raise_DblClick(button, shift, mousePosition.x, mousePosition.y);
	}

	return 0;
}

LRESULT ExplorerTreeView::OnRClickNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& wasHandled)
{
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	POINT pt = mousePosition;
	ScreenToClient(&mousePosition);

	if(!(properties.disabledEvents & deTreeClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		button |= 2/*MouseButtonConstants.vbRightButton*/;
		Raise_RClick(button, shift, mousePosition.x, mousePosition.y);
	}

	// SysTreeView32 seems to eat all WM_RBUTTONUP messages
	BOOL dummy = TRUE;
	OnRButtonUp(WM_RBUTTONUP, GetButtonStateBitField(), MAKELPARAM(mousePosition.x, mousePosition.y), dummy);
	// WM_CONTEXTMENU is fucked up, too
	dummy = TRUE;
	OnContextMenu(WM_CONTEXTMENU, reinterpret_cast<WPARAM>(static_cast<HWND>(*this)), MAKELPARAM(pt.x, pt.y), dummy);

	wasHandled = TRUE;
	return TRUE;					// SysTreeView32 won't send WM_CONTEXTMENU if we return TRUE here
}

LRESULT ExplorerTreeView::OnRDblClkNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& wasHandled)
{
	if(!(properties.disabledEvents & deTreeClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		button |= 2/*MouseButtonConstants.vbRightButton*/;
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ScreenToClient(&mousePosition);
		Raise_RDblClick(button, shift, mousePosition.x, mousePosition.y);
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerTreeView::OnSetCursorNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMMOUSE pDetails = reinterpret_cast<LPNMMOUSE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMMOUSE);

	HCURSOR hCursor = NULL;

	HotTrackingConstants hotTracking = htrNone;
	get_HotTracking(&hotTracking);
	switch(hotTracking) {
		case htrNone:
		case htrWinXPStyle:
			if(properties.mousePointer == mpCustom) {
				if(properties.mouseIcon.pPictureDisp) {
					CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
					if(pPicture) {
						OLE_HANDLE h = NULL;
						pPicture->get_Handle(&h);
						hCursor = static_cast<HCURSOR>(LongToHandle(h));
					}
				}
			} else {
				if((hotTracking == htrWinXPStyle) && (properties.mousePointer == mpDefault)) {
					hCursor = MousePointerConst2hCursor(mpArrow);
				} else {
					hCursor = MousePointerConst2hCursor(properties.mousePointer);
				}
			}
			break;

		case htrNormal:
			// if we're above an item, use standard message handler
			if(pDetails && !pDetails->dwItemSpec) {
				if(properties.mousePointer == mpCustom) {
					if(properties.mouseIcon.pPictureDisp) {
						CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
						if(pPicture) {
							OLE_HANDLE h = NULL;
							pPicture->get_Handle(&h);
							hCursor = static_cast<HCURSOR>(LongToHandle(h));
						}
					}
				} else {
					hCursor = MousePointerConst2hCursor(properties.mousePointer);
				}
			}
			break;
	}

	if(hCursor) {
		SetCursor(hCursor);
		return TRUE;
	}

	return FALSE;
}

LRESULT ExplorerTreeView::OnSetFocusNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& wasHandled)
{
	if(!IsInDesignMode()) {
		// now that we've the focus, we should configure the IME
		IMEModeConstants ime = GetCurrentIMEContextMode(*this);
		if(ime != imeInherit) {
			ime = GetEffectiveIMEMode();
			if(ime != imeNoControl) {
				SetCurrentIMEContextMode(*this, ime);
			}
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerTreeView::OnTVStateImageChangingNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTVSTATEIMAGECHANGING pDetails = reinterpret_cast<LPNMTVSTATEIMAGECHANGING>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTVSTATEIMAGECHANGING);

	stateImageChangeFlags.hItemToResetStateImageFor = NULL;
	if(!pDetails) {
		return FALSE;
	}

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(pDetails->hti, this);
	VARIANT_BOOL cancel = VARIANT_FALSE;
	LONG newStateImageIndex = pDetails->iNewStateImageIndex;
	Raise_ItemStateImageChanging(pTvwItem, pDetails->iOldStateImageIndex, &newStateImageIndex, stateImageChangeFlags.reasonForPotentialStateImageChange, &cancel);

	if(cancel != VARIANT_FALSE) {
		return TRUE;
	}
	if(newStateImageIndex != static_cast<LONG>(pDetails->iNewStateImageIndex)) {
		if(newStateImageIndex == static_cast<LONG>(pDetails->iOldStateImageIndex)) {
			// simply cancel the whole thing
			return TRUE;
		} else {
			/* We can't simply tell SysTreeView32 to use another state image index, so advise
			   OnItemChangedNotification to change the state image afterwards. */
			stateImageChangeFlags.hItemToResetStateImageFor = pDetails->hti;
			stateImageChangeFlags.stateImageIndexToResetTo = newStateImageIndex;
		}
	}
	return FALSE;
}

LRESULT ExplorerTreeView::OnAsyncDrawNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	NMTVASYNCDRAW* pDetails = reinterpret_cast<NMTVASYNCDRAW*>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTVASYNCDRAW);
	if(!pDetails) {
		return 0;
	}

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(pDetails->hItem, this);
	Raise_ItemAsynchronousDrawFailed(pTvwItem, pDetails);
	return 0;
}

LRESULT ExplorerTreeView::OnBeginDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	stateImageChangeFlags.hClickedItem = NULL;

	LPNMTREEVIEW pDetails = reinterpret_cast<LPNMTREEVIEW>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_BEGINDRAGA) {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWW);
		}
	#endif

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	button |= 1/*MouseButtonConstants.vbLeftButton*/;
	POINT mousePosition = pDetails->ptDrag;
	UINT hitTestDetails = 0;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(pDetails->itemNew.hItem, this);
	Raise_ItemBeginDrag(pTvwItem, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			UINT message = (pNotificationDetails->code == TVN_BEGINDRAGA ? SHTVM_BEGINDRAGA : SHTVM_BEGINDRAGW);
			shellBrowserInterface.pInternalMessageListener->ProcessMessage(message, FALSE, reinterpret_cast<LPARAM>(pNotificationDetails));
		}
	#endif

	return 0;
}

LRESULT ExplorerTreeView::OnBeginLabelEditNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(!labelEditStatus.watchingForEdit) {
		cancel = VARIANT_TRUE;
	}
	labelEditStatus.watchingForEdit = FALSE;

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			UINT message = (pNotificationDetails->code == TVN_BEGINLABELEDITA ? SHTVM_BEGINLABELEDITA : SHTVM_BEGINLABELEDITW);
			HRESULT hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(message, VARIANTBOOL2BOOL(cancel), reinterpret_cast<LPARAM>(pNotificationDetails));
			cancel = BOOL2VARIANTBOOL(hr != S_OK);
		}
		if(lstrlen(labelEditStatus.pTextToSetOnBegin) > 0) {
			containedEdit.SetWindowText(labelEditStatus.pTextToSetOnBegin);
			labelEditStatus.pTextToSetOnBegin[0] = TEXT('\0');
		}
	#endif

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(reinterpret_cast<LPNMTVDISPINFOA>(pNotificationDetails)->item.hItem, this);
	Raise_StartingLabelEditing(pTvwItem, &cancel);
	if(cancel != VARIANT_FALSE) {
		return TRUE;
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerTreeView::OnBeginRDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	labelEditStatus.watchingForEdit = FALSE;
	stateImageChangeFlags.hClickedItem = NULL;

	LPNMTREEVIEW pDetails = reinterpret_cast<LPNMTREEVIEW>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_BEGINRDRAGA) {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWW);
		}
	#endif

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	button |= 2/*MouseButtonConstants.vbRightButton*/;
	POINT mousePosition = pDetails->ptDrag;
	UINT hitTestDetails = 0;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(pDetails->itemNew.hItem, this);
	Raise_ItemBeginRDrag(pTvwItem, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			UINT message = (pNotificationDetails->code == TVN_BEGINRDRAGA ? SHTVM_BEGINDRAGA : SHTVM_BEGINDRAGW);
			shellBrowserInterface.pInternalMessageListener->ProcessMessage(message, TRUE, reinterpret_cast<LPARAM>(pNotificationDetails));
		}
	#endif

	return 0;
}

LRESULT ExplorerTreeView::OnDeleteItemNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	wasHandled = FALSE;

	LPNMTREEVIEW pDetails = reinterpret_cast<LPNMTREEVIEW>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_DELETEITEMA) {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWW);
		}
	#endif

	#ifdef USE_STL
		std::vector<HTREEITEM>::iterator iter = std::find(itemDeletionsToIgnore.begin(), itemDeletionsToIgnore.end(), pDetails->itemOld.hItem);
		if(iter == itemDeletionsToIgnore.end()) {
	#else
		size_t iIgnoredItemDeletion = static_cast<size_t>(-1);
		for(size_t i = 0; i < itemDeletionsToIgnore.GetCount(); ++i) {
			if(itemDeletionsToIgnore[i] == pDetails->itemOld.hItem) {
				iIgnoredItemDeletion = i;
				break;
			}
		}
		if(iIgnoredItemDeletion == -1) {
	#endif
		if(!(properties.disabledEvents & deFreeItemData)) {
			CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(pDetails->itemOld.hItem, this);
			Raise_FreeItemData(pTvwItem);
		}

		#ifdef USE_STL
			std::unordered_map<LONG, HTREEITEM>::iterator iter2 = itemHandles.find(static_cast<const LONG>(pDetails->itemOld.lParam));
			if(iter2 != itemHandles.end()) {
				RemoveItemFromItemContainers(iter2->first);
				itemHandles.erase(iter2);
			}
			std::unordered_map<HTREEITEM, ItemData>::iterator iter3 = itemParams.find(pDetails->itemOld.hItem);
			if(iter3 != itemParams.end()) {
				itemParams.erase(iter3);
			}
		#else
			CAtlMap<LONG, HTREEITEM>::CPair* pEntry = itemHandles.Lookup(static_cast<LONG>(pDetails->itemOld.lParam));
			if(pEntry) {
				RemoveItemFromItemContainers(pEntry->m_key);
				itemHandles.RemoveKey(pEntry->m_key);
			}
			itemParams.RemoveKey(pDetails->itemOld.hItem);
		#endif

		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pInternalMessageListener) {
				shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_FREEITEM, reinterpret_cast<WPARAM>(pDetails->itemOld.hItem), pDetails->itemOld.lParam);
			}
		#endif
	} else {
		#ifdef USE_STL
			itemDeletionsToIgnore.erase(iter);
		#else
			itemDeletionsToIgnore.RemoveAt(iIgnoredItemDeletion);
		#endif
	}

	// keep the status structs up to date
	if(pDetails->itemOld.hItem == dragDropStatus.hLastDropTarget) {
		dragDropStatus.hLastDropTarget = NULL;
	}
	if(pDetails->itemOld.hItem == insertMark.hItem) {
		insertMark.hItem = NULL;
	}
	if(pDetails->itemOld.hItem == stateImageChangeFlags.hClickedItem) {
		stateImageChangeFlags.hClickedItem = NULL;
	}
	return 0;
}

LRESULT ExplorerTreeView::OnEndLabelEditNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTVDISPINFO pDetails = reinterpret_cast<LPNMTVDISPINFO>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_ENDLABELEDITA) {
			ATLASSERT_POINTER(pDetails, NMTVDISPINFOA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTVDISPINFOW);
		}
	#endif

	if(pDetails->item.pszText) {
		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(pDetails->item.hItem, this);
		CComBSTR newItemText;
		if(pNotificationDetails->code == TVN_ENDLABELEDITA) {
			newItemText = reinterpret_cast<LPNMTVDISPINFOA>(pDetails)->item.pszText;
		} else {
			newItemText = reinterpret_cast<LPNMTVDISPINFOW>(pDetails)->item.pszText;
		}
		BSTR itemText = NULL;
		pTvwItem->get_Text(&itemText);

		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_RenamingItem(pTvwItem, itemText, newItemText, &cancel);
		if(cancel == VARIANT_FALSE) {
			labelEditStatus.hEditedItem = pDetails->item.hItem;
			labelEditStatus.previousText = itemText;
			labelEditStatus.acceptText = TRUE;

			#ifdef INCLUDESHELLBROWSERINTERFACE
				if(shellBrowserInterface.pInternalMessageListener) {
					WCHAR pNewName[MAX_ITEMTEXTLENGTH + 1];
					lstrcpynW(pNewName, newItemText, MAX_ITEMTEXTLENGTH + 1);
					#ifdef UNICODE
						lstrcpyn(labelEditStatus.pTextToSetOnBegin, pNewName, MAX_ITEMTEXTLENGTH + 1);
					#else
						CW2A converter(pNewName);
						lstrcpyn(labelEditStatus.pTextToSetOnBegin, converter, MAX_ITEMTEXTLENGTH + 1);
					#endif
					HRESULT hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_RENAMEDITEM, reinterpret_cast<WPARAM>(labelEditStatus.hEditedItem), reinterpret_cast<LPARAM>(pNewName));
					labelEditStatus.acceptText = (hr == S_OK);
					if(labelEditStatus.acceptText) {
						labelEditStatus.pTextToSetOnBegin[0] = TEXT('\0');
					}
				}
			#endif
			SysFreeString(itemText);
			return TRUE;
		}
	}
	return FALSE;
}

LRESULT ExplorerTreeView::OnGetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	if(properties.disabledEvents & deItemGetDisplayInfo) {
		return 0;
	}

	LPNMTVDISPINFO pDetails = reinterpret_cast<LPNMTVDISPINFO>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_GETDISPINFOA) {
			ATLASSERT_POINTER(pDetails, NMTVDISPINFOA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTVDISPINFOW);
		}
	#endif

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(pDetails->item.hItem, this);
	LONG iconIndex = -2;
	LONG selectedIconIndex = -2;
	LONG expandedIconIndex = -2;
	VARIANT_BOOL hasButton = VARIANT_FALSE;
	CComBSTR itemText = L"";
	LONG maxItemTextLength = 0;
	RequestedInfoConstants requestedInfo = static_cast<RequestedInfoConstants>(0);
	if(pDetails->item.mask & TVIF_IMAGE) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riIconIndex);
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pMessageListener) {
				iconIndex = pDetails->item.iImage;
			}
		#endif
	}
	if(pDetails->item.mask & TVIF_SELECTEDIMAGE) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riSelectedIconIndex);
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pMessageListener) {
				selectedIconIndex = pDetails->item.iSelectedImage;
			}
		#endif
	}
	if(pDetails->item.mask & TVIF_EXPANDEDIMAGE) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riExpandedIconIndex);
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pMessageListener) {
				expandedIconIndex = reinterpret_cast<LPNMTVDISPINFOEXA>(pDetails)->item.iExpandedImage;
			}
		#endif
	}
	if(pDetails->item.mask & TVIF_CHILDREN) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riHasExpando);
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pMessageListener) {
				hasButton = BOOL2VARIANTBOOL(pDetails->item.cChildren == heYes);
			}
		#endif
	}
	if(pDetails->item.mask & TVIF_TEXT) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riItemText);
		// exclude the terminating 0
		maxItemTextLength = pDetails->item.cchTextMax - 1;
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pMessageListener) {
				if(pNotificationDetails->code == TVN_GETDISPINFOA) {
					itemText = reinterpret_cast<LPNMTVDISPINFOA>(pDetails)->item.pszText;
				} else {
					itemText = reinterpret_cast<LPNMTVDISPINFOW>(pDetails)->item.pszText;
				}
			}
		#endif
	}

	VARIANT_BOOL dontAskAgain = VARIANT_FALSE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pMessageListener) {
			dontAskAgain = BOOL2VARIANTBOOL((pDetails->item.mask & TVIF_DI_SETITEM) == TVIF_DI_SETITEM);
		}
	#endif
	Raise_ItemGetDisplayInfo(pTvwItem, requestedInfo, &iconIndex, &selectedIconIndex, &expandedIconIndex, &hasButton, maxItemTextLength, &itemText, &dontAskAgain);

	if(pDetails->item.mask & TVIF_IMAGE) {
		pDetails->item.iImage = iconIndex;
	}
	if(pDetails->item.mask & TVIF_SELECTEDIMAGE) {
		pDetails->item.iSelectedImage = selectedIconIndex;
	}
	if(pDetails->item.mask & TVIF_EXPANDEDIMAGE) {
		reinterpret_cast<LPNMTVDISPINFOEXA>(pDetails)->item.iExpandedImage = expandedIconIndex;
	}
	if(pDetails->item.mask & TVIF_CHILDREN) {
		pDetails->item.cChildren = (VARIANTBOOL2BOOL(hasButton) ? heYes : heNo);
	}
	if(pDetails->item.mask & TVIF_TEXT) {
		// don't forget the terminating 0
		int bufferSize = itemText.Length() + 1;
		if(bufferSize > pDetails->item.cchTextMax) {
			bufferSize = pDetails->item.cchTextMax;
		}
		if(pNotificationDetails->code == TVN_GETDISPINFOA) {
			if(bufferSize > 1) {
				lstrcpynA(reinterpret_cast<LPNMTVDISPINFOA>(pDetails)->item.pszText, CW2A(itemText), bufferSize);
			} else {
				reinterpret_cast<LPNMTVDISPINFOA>(pDetails)->item.pszText[0] = '\0';
			}
		} else {
			if(bufferSize > 1) {
				lstrcpynW(reinterpret_cast<LPNMTVDISPINFOW>(pDetails)->item.pszText, itemText, bufferSize);
			} else {
				reinterpret_cast<LPNMTVDISPINFOW>(pDetails)->item.pszText[0] = L'\0';
			}
		}
	}

	if(dontAskAgain == VARIANT_FALSE) {
		pDetails->item.mask &= ~TVIF_DI_SETITEM;
	} else {
		pDetails->item.mask |= TVIF_DI_SETITEM;
	}
	return 0;
}

LRESULT ExplorerTreeView::OnGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTVGETINFOTIP pDetails = reinterpret_cast<LPNMTVGETINFOTIP>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_GETINFOTIPA) {
			ATLASSERT_POINTER(pDetails, NMTVGETINFOTIPA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTVGETINFOTIPW);
		}
	#endif

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(pDetails->hItem, this);
	CComBSTR text = L"";
	TCHAR pItemText[MAX_ITEMTEXTLENGTH + 1];
	BOOL isTruncated = IsTreeViewItemTruncated(*this, pDetails->hItem, pItemText);
	if(isTruncated) {
		text.Append(pItemText);
	}
	// exclude the terminating 0
	LONG maxTextLength = pDetails->cchTextMax - 1;

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			LPWSTR pInfoTip = NULL;
			shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_GETINFOTIP, reinterpret_cast<WPARAM>(pDetails->hItem), reinterpret_cast<LPARAM>(&pInfoTip));
			if(pInfoTip) {
				if(lstrlenW(pInfoTip) > 0) {
					if(text.Length() > 0) {
						text.Append(TEXT("\r\n"));
					}
					text.Append(pInfoTip);
				}
				CoTaskMemFree(pInfoTip);
			}
		}
	#endif
	VARIANT_BOOL abortToolTip = VARIANT_FALSE;
	Raise_ItemGetInfoTipText(pTvwItem, maxTextLength, &text, &abortToolTip);
	flags.cancelToolTip = VARIANTBOOL2BOOL(abortToolTip);

	if(!(isTruncated && (text == pItemText))) {
		// don't forget the terminating 0
		int bufferSize = text.Length() + 1;
		if(bufferSize > pDetails->cchTextMax) {
			bufferSize = pDetails->cchTextMax;
		}
		if(pNotificationDetails->code == TVN_GETINFOTIPA) {
			int requiredSize = WideCharToMultiByte(CP_ACP, 0, text, bufferSize, "", 0, "", NULL);
			if(requiredSize > 0) {
				LPSTR pBuffer = reinterpret_cast<LPSTR>(HeapAlloc(GetProcessHeap(), 0, (requiredSize + 1) * sizeof(CHAR)));
				if(pBuffer) {
					WideCharToMultiByte(CP_ACP, 0, text, bufferSize, pBuffer, requiredSize, "", NULL);
					/* On Vista and newer, many info tips contain Unicode characters that cannot be converted.
					 * We have told WideCharToMultiByte to replace them with \0. Now remove those \0.
					 */
					LPSTR pDst = reinterpret_cast<LPNMTVGETINFOTIPA>(pDetails)->pszText;
					ZeroMemory(pDst, pDetails->cchTextMax);
					int j = 0;
					for(int i = 0; i < min(bufferSize, requiredSize); i++) {
						if(pBuffer[i] != '\0') {
							pDst[j++] = pBuffer[i];
						}
					}
					HeapFree(GetProcessHeap(), 0, pBuffer);
				}
			}
		} else {
			lstrcpynW(reinterpret_cast<LPNMTVGETINFOTIPW>(pDetails)->pszText, text, bufferSize);
		}
	}

	return 0;
}

LRESULT ExplorerTreeView::OnItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	NMTVITEMCHANGE* pDetails = reinterpret_cast<NMTVITEMCHANGE*>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTVITEMCHANGE);
	if(!pDetails) {
		return 0;
	}

	if(pDetails->uChanged & TVIF_STATE) {
		if(!(properties.disabledEvents & deItemSelectionChangingEvents)) {
			if((pDetails->uStateOld & TVIS_SELECTED) != (pDetails->uStateNew & TVIS_SELECTED)) {
				Raise_ItemSelectionChanged(pDetails->hItem);
			}
		}

		if((stateImageChangeFlags.silentStateImageChanges == 0) && ((pDetails->uStateOld & TVIS_STATEIMAGEMASK) != (pDetails->uStateNew & TVIS_STATEIMAGEMASK))) {
			LONG newStateImageIndex = static_cast<LONG>((pDetails->uStateNew & TVIS_STATEIMAGEMASK) >> 12);
			if(pDetails->hItem == stateImageChangeFlags.hItemToResetStateImageFor) {
				// customize the state image index
				stateImageChangeFlags.EnterNoStateImageChangeEventsSection();

				TVITEMEX item = {0};
				item.hItem = pDetails->hItem;
				item.state = INDEXTOSTATEIMAGEMASK(stateImageChangeFlags.stateImageIndexToResetTo);
				item.stateMask = TVIS_STATEIMAGEMASK;
				item.mask = TVIF_HANDLE | TVIF_STATE;
				if(SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
					newStateImageIndex = stateImageChangeFlags.stateImageIndexToResetTo;
				}

				stateImageChangeFlags.LeaveNoStateImageChangeEventsSection();
			}

			if(stateImageChangeFlags.reasonForPotentialStateImageChange != static_cast<StateImageChangeCausedByConstants>(-1)) {
				CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(pDetails->hItem, this);
				Raise_ItemStateImageChanged(pTvwItem, static_cast<LONG>((pDetails->uStateOld & TVIS_STATEIMAGEMASK) >> 12), newStateImageIndex, stateImageChangeFlags.reasonForPotentialStateImageChange);
			}
		}
	}

	return 0;
}

LRESULT ExplorerTreeView::OnItemChangingNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	NMTVITEMCHANGE* pDetails = reinterpret_cast<NMTVITEMCHANGE*>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTVITEMCHANGE);
	if(!pDetails) {
		return FALSE;
	}

	if(pDetails->uChanged & TVIF_STATE) {
		if(!(properties.disabledEvents & deItemSelectionChangingEvents)) {
			if((pDetails->uStateOld & TVIS_SELECTED) != (pDetails->uStateNew & TVIS_SELECTED)) {
				VARIANT_BOOL cancel = VARIANT_FALSE;
				Raise_ItemSelectionChanging(pDetails->hItem, &cancel);
				if(cancel != VARIANT_FALSE) {
					return TRUE;
				}
			}
		}

		if((stateImageChangeFlags.silentStateImageChanges == 0) && (pDetails->uStateOld & TVIS_STATEIMAGEMASK) != (pDetails->uStateNew & TVIS_STATEIMAGEMASK)) {
			if(stateImageChangeFlags.hItemCurrentInternalStateImageChange == pDetails->hItem) {
				stateImageChangeFlags.reasonForPotentialStateImageChange = siccbCode;

				stateImageChangeFlags.hItemToResetStateImageFor = NULL;

				CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(pDetails->hItem, this);
				VARIANT_BOOL cancel = VARIANT_FALSE;
				LONG newStateImageIndex = static_cast<LONG>((pDetails->uStateNew & TVIS_STATEIMAGEMASK) >> 12);
				Raise_ItemStateImageChanging(pTvwItem, static_cast<LONG>((pDetails->uStateOld & TVIS_STATEIMAGEMASK) >> 12), &newStateImageIndex, stateImageChangeFlags.reasonForPotentialStateImageChange, &cancel);

				if(cancel != VARIANT_FALSE) {
					return TRUE;
				}
				if(newStateImageIndex != static_cast<LONG>((pDetails->uStateNew & TVIS_STATEIMAGEMASK) >> 12)) {
					if(newStateImageIndex == static_cast<LONG>((pDetails->uStateOld & TVIS_STATEIMAGEMASK) >> 12)) {
						// simply cancel the whole thing
						return TRUE;
					} else {
						/* We can't simply tell SysTreeView32 to use another state image index, so advise
						   OnItemChangedNotification to change the state image afterwards. */
						stateImageChangeFlags.hItemToResetStateImageFor = pDetails->hItem;
						stateImageChangeFlags.stateImageIndexToResetTo = newStateImageIndex;
					}
				}
			}
		}
	}

	return FALSE;
}

LRESULT ExplorerTreeView::OnItemExpandedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	wasHandled = FALSE;

	LPNMTREEVIEW pDetails = reinterpret_cast<LPNMTREEVIEW>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_ITEMEXPANDEDA) {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWW);
		}
	#endif

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(pDetails->itemNew.hItem, this);
	if(pDetails->action & TVE_COLLAPSE) {
		if(pDetails->action & TVE_COLLAPSERESET) {
			Raise_CollapsedItem(pTvwItem, VARIANT_TRUE);
		} else {
			Raise_CollapsedItem(pTvwItem, VARIANT_FALSE);
		}
	} else if(pDetails->action & TVE_EXPAND) {
		if(pDetails->action & TVE_EXPANDPARTIAL) {
			Raise_ExpandedItem(pTvwItem, VARIANT_TRUE);
		} else {
			Raise_ExpandedItem(pTvwItem, VARIANT_FALSE);
		}
	}

	if(properties.favoritesStyle) {
		Invalidate();
	}
	return 0;
}

LRESULT ExplorerTreeView::OnItemExpandingNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	LPNMTREEVIEW pDetails = reinterpret_cast<LPNMTREEVIEW>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_ITEMEXPANDINGA) {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWW);
		}
	#endif

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(pDetails->itemNew.hItem, this);
	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(pDetails->action & TVE_COLLAPSE) {
		VARIANT_BOOL removeAll = VARIANT_FALSE;
		if(pDetails->action & TVE_COLLAPSERESET) {
			removeAll = VARIANT_TRUE;
		}
		Raise_CollapsingItem(pTvwItem, removeAll, &cancel);
	} else if(pDetails->action & TVE_EXPAND) {
		VARIANT_BOOL expandPartial = VARIANT_FALSE;
		if(pDetails->action & TVE_EXPANDPARTIAL) {
			expandPartial = VARIANT_TRUE;
		}
		Raise_ExpandingItem(pTvwItem, expandPartial, &cancel);
	}

	if(cancel != VARIANT_FALSE) {
		return TRUE;
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerTreeView::OnKeyDownNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deTreeKeyboardEvents)) {
		BOOL isChar = FALSE;
		LPNMTVKEYDOWN pDetails = reinterpret_cast<LPNMTVKEYDOWN>(pNotificationDetails);
		ATLASSERT_POINTER(pDetails, NMTVKEYDOWN);
		if(!pDetails) {
			return FALSE;
		}

		////////////////////////////////////////////////////////////////
		// taken from wine\dlls\user32\message.c -> TranslateMessage() (Wine 0.9.52)
		WCHAR wp[2];
		BYTE state[256];

		GetKeyboardState(state);
		/* FIXME : should handle ToUnicode yielding 2 */
		switch(ToUnicode(pDetails->wVKey, HIWORD(lastKeyDownLParam), state, wp, 2, 0)) {
			case -1:
			case 1:
				isChar = TRUE;
				break;
		}
		////////////////////////////////////////////////////////////////

		VARIANT_BOOL cancel = VARIANT_FALSE;
		if(isChar) {
			BSTR searchString = SysAllocString(L"");
			get_IncrementalSearchString(&searchString);
			Raise_IncrementalSearchStringChanging(searchString, static_cast<SHORT>(pDetails->wVKey), &cancel);
			SysFreeString(searchString);
		}
		if(cancel != VARIANT_FALSE) {
			// exclude the char
			wasHandled = TRUE;
			return TRUE;
		}
	}
	return FALSE;
}

LRESULT ExplorerTreeView::OnCaretChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	labelEditStatus.watchingForEdit = FALSE;

	// complete any outstanding tasks
	caretChangeTasks.ProcessOnCaretChangedTasks(this, *this);

	LPNMTREEVIEW pDetails = reinterpret_cast<LPNMTREEVIEW>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_SELCHANGEDA) {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWW);
		}
	#endif

	if(properties.favoritesStyle) {
		Invalidate();
	}

	if(!flags.removingAllItems) {
		Raise_CaretChanged(pDetails->itemOld.hItem, pDetails->itemNew.hItem, pDetails->action);
	}
	return 0;
}

LRESULT ExplorerTreeView::OnCaretChangingNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	labelEditStatus.watchingForEdit = FALSE;

	// complete any outstanding tasks
	caretChangeTasks.ProcessOnCaretChangingTasks(this, *this);

	LPNMTREEVIEW pDetails = reinterpret_cast<LPNMTREEVIEW>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_SELCHANGINGA) {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWW);
		}
	#endif

	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(!flags.removingAllItems) {
		Raise_CaretChanging(pDetails->itemOld.hItem, pDetails->itemNew.hItem, pDetails->action, &cancel);
	}

	return VARIANTBOOL2BOOL(cancel);
}

LRESULT ExplorerTreeView::OnSetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTVDISPINFO pDetails = reinterpret_cast<LPNMTVDISPINFO>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_SETDISPINFOA) {
			ATLASSERT_POINTER(pDetails, NMTVDISPINFOA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTVDISPINFOW);
		}
	#endif

	ATLASSERT(pDetails->item.mask & TVIF_TEXT);
	if(pDetails->item.mask & TVIF_TEXT) {
		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(pDetails->item.hItem, this);
		CComBSTR itemText;
		if(pNotificationDetails->code == TVN_SETDISPINFOA) {
			itemText = reinterpret_cast<LPNMTVDISPINFOA>(pDetails)->item.pszText;
		} else {
			itemText = reinterpret_cast<LPNMTVDISPINFOW>(pDetails)->item.pszText;
		}
		Raise_ItemSetText(pTvwItem, itemText);
	}
	return 0;
}

LRESULT ExplorerTreeView::OnSingleExpandNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTREEVIEW pDetails = reinterpret_cast<LPNMTREEVIEW>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTREEVIEW);
	if(!pDetails) {
		return 0;
	}

	CComPtr<ITreeViewItem> pPreviousCaretItem = ClassFactory::InitTreeItem(pDetails->itemOld.hItem, this);
	CComPtr<ITreeViewItem> pNewCaretItem = ClassFactory::InitTreeItem(pDetails->itemNew.hItem, this);
	VARIANT_BOOL skipOld = VARIANT_FALSE;
	VARIANT_BOOL skipNew = VARIANT_FALSE;

	if(flags.dontSingleExpand > 0) {
		skipOld = VARIANT_TRUE;
		skipNew = VARIANT_TRUE;
	} else if(GetAsyncKeyState(VK_CONTROL) & 0x8000) {
		// SysTreeView32 will expand the new caret and leave the previous caret unchanged
		skipOld = VARIANT_TRUE;
		skipNew = BOOL2VARIANTBOOL(IsExpanded(pDetails->itemNew.hItem, FALSE));
	} else {
		SingleExpandConstants singleExpand = seNone;
		get_SingleExpand(&singleExpand);
		if(singleExpand == seWinXPStyle) {
			if(!HaveSameParentItem(pDetails->itemNew.hItem, pDetails->itemOld.hItem)) {
				// we're handling 2 different branches - skip handling the old caret
				skipOld = VARIANT_TRUE;
			}
			if(IsExpanded(pDetails->itemNew.hItem, FALSE)) {
				// the new caret already is expanded
				skipNew = VARIANT_TRUE;
			}
		}
	}

	Raise_SingleExpanding(pPreviousCaretItem, pNewCaretItem, &skipOld, &skipNew);

	LRESULT lr = 0;
	if(skipOld != VARIANT_FALSE) {
		lr |= TVNRET_SKIPOLD;
	}
	if(skipNew != VARIANT_FALSE) {
		lr |= TVNRET_SKIPNEW;
	}
	if(lr == 0) {
		lr = TVNRET_DEFAULT;
	}
	return lr;
}


LRESULT ExplorerTreeView::OnEditChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_EditChange();
	return 0;
}


LRESULT ExplorerTreeView::OnCustomDrawNotification(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(flags.noCustomDraw) {
		return CDRF_DODEFAULT;
	}

	LPNMTVCUSTOMDRAW pDetails = reinterpret_cast<LPNMTVCUSTOMDRAW>(lParam);
	ATLASSERT_POINTER(pDetails, NMTVCUSTOMDRAW);
	if(!pDetails) {
		return DefWindowProc(message, wParam, lParam);
	}

	ATLASSERT((pDetails->nmcd.uItemState & CDIS_NEARHOT) == 0);
	ATLASSERT((pDetails->nmcd.uItemState & CDIS_OTHERSIDEHOT) == 0);
	ATLASSERT((pDetails->nmcd.uItemState & CDIS_DROPHILITED) == 0);

	CustomDrawReturnValuesConstants returnValue = static_cast<CustomDrawReturnValuesConstants>(DefWindowProc(message, wParam, lParam));
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			returnValue = static_cast<CustomDrawReturnValuesConstants>(shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_CUSTOMDRAW, wParam, lParam));
		}
	#endif

	if(!(properties.disabledEvents & deCustomDraw)) {
		CComPtr<ITreeViewItem> pTvwItem = NULL;
		LONG level = 0;
		OLE_COLOR colorText = 0;
		OLE_COLOR colorTextBackground = 0;
		if(pDetails->nmcd.dwDrawStage & (CDDS_ITEM | CDDS_SUBITEM)) {
			pTvwItem = ClassFactory::InitTreeItem(reinterpret_cast<HTREEITEM>(pDetails->nmcd.dwItemSpec), this);
			level = pDetails->iLevel;
			colorText = pDetails->clrText;
			colorTextBackground = pDetails->clrTextBk;
		}
		Raise_CustomDraw(pTvwItem, &level, &colorText, &colorTextBackground, static_cast<CustomDrawStageConstants>(pDetails->nmcd.dwDrawStage), static_cast<CustomDrawItemStateConstants>(pDetails->nmcd.uItemState), HandleToLong(pDetails->nmcd.hdc), reinterpret_cast<RECTANGLE*>(&pDetails->nmcd.rc), &returnValue);

		if(pDetails->nmcd.dwDrawStage & (CDDS_ITEM | CDDS_SUBITEM)) {
			pDetails->iLevel = level;
			pDetails->clrText = colorText;
			pDetails->clrTextBk = colorTextBackground;
		}
	}

	// FavoritesStyle support
	if(properties.favoritesStyle) {
		if(pDetails->nmcd.dwDrawStage == CDDS_PREPAINT) {
			returnValue = static_cast<CustomDrawReturnValuesConstants>(static_cast<LRESULT>(returnValue) | CDRF_NOTIFYPOSTPAINT);
		} else if(pDetails->nmcd.dwDrawStage == CDDS_POSTPAINT) {
			// retrieve the group's start and end position
			HTREEITEM hCaretItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0));
			if(hCaretItem) {
				HTREEITEM hParentItem = NULL;
				HTREEITEM hLastSubItem = NULL;

				if(IsExpanded(hCaretItem)) {
					hParentItem = hCaretItem;
					HTREEITEM hItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(hParentItem)));
					while(hItem) {
						hLastSubItem = hItem;
						hItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hItem)));
					}
				} else {
					hParentItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hCaretItem)));
					if(hParentItem) {
						HTREEITEM hItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(hParentItem)));
						while(hItem) {
							hLastSubItem = hItem;
							hItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hItem)));
						}
					} else {
						hParentItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_ROOT), 0));
						hLastSubItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_LASTVISIBLE), 0));
					}
				}

				// calculate the rectangle
				RECT itemRectangle = {0};
				*reinterpret_cast<HTREEITEM*>(&itemRectangle) = hParentItem;
				SendMessage(TVM_GETITEMRECT, static_cast<WPARAM>(FALSE), reinterpret_cast<LPARAM>(&itemRectangle));
				RECT groupBoxRectangle = itemRectangle;
				*reinterpret_cast<HTREEITEM*>(&itemRectangle) = hLastSubItem;
				SendMessage(TVM_GETITEMRECT, static_cast<WPARAM>(FALSE), reinterpret_cast<LPARAM>(&itemRectangle));
				groupBoxRectangle.bottom = itemRectangle.bottom;

				// now draw the rectangle
				HPEN hPen = CreatePen(PS_SOLID, 1, OLECOLOR2COLORREF(properties.groupBoxColor));
				ATLASSERT(hPen != NULL);
				if(hPen) {
					HGDIOBJ hPreviousPen = SelectObject(pDetails->nmcd.hdc, hPen);

					MoveToEx(pDetails->nmcd.hdc, groupBoxRectangle.left, groupBoxRectangle.top, NULL);
					LineTo(pDetails->nmcd.hdc, groupBoxRectangle.right - 1, groupBoxRectangle.top);
					LineTo(pDetails->nmcd.hdc, groupBoxRectangle.right - 1, groupBoxRectangle.bottom);
					LineTo(pDetails->nmcd.hdc, groupBoxRectangle.left, groupBoxRectangle.bottom);
					LineTo(pDetails->nmcd.hdc, groupBoxRectangle.left, groupBoxRectangle.top - 1);

					SelectObject(pDetails->nmcd.hdc, hPreviousPen);
					DeleteObject(hPen);
				}
			}
		}
	}

	// BlendSelectedItemsIcons support
	if(cachedSettings.hImageList) {
		if(properties.blendSelectedItemsIcons) {
			if(pDetails->nmcd.dwDrawStage == CDDS_PREPAINT) {
				returnValue = static_cast<CustomDrawReturnValuesConstants>(static_cast<LRESULT>(returnValue) | CDRF_NOTIFYITEMDRAW);
			} else if(pDetails->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
				if(((returnValue & TVCDRF_NOIMAGES) == 0) && ((pDetails->nmcd.uItemState & CDIS_SELECTED) == CDIS_SELECTED)) {
					TVITEMEX item = {0};
					item.hItem = reinterpret_cast<HTREEITEM>(pDetails->nmcd.dwItemSpec);
					item.stateMask = TVIS_CUT | TVIS_OVERLAYMASK;
					item.mask = TVIF_HANDLE | TVIF_SELECTEDIMAGE | TVIF_STATE;
					if(SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
						if(!(item.state & TVIS_CUT)) {
							int iconIndex = item.iSelectedImage;

							/* It's the easiest way to retrieve the item's rectangle and calculate the icon's
							   position based on it. */
							RECT itemRectangle = {0};
							*reinterpret_cast<HTREEITEM*>(&itemRectangle) = reinterpret_cast<HTREEITEM>(pDetails->nmcd.dwItemSpec);
							SendMessage(TVM_GETITEMRECT, static_cast<WPARAM>(TRUE), reinterpret_cast<LPARAM>(&itemRectangle));
							int x = itemRectangle.left - cachedSettings.iconWidth - TV_TEXTINDENT;
							int y = itemRectangle.top + max(cachedSettings.itemHeight - cachedSettings.iconHeight, 0) / 2;

							ImageList_DrawEx(cachedSettings.hImageList, iconIndex, pDetails->nmcd.hdc, x, y, cachedSettings.iconWidth, cachedSettings.iconHeight, CLR_HILIGHT, CLR_HILIGHT, ILD_SELECTED | ILD_TRANSPARENT | (item.state & TVIS_OVERLAYMASK));

							returnValue = static_cast<CustomDrawReturnValuesConstants>(static_cast<LRESULT>(returnValue) | TVCDRF_NOIMAGES);
						}
					}
				}
			}
		}
	}

	return static_cast<LRESULT>(returnValue);
}

LRESULT ExplorerTreeView::OnToolTipGetDispInfoNotificationA(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(lParam);
	ATLASSERT_POINTER(pDetails, NMTTDISPINFOA);
	if(!pDetails) {
		return DefWindowProc(message, wParam, lParam);
	}

	NMTTDISPINFOA backup = *pDetails;

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(flags.cancelToolTip) {
		flags.cancelToolTip = FALSE;
		/* TODO: Find a better way than copying the old values back. Don't we produce a memleak here?
		         Who frees the string buffer (pDetails->lpszText) that the treeview set up? */
		*pDetails = backup;
		lr = 0;     // actually the return value is ignored
	}

	return lr;
}

LRESULT ExplorerTreeView::OnToolTipGetDispInfoNotificationW(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(lParam);
	ATLASSERT_POINTER(pDetails, NMTTDISPINFOW);
	if(!pDetails) {
		return DefWindowProc(message, wParam, lParam);
	}

	NMTTDISPINFOW backup = *pDetails;

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(flags.cancelToolTip) {
		flags.cancelToolTip = FALSE;
		/* TODO: Find a better way than copying the old values back. Don't we produce a memleak here?
		         Who frees the string buffer (pDetails->lpszText) that the treeview set up? */
		*pDetails = backup;
		lr = 0;     // actually the return value is ignored
	}

	return lr;
}


void ExplorerTreeView::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// retrieve the system font
			LOGFONT iconFont = {0};
			SystemParametersInfo(SPI_GETICONTITLELOGFONT, sizeof(LOGFONT), &iconFont, 0);
			properties.font.currentFont.CreateFontIndirect(&iconFont);
			properties.font.owningFontResource = TRUE;

			// apply the font
			SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
		} else {
			/* The whole font object or at least some of its attributes were changed. 'font.pFontDisp' is
			   still valid, so simply update the treeview's font. */
			if(properties.font.pFontDisp) {
				CComQIPtr<IFont, &IID_IFont> pFont(properties.font.pFontDisp);
				if(pFont) {
					HFONT hFont = NULL;
					pFont->get_hFont(&hFont);
					properties.font.currentFont.Attach(hFont);
					properties.font.owningFontResource = FALSE;

					SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
				} else {
					SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
				}
			} else {
				SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
			}
			Invalidate();
		}
	}
	properties.font.dontGetFontObject = FALSE;
	FireViewChange();
}


inline HRESULT ExplorerTreeView::Raise_AbortedDrag(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_AbortedDrag();
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_CaretChanged(HTREEITEM hPreviousCaretItem, HTREEITEM hNewCaretItem, UINT caretChangeReason, BOOL delay/* = TRUE*/)
{
	if(caretChangeReason == TVC_UNKNOWN) {
		// try to find out the reason
		#ifdef USE_STL
			std::unordered_map<HTREEITEM, UINT>::const_iterator iter = internalCaretChanges.find(hNewCaretItem);
			if(iter != internalCaretChanges.end()) {
				caretChangeReason = iter->second;
				if(caretChangeReason == static_cast<UINT>(-1)) {
					// don't raise the event
					return S_OK;
				}
			}
		#else
			CAtlMap<HTREEITEM, UINT>::CPair* pEntry = internalCaretChanges.Lookup(hNewCaretItem);
			if(pEntry) {
				caretChangeReason = pEntry->m_value;
				if(caretChangeReason == static_cast<UINT>(-1)) {
					// don't raise the event
					return S_OK;
				}
			}
		#endif
	}

	if(delay && (properties.caretChangedDelayTime > 0)) {
		if(caretChangeReason == TVC_BYKEYBOARD) {
			// delay this event
			if(caretChangedDelayStatus.timeOfCaretChange == caretChangedDelayStatus.UNDEFINEDTIME) {
				// NOTE: We use a smaller interval because this increases precision.
				SetTimer(timers.ID_DELAYEDCARETCHANGE, properties.caretChangedDelayTime / 2);
			}
			caretChangedDelayStatus.Initialize(hNewCaretItem, hPreviousCaretItem);
			return S_OK;
		}
	}

	//if(m_nFreezeEvents == 0) {
		// fire the event
		CComPtr<ITreeViewItem> pTvwNewCaretItem = ClassFactory::InitTreeItem(hNewCaretItem, this);
		CComPtr<ITreeViewItem> pTvwPrevCaretItem = ClassFactory::InitTreeItem(hPreviousCaretItem, this);
		switch(caretChangeReason) {
			case TVC_BYKEYBOARD:
				return Fire_CaretChanged(pTvwPrevCaretItem, pTvwNewCaretItem, cccbKeyboard);
			case TVC_BYMOUSE:
				return Fire_CaretChanged(pTvwPrevCaretItem, pTvwNewCaretItem, cccbMouse);
			case TVC_UNKNOWN:
				return Fire_CaretChanged(pTvwPrevCaretItem, pTvwNewCaretItem, cccbUnknown);
			case TVC_INTERNAL:
				return Fire_CaretChanged(pTvwPrevCaretItem, pTvwNewCaretItem, cccbInternal);
			default:
				return Fire_CaretChanged(pTvwPrevCaretItem, pTvwNewCaretItem, cccbUnknown);
		}
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_CaretChanging(HTREEITEM hPreviousCaretItem, HTREEITEM hNewCaretItem, UINT caretChangeReason, VARIANT_BOOL* pCancel)
{
	if(caretChangeReason == TVC_UNKNOWN) {
		// try to find out the reason
		#ifdef USE_STL
			std::unordered_map<HTREEITEM, UINT>::const_iterator iter = internalCaretChanges.find(hNewCaretItem);
			if(iter != internalCaretChanges.end()) {
				caretChangeReason = iter->second;
				if(caretChangeReason == static_cast<UINT>(-1)) {
					// don't raise the event
					return S_OK;
				}
			}
		#else
			CAtlMap<HTREEITEM, UINT>::CPair* pEntry = internalCaretChanges.Lookup(hNewCaretItem);
			if(pEntry) {
				caretChangeReason = pEntry->m_value;
				if(caretChangeReason == static_cast<UINT>(-1)) {
					// don't raise the event
					return S_OK;
				}
			}
		#endif
	}

	//if(m_nFreezeEvents == 0) {
		// fire the event
		CComPtr<ITreeViewItem> pTvwNewCaretItem = ClassFactory::InitTreeItem(hNewCaretItem, this);
		CComPtr<ITreeViewItem> pTvwPrevCaretItem = ClassFactory::InitTreeItem(hPreviousCaretItem, this);
		switch(caretChangeReason) {
			case TVC_BYKEYBOARD:
				return Fire_CaretChanging(pTvwPrevCaretItem, pTvwNewCaretItem, cccbKeyboard, pCancel);
			case TVC_BYMOUSE:
				return Fire_CaretChanging(pTvwPrevCaretItem, pTvwNewCaretItem, cccbMouse, pCancel);
			case TVC_UNKNOWN:
				return Fire_CaretChanging(pTvwPrevCaretItem, pTvwNewCaretItem, cccbUnknown, pCancel);
			case TVC_INTERNAL:
				return Fire_CaretChanging(pTvwPrevCaretItem, pTvwNewCaretItem, cccbInternal, pCancel);
			default:
				return Fire_CaretChanging(pTvwPrevCaretItem, pTvwNewCaretItem, cccbUnknown, pCancel);
		}
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ChangedSortOrder(SortOrderConstants previousSortOrder, SortOrderConstants newSortOrder)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangedSortOrder(previousSortOrder, newSortOrder);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ChangingSortOrder(SortOrderConstants previousSortOrder, SortOrderConstants newSortOrder, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangingSortOrder(previousSortOrder, newSortOrder, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.hLastClickedItem = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(mouseStatus.hLastClickedItem, this);
		return Fire_Click(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_CollapsedItem(ITreeViewItem* pTreeItem, VARIANT_BOOL removedAllSubItems)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CollapsedItem(pTreeItem, removedAllSubItems);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_CollapsingItem(ITreeViewItem* pTreeItem, VARIANT_BOOL removingAllSubItems, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CollapsingItem(pTreeItem, removingAllSubItems, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_CompareItems(ITreeViewItem* pFirstItem, ITreeViewItem* pSecondItem, CompareResultConstants* pResult)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CompareItems(pFirstItem, pSecondItem, pResult);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		#ifdef INCLUDESHELLBROWSERINTERFACE
			POINT menuPosition = {x, y};
		#endif
		if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
			if(properties.processContextMenuKeys) {
				// retrieve the caret item and propose its rectangle's middle as the menu's position
				HTREEITEM hItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, TVGN_CARET, 0));
				if(hItem) {
					CRect itemRectangle;
					*reinterpret_cast<HTREEITEM*>(&itemRectangle) = hItem;
					if(SendMessage(TVM_GETITEMRECT, TRUE, reinterpret_cast<LPARAM>(&itemRectangle))) {
						CPoint centerPoint = itemRectangle.CenterPoint();
						x = centerPoint.x;
						y = centerPoint.y;
					}
				}
			} else {
				return S_OK;
			}
		}

		UINT hitTestDetails = 0;
		HTREEITEM hItem = HitTest(x, y, &hitTestDetails);

		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItem, this);
		VARIANT_BOOL showDefaultMenu = VARIANT_TRUE;
		HRESULT hr = Fire_ContextMenu(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails), &showDefaultMenu);
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(SUCCEEDED(hr) && shellBrowserInterface.pInternalMessageListener) {
				if(showDefaultMenu != VARIANT_FALSE) {
					shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_CONTEXTMENU, 0, reinterpret_cast<LPARAM>(&menuPosition));
				}
			}
		#endif
		return hr;
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_CreatedEditControlWindow(LONG hWndEdit)
{
	// initialize some values that are needed during label-editing
	labelEditStatus.Reset(FALSE);
	hEditBackColorBrush = CreateSolidBrush(OLECOLOR2COLORREF(properties.editBackColor));
	properties.hWndEdit = static_cast<HWND>(LongToHandle(hWndEdit));

	// subclass the edit control
	containedEdit.SubclassWindow(properties.hWndEdit);

	//if(m_nFreezeEvents == 0) {
		return Fire_CreatedEditControlWindow(hWndEdit);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_CustomDraw(ITreeViewItem* pTreeItem, LONG* pItemLevel, OLE_COLOR* pTextColor, OLE_COLOR* pTextBackColor, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants itemState, LONG hDC, RECTANGLE* pDrawingRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CustomDraw(pTreeItem, pItemLevel, pTextColor, pTextBackColor, drawStage, itemState, hDC, pDrawingRectangle, pFurtherProcessing);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.raiseDblClkEvent = FALSE;
	UINT hitTestDetails = 0;
	HTREEITEM hItem = HitTest(x, y, &hitTestDetails);
	if(hItem != mouseStatus.hLastClickedItem) {
		hItem = NULL;
	}
	mouseStatus.hLastClickedItem = NULL;

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItem, this);
		return Fire_DblClick(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_DestroyedControlWindow(LONG hWnd)
{
	// clean up
	if(hBuiltInStateImageList) {
		/* We had a state imagelist. If it's still in use, we must destroy it.
		   NOTE: Never destroy a custom imagelist at this place. */
		if(hBuiltInStateImageList == reinterpret_cast<HIMAGELIST>(SendMessage(TVM_GETIMAGELIST, TVSIL_STATE, 0))) {
			ImageList_Destroy(hBuiltInStateImageList);
			if(cachedSettings.hStateImageList == hBuiltInStateImageList) {
				cachedSettings.hStateImageList = NULL;
			}
			hBuiltInStateImageList = NULL;
		}
	}

	KillTimer(timers.ID_DELAYEDCARETCHANGE);
	KillTimer(timers.ID_REDRAW);
	RemoveItemFromItemContainers(0);
	#ifdef USE_STL
		itemHandles.clear();
		itemParams.clear();
		itemDeletionsToIgnore.clear();
		internalCaretChanges.clear();
	#else
		itemHandles.RemoveAll();
		itemParams.RemoveAll();
		itemDeletionsToIgnore.RemoveAll();
		internalCaretChanges.RemoveAll();
	#endif

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
		//CoLockObjectExternal(static_cast<IDropTarget*>(this), FALSE, TRUE);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_DestroyedEditControlWindow(LONG hWndEdit)
{
	BOOL acceptText = labelEditStatus.acceptText;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		HTREEITEM hEditedItem = labelEditStatus.hEditedItem;
		TCHAR pPreviousItemText[MAX_ITEMTEXTLENGTH + 1];
		if(labelEditStatus.previousText) {
			COLE2T converter(labelEditStatus.previousText);
			lstrcpyn(pPreviousItemText, converter, MAX_ITEMTEXTLENGTH + 1);
		} else {
			pPreviousItemText[0] = TEXT('\0');
		}
	#endif

	if(labelEditStatus.hEditedItem) {
		if(acceptText) {
			CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(labelEditStatus.hEditedItem, this);
			BSTR itemText = NULL;
			pTvwItem->get_Text(&itemText);
			Raise_RenamedItem(pTvwItem, labelEditStatus.previousText, itemText);
			SysFreeString(itemText);
		}
		labelEditStatus.Reset();
	}

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		TRACKMOUSEEVENT trackingOptions = {0};
		trackingOptions.cbSize = sizeof(trackingOptions);
		trackingOptions.hwndTrack = containedEdit.m_hWnd;
		trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
		TrackMouseEvent(&trackingOptions);

		if(mouseStatus_Edit.enteredControl) {
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			Raise_EditMouseLeave(button, shift, mouseStatus_Edit.lastPosition.x, mouseStatus_Edit.lastPosition.y);
		}
	}

	containedEdit.UnsubclassWindow();
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		hr = Fire_DestroyedEditControlWindow(hWndEdit);
	//}

	// clean up
	DeleteObject(hEditBackColorBrush);
	hEditBackColorBrush = NULL;
	properties.hWndEdit = NULL;

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(!acceptText) {
			// the shell item could not be renamed, so go back into edit mode
			TVITEMEX item = {0};
			item.mask = TVIF_HANDLE | TVIF_TEXT;
			item.hItem = hEditedItem;
			item.pszText = pPreviousItemText;
			item.cchTextMax = lstrlen(item.pszText);
			SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
			PostMessage(TVM_EDITLABEL, 0, reinterpret_cast<LPARAM>(hEditedItem));
		}
	#endif

	return hr;
}

inline HRESULT ExplorerTreeView::Raise_DragMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(dragDropStatus.hDragImageList) {
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ImageList_DragMove(mousePosition.x, mousePosition.y);
	}

	UINT hitTestDetails = 0;
	HTREEITEM hDropTarget = HitTest(x, y, &hitTestDetails, TRUE);
	if(!hDropTarget) {
		if(hitTestDetails & htBelowLastItem) {
			hDropTarget = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_LASTVISIBLE), 0));
		}
	}
	ITreeViewItem* pDropTarget = NULL;
	ClassFactory::InitTreeItem(hDropTarget, this, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	LONG yToItemTop = 0;
	if(pDropTarget) {
		LONG yItem = 0;
		pDropTarget->GetRectangle(irtItem, NULL, &yItem);
		yToItemTop = y - yItem;
	}

	VARIANT_BOOL autoExpandItem = VARIANT_TRUE;
	if(!hDropTarget) {
		autoExpandItem = VARIANT_FALSE;
	}

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePosition(x, y);
		CRect noScrollZone(0, 0, 0, 0);
		GetClientRect(&noScrollZone);
		BOOL isInScrollZone = noScrollZone.PtInRect(mousePosition);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			isInScrollZone = !noScrollZone.PtInRect(mousePosition);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePosition.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePosition.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePosition.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePosition.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		hr = Fire_DragMouseMove(&pDropTarget, button, shift, x, y, yToItemTop, static_cast<HitTestConstants>(hitTestDetails), &autoExpandItem, &autoHScrollVelocity, &autoVScrollVelocity);
	//}

	if(pDropTarget) {
		OLE_HANDLE h = NULL;
		pDropTarget->get_Handle(&h);
		hDropTarget = static_cast<HTREEITEM>(LongToHandle(h));
		// we don't want to produce a mem-leak...
		pDropTarget->Release();
	} else {
		hDropTarget = NULL;
	}

	if(autoExpandItem == VARIANT_FALSE) {
		// cancel auto-expansion
		KillTimer(timers.ID_DRAGEXPAND);
	} else {
		if(hDropTarget != dragDropStatus.hLastDropTarget) {
			// cancel auto-expansion of previous target
			KillTimer(timers.ID_DRAGEXPAND);
			if(properties.dragExpandTime != 0) {
				// start timered auto-expansion of new target
				SetTimer(timers.ID_DRAGEXPAND, (properties.dragExpandTime == -1 ? GetDoubleClickTime() * 4 : properties.dragExpandTime));
			}
		}
	}
	dragDropStatus.hLastDropTarget = hDropTarget;

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT ExplorerTreeView::Raise_Drop(void)
{
	//if(m_nFreezeEvents == 0) {
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ScreenToClient(&mousePosition);

		UINT hitTestDetails = 0;
		HTREEITEM hDropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, TRUE);
		if(!hDropTarget) {
			if(hitTestDetails & htBelowLastItem) {
				hDropTarget = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_LASTVISIBLE), 0));
			}
		}
		CComPtr<ITreeViewItem> pDropTarget = ClassFactory::InitTreeItem(hDropTarget, this);

		LONG yToItemTop = 0;
		if(pDropTarget) {
			LONG yItem = 0;
			pDropTarget->GetRectangle(irtItem, NULL, &yItem);
			yToItemTop = mousePosition.y - yItem;
		}

		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		return Fire_Drop(pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditChange(void)
{
	if(!containedEdit.IsWindow()) {
		/* We ignore the event because the user would probably get/set the EditText property which
		   fails as long as the window isn't created. */
		return S_OK;
	}
	//if(m_nFreezeEvents == 0) {
		return Fire_EditChange();
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu)
{
	if((x == -1) && (y == -1)) {
		// the event was caused by the keyboard
		if(properties.processContextMenuKeys) {
			// propose the middle of the edit control's client rectangle as the menu's position
			CRect clientArea;
			containedEdit.GetClientRect(&clientArea);
			CPoint centerPoint = clientArea.CenterPoint();
			x = centerPoint.x;
			y = centerPoint.y;
		} else {
			return S_OK;
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_EditContextMenu(button, shift, x, y, pShowDefaultMenu);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditGotFocus(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditGotFocus();
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditKeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditKeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditKeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditKeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditKeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditKeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditLostFocus(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditLostFocus();
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditMClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditMDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditMouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus_Edit.enteredControl) {
		Raise_EditMouseEnter(button, shift, x, y);
	}
	if(!mouseStatus_Edit.hoveredControl) {
		TRACKMOUSEEVENT trackingOptions = {0};
		trackingOptions.cbSize = sizeof(trackingOptions);
		trackingOptions.hwndTrack = containedEdit.m_hWnd;
		trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
		TrackMouseEvent(&trackingOptions);

		Raise_EditMouseHover(button, shift, x, y);
	}
	mouseStatus_Edit.StoreClickCandidate(button);

	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseDown(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditMouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = containedEdit.m_hWnd;
	trackingOptions.dwHoverTime = (properties.editHoverTime == -1 ? HOVER_DEFAULT : properties.editHoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus_Edit.EnterControl();
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseEnter(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditMouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus_Edit.HoverControl();
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseHover(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditMouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus_Edit.LeaveControl();
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseLeave(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus_Edit.enteredControl) {
		Raise_EditMouseEnter(button, shift, x, y);
	}
	mouseStatus_Edit.lastPosition.x = x;
	mouseStatus_Edit.lastPosition.y = y;
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseMove(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditMouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(mouseStatus_Edit.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea;
		containedEdit.GetClientRect(&clientArea);
		containedEdit.ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != containedEdit.m_hWnd) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					Raise_EditClick(button, shift, x, y);
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					Raise_EditRClick(button, shift, x, y);
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					Raise_EditMClick(button, shift, x, y);
					break;
				case embXButton1:
				case embXButton2:
					Raise_EditXClick(button, shift, x, y);
					break;
			}
		}
		mouseStatus_Edit.RemoveClickCandidate(button);
	}
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseUp(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditMouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	if(!mouseStatus_Edit.enteredControl) {
		Raise_EditMouseEnter(button, shift, x, y);
	}
	mouseStatus_Edit.lastPosition.x = x;
	mouseStatus_Edit.lastPosition.y = y;
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseWheel(button, shift, x, y, scrollAxis, wheelDelta);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditRClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditRClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditRDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditRDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditXClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditXClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_EditXDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditXDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ExpandedItem(ITreeViewItem* pTreeItem, VARIANT_BOOL expandedPartially)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ExpandedItem(pTreeItem, expandedPartially);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ExpandingItem(ITreeViewItem* pTreeItem, VARIANT_BOOL expandingPartially, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ExpandingItem(pTreeItem, expandingPartially, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_FreeItemData(ITreeViewItem* pTreeItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FreeItemData(pTreeItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_IncrementalSearchStringChanging(BSTR currentSearchString, SHORT keyCodeOfCharToBeAdded, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_IncrementalSearchStringChanging(currentSearchString, keyCodeOfCharToBeAdded, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_InsertedItem(ITreeViewItem* pTreeItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedItem(pTreeItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_InsertingItem(IVirtualTreeViewItem* pTreeItem, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertingItem(pTreeItem, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ItemAsynchronousDrawFailed(ITreeViewItem* pTreeItem, NMTVASYNCDRAW* pNotificationDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemAsynchronousDrawFailed(pTreeItem, pNotificationDetails);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ItemBeginDrag(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemBeginDrag(pTreeItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ItemBeginRDrag(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemBeginRDrag(pTreeItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ItemGetDisplayInfo(ITreeViewItem* pTreeItem, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG* pSelectedIconIndex, LONG* pExpandedIconIndex, VARIANT_BOOL* pHasButton, LONG maxItemTextLength, BSTR* pItemText, VARIANT_BOOL* pDontAskAgain)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemGetDisplayInfo(pTreeItem, requestedInfo, pIconIndex, pSelectedIconIndex, pExpandedIconIndex, pHasButton, maxItemTextLength, pItemText, pDontAskAgain);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ItemGetInfoTipText(ITreeViewItem* pTreeItem, LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemGetInfoTipText(pTreeItem, maxInfoTipLength, pInfoTipText, pAbortToolTip);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ItemMouseEnter(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_ItemMouseEnter(pTreeItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ItemMouseLeave(ITreeViewItem* pTreeItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_ItemMouseLeave(pTreeItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ItemSelectionChanged(HTREEITEM hItem)
{
	//if(m_nFreezeEvents == 0) {
		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItem, this);
		return Fire_ItemSelectionChanged(pTvwItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ItemSelectionChanging(HTREEITEM hItem, VARIANT_BOOL* pCancelChange)
{
	//if(m_nFreezeEvents == 0) {
		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItem, this);
		return Fire_ItemSelectionChanging(pTvwItem, pCancelChange);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ItemSetText(ITreeViewItem* pTreeItem, BSTR itemText)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemSetText(pTreeItem, itemText);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ItemStateImageChanged(ITreeViewItem* pTreeItem, LONG previousStateImageIndex, LONG newStateImageIndex, StateImageChangeCausedByConstants causedBy)
{
	stateImageChangeFlags.reasonForPotentialStateImageChange = static_cast<StateImageChangeCausedByConstants>(-1);
	if(newStateImageIndex != previousStateImageIndex) {
		//if(m_nFreezeEvents == 0) {
			return Fire_ItemStateImageChanged(pTreeItem, previousStateImageIndex, newStateImageIndex, causedBy);
		//}
	}
	return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ItemStateImageChanging(ITreeViewItem* pTreeItem, LONG previousStateImageIndex, LONG* pNewStateImageIndex, StateImageChangeCausedByConstants causedBy, VARIANT_BOOL* pCancel)
{
	ATLASSERT(*pNewStateImageIndex != previousStateImageIndex);
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemStateImageChanging(pTreeItem, previousStateImageIndex, pNewStateImageIndex, causedBy, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.hLastClickedItem = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(mouseStatus.hLastClickedItem, this);
		return Fire_MClick(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	HTREEITEM hItem = HitTest(x, y, &hitTestDetails);
	if(hItem != mouseStatus.hLastClickedItem) {
		hItem = NULL;
	}
	mouseStatus.hLastClickedItem = NULL;

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItem, this);
		return Fire_MDblClick(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	if(!mouseStatus.hoveredControl) {
		TRACKMOUSEEVENT trackingOptions = {0};
		trackingOptions.cbSize = sizeof(trackingOptions);
		trackingOptions.hwndTrack = *this;
		trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
		TrackMouseEvent(&trackingOptions);

		Raise_MouseHover(button, shift, x, y);
	}
	mouseStatus.StoreClickCandidate(button);

	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		HTREEITEM hItem = HitTest(x, y, &hitTestDetails);

		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItem, this);
		return Fire_MouseDown(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = *this;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	UINT hitTestDetails = 0;
	HTREEITEM hItem = HitTest(x, y, &hitTestDetails);
	hItemUnderMouse = hItem;

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItem, this);
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		hr = Fire_MouseEnter(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	if(pTvwItem) {
		Raise_ItemMouseEnter(pTvwItem, button, shift, x, y, hitTestDetails);
	}
	return hr;
}

inline HRESULT ExplorerTreeView::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.HoverControl();

	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		HTREEITEM hItem = HitTest(x, y, &hitTestDetails);

		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItem, this);
		return Fire_MouseHover(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	HTREEITEM hItem = HitTest(x, y, &hitTestDetails);

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItemUnderMouse, this);
	if(pTvwItem) {
		Raise_ItemMouseLeave(pTvwItem, button, shift, x, y, hitTestDetails);
	}
	hItemUnderMouse = NULL;

	mouseStatus.LeaveControl();

	//if(m_nFreezeEvents == 0) {
		pTvwItem = ClassFactory::InitTreeItem(hItem, this);
		return Fire_MouseLeave(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	UINT hitTestDetails = 0;
	HTREEITEM hItem = HitTest(x, y, &hitTestDetails);

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItem, this);
	// Do we move over another item than before?
	if(hItem != hItemUnderMouse) {
		CComPtr<ITreeViewItem> pPrevTvwItem = ClassFactory::InitTreeItem(hItemUnderMouse, this);
		if(pPrevTvwItem) {
			Raise_ItemMouseLeave(pPrevTvwItem, button, shift, x, y, hitTestDetails);
		}
		hItemUnderMouse = hItem;
		if(pTvwItem) {
			Raise_ItemMouseEnter(pTvwItem, button, shift, x, y, hitTestDetails);
		}
	}
	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea;
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					/*if(!(properties.disabledEvents & deTreeClickEvents)) {
						Raise_Click(button, shift, x, y);
					}*/
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					/*if(!(properties.disabledEvents & deTreeClickEvents)) {
						Raise_RClick(button, shift, x, y);
					}*/
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deTreeClickEvents)) {
						Raise_MClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deTreeClickEvents)) {
						Raise_XClick(button, shift, x, y);
					}
					break;
			}
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		HTREEITEM hItem = HitTest(x, y, &hitTestDetails);
		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItem, this);
		return Fire_MouseUp(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}

	UINT hitTestDetails = 0;
	HTREEITEM hItem = HitTest(x, y, &hitTestDetails);

	CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItem, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseWheel(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails), scrollAxis, wheelDelta);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLECompleteDrag(pData, performedEffect);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_OLEDragDrop(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGSCROLL);
	KillTimer(timers.ID_DRAGEXPAND);

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	HTREEITEM hDropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, TRUE);
	if(!hDropTarget) {
		if(hitTestDetails & htBelowLastItem) {
			hDropTarget = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_LASTVISIBLE), 0));
		}
	}
	ITreeViewItem* pDropTarget = NULL;
	ClassFactory::InitTreeItem(hDropTarget, this, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	LONG yToItemTop = 0;
	if(pDropTarget) {
		LONG yItem = 0;
		pDropTarget->GetRectangle(irtItem, NULL, &yItem);
		yToItemTop = mousePosition.y - yItem;
	}

	if(pData) {
		/* Actually we wouldn't need the next line, because the data object passed to this method should
		   always be the same as the data object that was passed to Raise_OLEDragEnter. */
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	BOOL fireEvent = TRUE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			SHTVMDRAGDROPEVENTDATA eventDetails = {0};
			eventDetails.structSize = sizeof(SHTVMDRAGDROPEVENTDATA);
			eventDetails.pDataObject = reinterpret_cast<LPUNKNOWN>(dragDropStatus.pActiveDataObject.p);
			eventDetails.effect = *pEffect;
			eventDetails.pDropTarget = reinterpret_cast<LPUNKNOWN>(pDropTarget);
			eventDetails.hDropTarget = hDropTarget;
			eventDetails.keyState = keyState;
			eventDetails.draggingMouseButton = dragDropStatus.draggingMouseButton;
			eventDetails.cursorPosition = mousePosition;
			ClientToScreen(reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition));
			eventDetails.yToItemTop = yToItemTop;
			eventDetails.hitTestDetails = hitTestDetails;
			eventDetails.pDropTargetHelper = dragDropStatus.pDropTargetHelper;

			hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_HANDLEDRAGEVENT, OLEDRAGEVENT_DROP, reinterpret_cast<LPARAM>(&eventDetails));
			if(hr == S_OK) {
				fireEvent = FALSE;
				if(pCallDropTargetHelper) {
					*pCallDropTargetHelper = FALSE;
					dragDropStatus.drop_pDataObject = NULL;
				}

				*pEffect = eventDetails.effect;
				pDropTarget = reinterpret_cast<ITreeViewItem*>(eventDetails.pDropTarget);
			}
		}
	#endif

	//if(m_nFreezeEvents == 0) {
		if(fireEvent && dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	if(pDropTarget) {
		// we're using a raw pointer
		pDropTarget->Release();
	}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT ExplorerTreeView::Raise_OLEDragEnter(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	UINT hitTestDetails = 0;
	HTREEITEM hDropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	if(!hDropTarget) {
		if(hitTestDetails & htBelowLastItem) {
			hDropTarget = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_LASTVISIBLE), 0));
		}
	}
	ITreeViewItem* pDropTarget = NULL;
	ClassFactory::InitTreeItem(hDropTarget, this, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	LONG yToItemTop = 0;
	if(pDropTarget) {
		LONG yItem = 0;
		pDropTarget->GetRectangle(irtItem, NULL, &yItem);
		yToItemTop = mousePosition.y - yItem;
	}

	VARIANT_BOOL autoExpandItem = VARIANT_TRUE;
	if(!hDropTarget) {
		autoExpandItem = VARIANT_FALSE;
	}

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePos(mousePosition.x, mousePosition.y);
		CRect noScrollZone(0, 0, 0, 0);
		GetClientRect(&noScrollZone);
		BOOL isInScrollZone = noScrollZone.PtInRect(mousePos);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			isInScrollZone = !noScrollZone.PtInRect(mousePos);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePos.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePos.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePos.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePos.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	BOOL fireEvent = TRUE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			SHTVMDRAGDROPEVENTDATA eventDetails = {0};
			eventDetails.structSize = sizeof(SHTVMDRAGDROPEVENTDATA);
			eventDetails.pDataObject = reinterpret_cast<LPUNKNOWN>(dragDropStatus.pActiveDataObject.p);
			eventDetails.effect = *pEffect;
			eventDetails.pDropTarget = reinterpret_cast<LPUNKNOWN>(pDropTarget);
			eventDetails.hDropTarget = hDropTarget;
			eventDetails.keyState = keyState;
			eventDetails.draggingMouseButton = dragDropStatus.draggingMouseButton;
			eventDetails.cursorPosition = mousePosition;
			ClientToScreen(reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition));
			eventDetails.yToItemTop = yToItemTop;
			eventDetails.hitTestDetails = hitTestDetails;
			eventDetails.autoExpandItem = autoExpandItem;
			eventDetails.autoHScrollVelocity = autoHScrollVelocity;
			eventDetails.autoVScrollVelocity = autoVScrollVelocity;
			eventDetails.pDropTargetHelper = dragDropStatus.pDropTargetHelper;

			hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_HANDLEDRAGEVENT, OLEDRAGEVENT_DRAGENTER, reinterpret_cast<LPARAM>(&eventDetails));
			if(hr == S_OK) {
				fireEvent = FALSE;
				if(pCallDropTargetHelper) {
					*pCallDropTargetHelper = FALSE;
				}

				*pEffect = eventDetails.effect;
				pDropTarget = reinterpret_cast<ITreeViewItem*>(eventDetails.pDropTarget);
				autoExpandItem = eventDetails.autoExpandItem;
				autoHScrollVelocity = eventDetails.autoHScrollVelocity;
				autoVScrollVelocity = eventDetails.autoVScrollVelocity;
			}
		}
	#endif

	//if(m_nFreezeEvents == 0) {
		if(fireEvent && dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, static_cast<HitTestConstants>(hitTestDetails), &autoExpandItem, &autoHScrollVelocity, &autoVScrollVelocity);
		}
	//}

	if(pDropTarget) {
		OLE_HANDLE h = NULL;
		pDropTarget->get_Handle(&h);
		hDropTarget = static_cast<HTREEITEM>(LongToHandle(h));
		// we're using a raw pointer
		pDropTarget->Release();
	} else {
		hDropTarget = NULL;
	}

	if(autoExpandItem == VARIANT_FALSE) {
		// cancel auto-expansion
		KillTimer(timers.ID_DRAGEXPAND);
	} else {
		if(hDropTarget != dragDropStatus.hLastDropTarget) {
			// cancel auto-expansion of previous target
			KillTimer(timers.ID_DRAGEXPAND);
			if(properties.dragExpandTime != 0) {
				// start timered auto-expansion of new target
				SetTimer(timers.ID_DRAGEXPAND, (properties.dragExpandTime == -1 ? GetDoubleClickTime() * 4 : properties.dragExpandTime));
			}
		}
	}
	dragDropStatus.hLastDropTarget = hDropTarget;

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT ExplorerTreeView::Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragEnterPotentialTarget(hWndPotentialTarget);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_OLEDragLeave(BOOL* pCallDropTargetHelper)
{
	KillTimer(timers.ID_DRAGSCROLL);
	KillTimer(timers.ID_DRAGEXPAND);

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	UINT hitTestDetails = 0;
	HTREEITEM hDropTarget = HitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails);
	if(!hDropTarget) {
		if(hitTestDetails & htBelowLastItem) {
			hDropTarget = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_LASTVISIBLE), 0));
		}
	}
	CComPtr<ITreeViewItem> pDropTarget = ClassFactory::InitTreeItem(hDropTarget, this);

	LONG yToItemTop = 0;
	if(pDropTarget) {
		LONG yItem = 0;
		pDropTarget->GetRectangle(irtItem, NULL, &yItem);
		yToItemTop = dragDropStatus.lastMousePosition.y - yItem;
	}

	HRESULT hr = S_OK;
	BOOL fireEvent = TRUE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			SHTVMDRAGDROPEVENTDATA eventDetails = {0};
			eventDetails.structSize = sizeof(SHTVMDRAGDROPEVENTDATA);
			eventDetails.pDataObject = reinterpret_cast<LPUNKNOWN>(dragDropStatus.pActiveDataObject.p);
			eventDetails.pDropTarget = reinterpret_cast<LPUNKNOWN>(pDropTarget.p);
			eventDetails.hDropTarget = hDropTarget;
			BUTTONANDSHIFT2OLEKEYSTATE(button, shift, eventDetails.keyState);
			eventDetails.draggingMouseButton = dragDropStatus.draggingMouseButton;
			eventDetails.cursorPosition = dragDropStatus.lastMousePosition;
			ClientToScreen(reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition));
			eventDetails.yToItemTop = yToItemTop;
			eventDetails.hitTestDetails = hitTestDetails;
			eventDetails.pDropTargetHelper = dragDropStatus.pDropTargetHelper;

			hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_HANDLEDRAGEVENT, OLEDRAGEVENT_DRAGLEAVE, reinterpret_cast<LPARAM>(&eventDetails));
			if(hr == S_OK) {
				fireEvent = FALSE;
				if(pCallDropTargetHelper) {
					*pCallDropTargetHelper = FALSE;
				}
			}
		}
	#endif

	//if(m_nFreezeEvents == 0) {
		if(fireEvent && dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, pDropTarget, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, yToItemTop, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT ExplorerTreeView::Raise_OLEDragLeavePotentialTarget(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragLeavePotentialTarget();
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_OLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	HTREEITEM hDropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	if(!hDropTarget) {
		if(hitTestDetails & htBelowLastItem) {
			hDropTarget = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_LASTVISIBLE), 0));
		}
	}
	ITreeViewItem* pDropTarget = NULL;
	ClassFactory::InitTreeItem(hDropTarget, this, IID_ITreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	LONG yToItemTop = 0;
	if(pDropTarget) {
		LONG yItem = 0;
		pDropTarget->GetRectangle(irtItem, NULL, &yItem);
		yToItemTop = mousePosition.y - yItem;
	}

	VARIANT_BOOL autoExpandItem = VARIANT_TRUE;
	if(!hDropTarget) {
		autoExpandItem = VARIANT_FALSE;
	}

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePos(mousePosition.x, mousePosition.y);
		CRect noScrollZone(0, 0, 0, 0);
		GetClientRect(&noScrollZone);
		BOOL isInScrollZone = noScrollZone.PtInRect(mousePos);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			isInScrollZone = !noScrollZone.PtInRect(mousePos);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePos.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePos.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePos.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePos.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	HRESULT hr = S_OK;
	BOOL fireEvent = TRUE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			SHTVMDRAGDROPEVENTDATA eventDetails = {0};
			eventDetails.structSize = sizeof(SHTVMDRAGDROPEVENTDATA);
			eventDetails.pDataObject = reinterpret_cast<LPUNKNOWN>(dragDropStatus.pActiveDataObject.p);
			eventDetails.effect = *pEffect;
			eventDetails.pDropTarget = reinterpret_cast<LPUNKNOWN>(pDropTarget);
			eventDetails.hDropTarget = hDropTarget;
			eventDetails.keyState = keyState;
			eventDetails.draggingMouseButton = dragDropStatus.draggingMouseButton;
			eventDetails.cursorPosition = mousePosition;
			ClientToScreen(reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition));
			eventDetails.yToItemTop = yToItemTop;
			eventDetails.hitTestDetails = hitTestDetails;
			eventDetails.autoExpandItem = autoExpandItem;
			eventDetails.autoHScrollVelocity = autoHScrollVelocity;
			eventDetails.autoVScrollVelocity = autoVScrollVelocity;
			eventDetails.pDropTargetHelper = dragDropStatus.pDropTargetHelper;

			hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_HANDLEDRAGEVENT, OLEDRAGEVENT_DRAGOVER, reinterpret_cast<LPARAM>(&eventDetails));
			if(hr == S_OK) {
				fireEvent = FALSE;
				if(pCallDropTargetHelper) {
					*pCallDropTargetHelper = FALSE;
				}

				*pEffect = eventDetails.effect;
				pDropTarget = reinterpret_cast<ITreeViewItem*>(eventDetails.pDropTarget);
				autoExpandItem = eventDetails.autoExpandItem;
				autoHScrollVelocity = eventDetails.autoHScrollVelocity;
				autoVScrollVelocity = eventDetails.autoVScrollVelocity;
			}
		}
	#endif

	//if(m_nFreezeEvents == 0) {
		if(fireEvent && dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, static_cast<HitTestConstants>(hitTestDetails), &autoExpandItem, &autoHScrollVelocity, &autoVScrollVelocity);
		}
	//}

	if(pDropTarget) {
		OLE_HANDLE h = NULL;
		pDropTarget->get_Handle(&h);
		hDropTarget = static_cast<HTREEITEM>(LongToHandle(h));
		// we're using a raw pointer
		pDropTarget->Release();
	} else {
		hDropTarget = NULL;
	}

	if(autoExpandItem == VARIANT_FALSE) {
		// cancel auto-expansion
		KillTimer(timers.ID_DRAGEXPAND);
	} else {
		if(hDropTarget != dragDropStatus.hLastDropTarget) {
			// cancel auto-expansion of previous target
			KillTimer(timers.ID_DRAGEXPAND);
			if(properties.dragExpandTime != 0) {
				// start timered auto-expansion of new target
				SetTimer(timers.ID_DRAGEXPAND, (properties.dragExpandTime == -1 ? GetDoubleClickTime() * 4 : properties.dragExpandTime));
			}
		}
	}
	dragDropStatus.hLastDropTarget = hDropTarget;

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT ExplorerTreeView::Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEGiveFeedback(static_cast<OLEDropEffectConstants>(effect), pUseDefaultCursors);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith)
{
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);
		return Fire_OLEQueryContinueDrag(BOOL2VARIANTBOOL(pressedEscape), button, shift, reinterpret_cast<OLEActionToContinueWithConstants*>(pActionToContinueWith));
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT ExplorerTreeView::Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEReceivedNewData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT ExplorerTreeView::Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLESetData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_OLEStartDrag(IOLEDataObject* pData)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEStartDrag(pData);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.hLastClickedItem = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(mouseStatus.hLastClickedItem, this);
		return Fire_RClick(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.raiseDblClkEvent = FALSE;
	UINT hitTestDetails = 0;
	HTREEITEM hItem = HitTest(x, y, &hitTestDetails);
	if(hItem != mouseStatus.hLastClickedItem) {
		hItem = NULL;
	}
	mouseStatus.hLastClickedItem = NULL;

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItem, this);
		return Fire_RDblClick(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_RecreatedControlWindow(LONG hWnd)
{
	dragDropStatus.Reset();
	// configure the treeview control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop) {
		//CoLockObjectExternal(static_cast<IDropTarget*>(this), TRUE, FALSE);
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	/* Comctl32.dll older than version 6.0 seems to simply toggle the internal flag instead of checking the
	   wParam value. So it's a bad idea to call SetRedraw(TRUE) when it already is set to TRUE. */
	if(properties.dontRedraw) {
		SetRedraw(!properties.dontRedraw);
	}
	SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_RemovedItem(IVirtualTreeViewItem* pTreeItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovedItem(pTreeItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_RemovingItem(ITreeViewItem* pTreeItem, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingItem(pTreeItem, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_RenamedItem(ITreeViewItem* pTreeItem, BSTR previousItemText, BSTR newItemText)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RenamedItem(pTreeItem, previousItemText, newItemText);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_RenamingItem(ITreeViewItem* pTreeItem, BSTR previousItemText, BSTR newItemText, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RenamingItem(pTreeItem, previousItemText, newItemText, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_SingleExpanding(ITreeViewItem* pPreviousCaretItem, ITreeViewItem* pNewCaretItem, VARIANT_BOOL* pDontChangePreviousItem, VARIANT_BOOL* pDontChangeNewItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SingleExpanding(pPreviousCaretItem, pNewCaretItem, pDontChangePreviousItem, pDontChangeNewItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_StartingLabelEditing(ITreeViewItem* pTreeItem, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_StartingLabelEditing(pTreeItem, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.hLastClickedItem = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(mouseStatus.hLastClickedItem, this);
		return Fire_XClick(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerTreeView::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	HTREEITEM hItem = HitTest(x, y, &hitTestDetails);
	if(hItem != mouseStatus.hLastClickedItem) {
		hItem = NULL;
	}
	mouseStatus.hLastClickedItem = NULL;

	//if(m_nFreezeEvents == 0) {
		CComPtr<ITreeViewItem> pTvwItem = ClassFactory::InitTreeItem(hItem, this);
		return Fire_XDblClick(pTvwItem, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}


void ExplorerTreeView::AddCaretChangeTask(BOOL executeOnCaretChanging, BOOL executeOnCaretChanged, ICaretChangeTask* pTask)
{
	caretChangeTasks.AddTask(executeOnCaretChanging, executeOnCaretChanged, pTask);
}

void ExplorerTreeView::AddInternalCaretChange(HTREEITEM hNewCaretItem, UINT realCaretChangeCause/* = (UINT) -1*/)
{
	internalCaretChanges[hNewCaretItem] = realCaretChangeCause;
}

void ExplorerTreeView::RemoveInternalCaretChange(HTREEITEM hNewCaretItem)
{
	#ifdef USE_STL
		internalCaretChanges.erase(hNewCaretItem);
	#else
		internalCaretChanges.RemoveKey(hNewCaretItem);
	#endif
}

void ExplorerTreeView::EnterSilentItemInsertionSection(void)
{
	++flags.silentItemInsertions;
}

void ExplorerTreeView::LeaveSilentItemInsertionSection(void)
{
	--flags.silentItemInsertions;
}

void ExplorerTreeView::AddSilentItemDeletion(HTREEITEM hDeletedItem)
{
	#ifdef USE_STL
		std::vector<HTREEITEM>::iterator iter = std::find(itemDeletionsToIgnore.begin(), itemDeletionsToIgnore.end(), hDeletedItem);
		if(iter == itemDeletionsToIgnore.end()) {
			itemDeletionsToIgnore.push_back(hDeletedItem);
		}
	#else
		size_t iIgnoredItemDeletion = static_cast<size_t>(-1);
		for(size_t i = 0; i < itemDeletionsToIgnore.GetCount(); ++i) {
			if(itemDeletionsToIgnore[i] == hDeletedItem) {
				iIgnoredItemDeletion = i;
				break;
			}
		}
		if(iIgnoredItemDeletion == -1) {
			itemDeletionsToIgnore.Add(hDeletedItem);
		}
	#endif
}

void ExplorerTreeView::EnterNoSingleExpandSection(void)
{
	++flags.dontSingleExpand;
}

void ExplorerTreeView::LeaveNoSingleExpandSection(void)
{
	--flags.dontSingleExpand;
}


void ExplorerTreeView::RecreateControlWindow(void)
{
	if(m_bInPlaceActive) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}
}

IMEModeConstants ExplorerTreeView::GetCurrentIMEContextMode(HWND hWnd)
{
	IMEModeConstants imeContextMode = imeNoControl;

	IMEModeConstants countryTable[10];
	IMEFlags.GetIMECountryTable(countryTable);
	if(countryTable[0] == -10) {
		imeContextMode = imeInherit;
	} else {
		HIMC hIMC = ImmGetContext(hWnd);
		if(hIMC) {
			if(ImmGetOpenStatus(hIMC)) {
				DWORD conversionMode = 0;
				DWORD sentenceMode = 0;
				ImmGetConversionStatus(hIMC, &conversionMode, &sentenceMode);
				if(conversionMode & IME_CMODE_NATIVE) {
					if(conversionMode & IME_CMODE_KATAKANA) {
						if(conversionMode & IME_CMODE_FULLSHAPE) {
							imeContextMode = countryTable[imeKatakanaHalf];
						} else {
							imeContextMode = countryTable[imeAlphaFull];
						}
					} else {
						if(conversionMode & IME_CMODE_FULLSHAPE) {
							imeContextMode = countryTable[imeHiragana];
						} else {
							imeContextMode = countryTable[imeKatakana];
						}
					}
				} else {
					if(conversionMode & IME_CMODE_FULLSHAPE) {
						imeContextMode = countryTable[imeAlpha];
					} else {
						imeContextMode = countryTable[imeHangulFull];
					}
				}
			} else {
				imeContextMode = countryTable[imeDisable];
			}
			ImmReleaseContext(hWnd, hIMC);
		} else {
			imeContextMode = countryTable[imeOn];
		}
	}
	return imeContextMode;
}

void ExplorerTreeView::SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode)
{
	if((IMEMode == imeInherit) || (IMEMode == imeNoControl) || !::IsWindow(hWnd)) {
		return;
	}

	IMEModeConstants countryTable[10];
	IMEFlags.GetIMECountryTable(countryTable);
	if(countryTable[0] == -10) {
		return;
	}

	// update IME mode
	HIMC hIMC = ImmGetContext(hWnd);
	if(IMEMode == imeDisable) {
		// disable IME
		if(hIMC) {
			// close the IME
			if(ImmGetOpenStatus(hIMC)) {
				ImmSetOpenStatus(hIMC, FALSE);
			}
			// each ImmGetContext() needs a ImmReleaseContext()
			ImmReleaseContext(hWnd, hIMC);
			hIMC = NULL;
		}
		// remove the control's association to the IME context
		HIMC h = ImmAssociateContext(hWnd, NULL);
		if(h) {
			IMEFlags.hDefaultIMC = h;
		}
		return;
	} else {
		// enable IME
		if(!hIMC) {
			if(!IMEFlags.hDefaultIMC) {
				// create an IME context
				hIMC = ImmCreateContext();
				if(hIMC) {
					// associate the control with the IME context
					ImmAssociateContext(hWnd, hIMC);
				}
			} else {
				// associate the control with the default IME context
				ImmAssociateContext(hWnd, IMEFlags.hDefaultIMC);
			}
		} else {
			// each ImmGetContext() needs a ImmReleaseContext()
			ImmReleaseContext(hWnd, hIMC);
			hIMC = NULL;
		}
	}

	hIMC = ImmGetContext(hWnd);
	if(hIMC) {
		DWORD conversionMode = 0;
		DWORD sentenceMode = 0;
		switch(IMEMode) {
			case imeOn:
				// open IME
				ImmSetOpenStatus(hIMC, TRUE);
				break;
			case imeOff:
				// close IME
				ImmSetOpenStatus(hIMC, FALSE);
				break;
			default:
				// open IME
				ImmSetOpenStatus(hIMC, TRUE);
				ImmGetConversionStatus(hIMC, &conversionMode, &sentenceMode);
				// switch conversion
				switch(IMEMode) {
					case imeHiragana:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_KATAKANA;
						break;
					case imeKatakana:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_ALPHANUMERIC;
						break;
					case imeKatakanaHalf:
						conversionMode |= (IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_FULLSHAPE;
						break;
					case imeAlphaFull:
						conversionMode |= IME_CMODE_FULLSHAPE;
						conversionMode &= ~(IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						break;
					case imeAlpha:
						conversionMode |= IME_CMODE_ALPHANUMERIC;
						conversionMode &= ~(IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						break;
					case imeHangulFull:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_ALPHANUMERIC;
						break;
					case imeHangul:
						conversionMode |= IME_CMODE_NATIVE;
						conversionMode &= ~IME_CMODE_FULLSHAPE;
						break;
				}
				ImmSetConversionStatus(hIMC, conversionMode, sentenceMode);
				break;
		}
		// each ImmGetContext() needs a ImmReleaseContext()
		ImmReleaseContext(hWnd, hIMC);
		hIMC = NULL;
	}
}

IMEModeConstants ExplorerTreeView::GetEffectiveIMEMode(void)
{
	IMEModeConstants IMEMode = properties.IMEMode;
	if((IMEMode == imeInherit) && IsWindow()) {
		CWindow wnd = GetParent();
		while((IMEMode == imeInherit) && wnd.IsWindow()) {
			// retrieve the parent's IME mode
			IMEMode = GetCurrentIMEContextMode(wnd);
			wnd = wnd.GetParent();
		}
	}

	if(IMEMode == imeInherit) {
		// use imeNoControl as fallback
		IMEMode = imeNoControl;
	}
	return IMEMode;
}

IMEModeConstants ExplorerTreeView::GetEffectiveIMEMode_Edit(void)
{
	IMEModeConstants IMEMode = properties.editIMEMode;
	if(IMEMode == imeInherit) {
		// retrieve the parent's IME mode
		IMEMode = GetEffectiveIMEMode();
	}

	if((IMEMode == imeInherit) && IsWindow()) {
		CWindow wnd = GetParent();
		while((IMEMode == imeInherit) && wnd.IsWindow()) {
			// retrieve the parent's IME mode
			IMEMode = GetCurrentIMEContextMode(wnd);
			wnd = wnd.GetParent();
		}
	}

	if(IMEMode == imeInherit) {
		// use imeNoControl as fallback
		IMEMode = imeNoControl;
	}
	return IMEMode;
}

DWORD ExplorerTreeView::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING | WS_EX_RIGHTSCROLLBAR;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeft & rtlLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	if(properties.rightToLeft & rtlText) {
		extendedStyle |= WS_EX_RTLREADING;
	}
	return extendedStyle;
}

DWORD ExplorerTreeView::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}

	style |= TVS_NONEVENHEIGHT;
	if(!properties.allowDragDrop) {
		style |= TVS_DISABLEDRAGDROP;
	}
	if(properties.allowLabelEditing) {
		style |= TVS_EDITLABELS;
	}
	if(properties.alwaysShowSelection) {
		style |= TVS_SHOWSELALWAYS;
	}
	if(!(properties.disabledEvents & deItemGetInfoTipText)) {
		style |= TVS_INFOTIP;
	}
	if(properties.fullRowSelect) {
		style |= TVS_FULLROWSELECT;
	}
	switch(properties.hotTracking) {
		case htrNormal:
		case htrWinXPStyle:
			style |= TVS_TRACKSELECT;
			break;
	}
	switch(properties.lineStyle) {
		case lsLinesAtRoot:
			style |= TVS_LINESATROOT;
			break;
	}
	if(properties.rightToLeft & rtlText) {
		style |= TVS_RTLREADING;
	}
	switch(properties.scrollBars) {
		case sbNone:
			style |= TVS_NOSCROLL;
			break;
		case sbVerticalOnly:
			style |= TVS_NOHSCROLL;
			break;
	}
	if(properties.showStateImages) {
		if(IsComctl32Version610OrNewer()) {
			if(properties.builtInStateImages == bisiSelectedAndUnselected) {
				style |= TVS_CHECKBOXES;
			}
		} else {
			style |= TVS_CHECKBOXES;
		}
	}
	if(!properties.showToolTips) {
		style |= TVS_NOTOOLTIPS;
	}
	if(properties.treeViewStyle & tvsExpandos) {
		style |= TVS_HASBUTTONS;
	}
	if(properties.treeViewStyle & tvsLines) {
		style |= TVS_HASLINES;
	}
	switch(properties.singleExpand) {
		case seNormal:
		case seWinXPStyle:
			style |= TVS_SINGLEEXPAND;
			break;
	}

	return style;
}

void ExplorerTreeView::SendConfigurationMessages(void)
{
	if(IsComctl32Version610OrNewer()) {
		// TVS_EX_DOUBLEBUFFER eliminates flicker
		DWORD extendedStyle = TVS_EX_DOUBLEBUFFER;
		if(properties.autoHScroll) {
			extendedStyle |= TVS_EX_AUTOHSCROLL;
		}
		if(properties.showStateImages) {
			if(properties.builtInStateImages & bisiIndeterminate) {
				extendedStyle |= TVS_EX_PARTIALCHECKBOXES;
			}
			if(properties.builtInStateImages & bisiSelectedDimmed) {
				extendedStyle |= TVS_EX_DIMMEDCHECKBOXES;
			}
			if(properties.builtInStateImages & bisiExcluded) {
				extendedStyle |= TVS_EX_EXCLUSIONCHECKBOXES;
			}
		}
		if(properties.drawImagesAsynchronously) {
			extendedStyle |= TVS_EX_DRAWIMAGEASYNC;
		}
		if(properties.fadeExpandos) {
			extendedStyle |= TVS_EX_FADEINOUTEXPANDOS;
		}
		if(!properties.indentStateImages) {
			extendedStyle |= TVS_EX_NOINDENTSTATE;
		}
		if(properties.richToolTips) {
			extendedStyle |= TVS_EX_RICHTOOLTIP;
		}
		SendMessage(TVM_SETEXTENDEDSTYLE, 0, extendedStyle);
		SendMessage(TVM_SETAUTOSCROLLINFO, properties.autoHScrollPixelsPerSecond, properties.autoHScrollRedrawInterval);
	}

	SendMessage(TVM_SETBORDER, TVSBF_XBORDER | TVSBF_YBORDER, properties.itemBorders);
	SendMessage(TVM_SETBKCOLOR, 0, OLECOLOR2COLORREF(properties.backColor));
	SendMessage(TVM_SETTEXTCOLOR, 0, OLECOLOR2COLORREF(properties.foreColor));
	SendMessage(TVM_SETINDENT, properties.indent, 0);
	SendMessage(TVM_SETINSERTMARKCOLOR, 0, OLECOLOR2COLORREF(properties.insertMarkColor));
	SendMessage(TVM_SETITEMHEIGHT, properties.itemHeight, 0);
	SendMessage(TVM_SETLINECOLOR, 0, OLECOLOR2COLORREF(properties.lineColor));
	SendMessage(TVM_SETSCROLLTIME, properties.maxScrollTime, 0);

	ApplyFont();

	if(IsInDesignMode()) {
		// insert some dummy items
		TVINSERTSTRUCT insertionData = {0};
		insertionData.itemex.mask = TVIF_CHILDREN | TVIF_TEXT;
		insertionData.hInsertAfter = TVI_LAST;

		insertionData.hParent = TVI_ROOT;
		insertionData.itemex.cChildren = 1;
		insertionData.itemex.pszText = TEXT("Dummy Root Item 1");
		insertionData.itemex.cchTextMax = lstrlen(insertionData.itemex.pszText);
		HTREEITEM hRootItem1 = reinterpret_cast<HTREEITEM>(SendMessage(TVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData)));
		insertionData.itemex.cChildren = 0;
		insertionData.itemex.pszText = TEXT("Dummy Root Item 2");
		insertionData.itemex.cchTextMax = lstrlen(insertionData.itemex.pszText);
		SendMessage(TVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData));

		insertionData.hParent = hRootItem1;
		insertionData.itemex.cChildren = 1;
		insertionData.itemex.pszText = TEXT("Dummy Item 1");
		insertionData.itemex.cchTextMax = lstrlen(insertionData.itemex.pszText);
		HTREEITEM hItem1 = reinterpret_cast<HTREEITEM>(SendMessage(TVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData)));
		insertionData.itemex.pszText = TEXT("Dummy Item 2");
		insertionData.itemex.cchTextMax = lstrlen(insertionData.itemex.pszText);
		HTREEITEM hItem2 = reinterpret_cast<HTREEITEM>(SendMessage(TVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData)));
		insertionData.itemex.cChildren = 0;
		insertionData.itemex.pszText = TEXT("Dummy Item 3");
		insertionData.itemex.cchTextMax = lstrlen(insertionData.itemex.pszText);
		SendMessage(TVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData));

		insertionData.hParent = hItem1;
		insertionData.itemex.pszText = TEXT("Dummy Sub Item 1");
		insertionData.itemex.cchTextMax = lstrlen(insertionData.itemex.pszText);
		SendMessage(TVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData));
		insertionData.itemex.pszText = TEXT("Dummy Sub Item 2");
		insertionData.itemex.cchTextMax = lstrlen(insertionData.itemex.pszText);
		SendMessage(TVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData));

		insertionData.hParent = hItem2;
		insertionData.itemex.pszText = TEXT("Dummy Sub Item 3");
		insertionData.itemex.cchTextMax = lstrlen(insertionData.itemex.pszText);
		SendMessage(TVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData));

		SendMessage(TVM_EXPAND, TVE_EXPAND, reinterpret_cast<LPARAM>(hRootItem1));
		SendMessage(TVM_EXPAND, TVE_EXPAND, reinterpret_cast<LPARAM>(hItem1));
		SendMessage(TVM_EXPAND, TVE_EXPAND, reinterpret_cast<LPARAM>(hItem2));
	}
}

HCURSOR ExplorerTreeView::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

#ifdef USE_STL
	HTREEITEM ExplorerTreeView::IDToItemHandle(std::unordered_map<LONG, HTREEITEM>::key_type ID)
	{
		std::unordered_map<LONG, HTREEITEM>::iterator iter = itemHandles.find(ID);
		if(iter != itemHandles.end()) {
			return iter->second;
		}
		return NULL;
	}
#else
	HTREEITEM ExplorerTreeView::IDToItemHandle(CAtlMap<LONG, HTREEITEM>::KINARGTYPE ID)
	{
		CAtlMap<LONG, HTREEITEM>::CPair* pEntry = itemHandles.Lookup(ID);
		if(pEntry) {
			return pEntry->m_value;
		}
		return NULL;
	}
#endif

LONG ExplorerTreeView::ItemHandleToID(HTREEITEM hItem)
{
	ATLASSERT(IsWindow());

	TVITEMEX item = {0};
	item.hItem = hItem;
	item.mask = TVIF_HANDLE | TVIF_PARAM | TVIF_NOINTERCEPTION;
	if(SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return static_cast<LONG>(item.lParam);
	}
	return 0;
}

HTREEITEM ExplorerTreeView::HitTest(LONG x, LONG y, UINT* pFlags, BOOL ignoreBoundingBoxDefinition/* = FALSE*/)
{
	ATLASSERT(IsWindow());

	UINT hitTestFlags = 0;
	if(pFlags) {
		hitTestFlags = *pFlags;
	}
	TVHITTESTINFO hitTestInfo = {{x, y}, hitTestFlags, 0};
	HTREEITEM hItem = reinterpret_cast<HTREEITEM>(SendMessage(TVM_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo)));
	hitTestFlags = hitTestInfo.flags;
	if(pFlags) {
		*pFlags = hitTestFlags;
	}
	if(!ignoreBoundingBoxDefinition && ((properties.itemBoundingBoxDefinition & hitTestFlags) != hitTestFlags)) {
		hItem = NULL;
	}
	return hItem;
}

BOOL ExplorerTreeView::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}

void ExplorerTreeView::ReplaceItemHandle(HTREEITEM oldItemHandle, HTREEITEM newItemHandle)
{
	// TODO: find a more performant solution
	LONG itemID = ItemHandleToID(oldItemHandle);
	if(itemID) {
		// replace the item handle
		itemHandles[itemID] = newItemHandle;
	}

	#ifdef USE_STL
		std::unordered_map<HTREEITEM, ItemData>::iterator iter = itemParams.find(oldItemHandle);
		if(iter != itemParams.end()) {
			// re-insert the entry using the new key
			ItemData tmp = iter->second;
			itemParams.erase(iter);
			itemParams[newItemHandle] = tmp;
		}
	#else
		CAtlMap<HTREEITEM, ItemData>::CPair* pEntry = itemParams.Lookup(oldItemHandle);
		if(pEntry) {
			// re-insert the entry using the new key
			ItemData tmp = pEntry->m_value;
			itemParams.RemoveKey(oldItemHandle);
			itemParams[newItemHandle] = tmp;
		}
	#endif

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHTVM_UPDATEDITEMHANDLE, reinterpret_cast<WPARAM>(oldItemHandle), reinterpret_cast<LPARAM>(newItemHandle));
		}
	#endif
}

void ExplorerTreeView::AutoScroll(void)
{
	LONG realScrollTimeBase = properties.dragScrollTimeBase;
	if(realScrollTimeBase == -1) {
		realScrollTimeBase = GetDoubleClickTime() / 4;
	}

	if((dragDropStatus.autoScrolling.currentHScrollVelocity != 0) && ((GetStyle() & WS_HSCROLL) == WS_HSCROLL)) {
		if(dragDropStatus.autoScrolling.currentHScrollVelocity < 0) {
			// Have we been waiting long enough since the last scroll to the left?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Left) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE;
				GetScrollInfo(SB_HORZ, &scrollInfo);
				if(scrollInfo.nPos > scrollInfo.nMin) {
					// scroll left
					dragDropStatus.autoScrolling.lastScroll_Left = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_HSCROLL, SB_LINELEFT, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		} else {
			// Have we been waiting long enough since the last scroll to the right?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Right) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE;
				GetScrollInfo(SB_HORZ, &scrollInfo);
				if(scrollInfo.nPage) {
					scrollInfo.nMax -= scrollInfo.nPage - 1;
				}
				if(scrollInfo.nPos < scrollInfo.nMax) {
					// scroll right
					dragDropStatus.autoScrolling.lastScroll_Right = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_HSCROLL, SB_LINERIGHT, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		}
	}

	if((dragDropStatus.autoScrolling.currentVScrollVelocity != 0) && ((GetStyle() & WS_VSCROLL) == WS_VSCROLL)) {
		if(dragDropStatus.autoScrolling.currentVScrollVelocity < 0) {
			// Have we been waiting long enough since the last scroll upwardly?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Up) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentVScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE | SIF_PAGE;
				GetScrollInfo(SB_VERT, &scrollInfo);
				if(scrollInfo.nPos > scrollInfo.nMin) {
					// scroll up
					dragDropStatus.autoScrolling.lastScroll_Up = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_VSCROLL, SB_LINEUP, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		} else {
			// Have we been waiting long enough since the last scroll downwards?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Down) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentVScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE | SIF_PAGE;
				GetScrollInfo(SB_VERT, &scrollInfo);
				if(scrollInfo.nPage) {
					scrollInfo.nMax -= scrollInfo.nPage - 1;
				}
				if(scrollInfo.nPos < scrollInfo.nMax) {
					// scroll down
					dragDropStatus.autoScrolling.lastScroll_Down = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_VSCROLL, SB_LINEDOWN, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		}
	}
}

BOOL ExplorerTreeView::HasSubItems(HTREEITEM hItem)
{
	ATLASSERT(IsWindow());

	if((hItem != NULL) && !ISPREDEFINEDHTREEITEM(hItem)) {
		HTREEITEM hChild = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(hItem)));
		return (hChild != NULL);
	} else if(hItem == TVI_ROOT) {
		HTREEITEM hChild = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_ROOT), 0));
		return (hChild != NULL);
	}
	return FALSE;
}

BOOL ExplorerTreeView::HaveSameParentItem(HTREEITEM hItem1, HTREEITEM hItem2)
{
	ATLASSERT(IsWindow());

	HTREEITEM hParent1 = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hItem1)));
	HTREEITEM hParent2 = reinterpret_cast<HTREEITEM>(SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hItem2)));
	return (hParent1 == hParent2);
}

BOOL ExplorerTreeView::IsExpanded(HTREEITEM hItem, BOOL allowPartialExpansion/* = TRUE*/)
{
	ATLASSERT(IsWindow());

	TVITEMEX item = {0};
	item.hItem = hItem;
	item.stateMask = TVIS_EXPANDED | TVIS_EXPANDPARTIAL;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		if(allowPartialExpansion) {
			return (item.state & (TVIS_EXPANDED | TVIS_EXPANDPARTIAL));
		} else {
			return (item.state & TVIS_EXPANDED);
		}
	}
	return FALSE;
}

BOOL ExplorerTreeView::IsLeftMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	}
}

BOOL ExplorerTreeView::IsRightMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	}
}

LONG ExplorerTreeView::GetNewItemID(void)
{
	static LONG lastID = 0;

	return ++lastID;
}

void ExplorerTreeView::RegisterItemContainer(IItemContainer* pContainer)
{
	ATLASSUME(pContainer);
	#ifdef _DEBUG
		#ifdef USE_STL
			std::unordered_map<DWORD, IItemContainer*>::iterator iter = itemContainers.find(pContainer->GetID());
			ATLASSERT(iter == itemContainers.end());
		#else
			CAtlMap<DWORD, IItemContainer*>::CPair* pEntry = itemContainers.Lookup(pContainer->GetID());
			ATLASSERT(!pEntry);
		#endif
	#endif
	itemContainers[pContainer->GetID()] = pContainer;
}

void ExplorerTreeView::DeregisterItemContainer(DWORD containerID)
{
	#ifdef USE_STL
		std::unordered_map<DWORD, IItemContainer*>::iterator iter = itemContainers.find(containerID);
		ATLASSERT(iter != itemContainers.end());
		if(iter != itemContainers.end()) {
			itemContainers.erase(iter);
		}
	#else
		itemContainers.RemoveKey(containerID);
	#endif
}

void ExplorerTreeView::RemoveItemFromItemContainers(LONG itemID)
{
	#ifdef USE_STL
		for(std::unordered_map<DWORD, IItemContainer*>::const_iterator iter = itemContainers.begin(); iter != itemContainers.end(); ++iter) {
			iter->second->RemovedItem(itemID);
		}
	#else
		POSITION p = itemContainers.GetStartPosition();
		while(p) {
			itemContainers.GetValueAt(p)->RemovedItem(itemID);
			itemContainers.GetNextValue(p);
		}
	#endif
}

BOOL ExplorerTreeView::ValidateItemHandle(HTREEITEM itemHandle)
{
	ATLASSERT(IsWindow());

	TVITEMEX item = {0};
	item.hItem = itemHandle;
	item.mask = TVIF_HANDLE;
	return static_cast<BOOL>(DefWindowProc(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item)));
}


HRESULT ExplorerTreeView::CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing)
{
	if(!hIcon || !pBits || !pWICImagingFactory) {
		return E_FAIL;
	}

	ICONINFO iconInfo;
	GetIconInfo(hIcon, &iconInfo);
	ATLASSERT(iconInfo.hbmColor);
	BITMAP bitmapInfo = {0};
	if(iconInfo.hbmColor) {
		GetObject(iconInfo.hbmColor, sizeof(BITMAP), &bitmapInfo);
	} else if(iconInfo.hbmMask) {
		GetObject(iconInfo.hbmMask, sizeof(BITMAP), &bitmapInfo);
	}
	bitmapInfo.bmHeight = abs(bitmapInfo.bmHeight);
	BOOL needsFrame = ((bitmapInfo.bmWidth < size.cx) || (bitmapInfo.bmHeight < size.cy));
	if(iconInfo.hbmColor) {
		DeleteObject(iconInfo.hbmColor);
	}
	if(iconInfo.hbmMask) {
		DeleteObject(iconInfo.hbmMask);
	}

	HRESULT hr = E_FAIL;

	CComPtr<IWICBitmapScaler> pWICBitmapScaler = NULL;
	if(!needsFrame) {
		hr = pWICImagingFactory->CreateBitmapScaler(&pWICBitmapScaler);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapScaler);
	}
	if(needsFrame || SUCCEEDED(hr)) {
		CComPtr<IWICBitmap> pWICBitmapSource = NULL;
		hr = pWICImagingFactory->CreateBitmapFromHICON(hIcon, &pWICBitmapSource);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapSource);
		if(SUCCEEDED(hr)) {
			if(!needsFrame) {
				hr = pWICBitmapScaler->Initialize(pWICBitmapSource, size.cx, size.cy, WICBitmapInterpolationModeFant);
			}
			if(SUCCEEDED(hr)) {
				WICRect rc = {0};
				if(needsFrame) {
					rc.Height = bitmapInfo.bmHeight;
					rc.Width = bitmapInfo.bmWidth;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					LPRGBQUAD pIconBits = reinterpret_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, rc.Width * rc.Height * sizeof(RGBQUAD)));
					hr = pWICBitmapSource->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pIconBits));
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						// center the icon
						int xIconStart = (size.cx - bitmapInfo.bmWidth) / 2;
						int yIconStart = (size.cy - bitmapInfo.bmHeight) / 2;
						LPRGBQUAD pIconPixel = pIconBits;
						LPRGBQUAD pPixel = pBits;
						pPixel += yIconStart * size.cx;
						for(int y = yIconStart; y < yIconStart + bitmapInfo.bmHeight; ++y, pPixel += size.cx, pIconPixel += bitmapInfo.bmWidth) {
							CopyMemory(pPixel + xIconStart, pIconPixel, bitmapInfo.bmWidth * sizeof(RGBQUAD));
						}
						HeapFree(GetProcessHeap(), 0, pIconBits);

						rc.Height = size.cy;
						rc.Width = size.cx;

						// TODO: now draw a frame around it
					}
				} else {
					rc.Height = size.cy;
					rc.Width = size.cx;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					hr = pWICBitmapScaler->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pBits));
					ATLASSERT(SUCCEEDED(hr));

					if(SUCCEEDED(hr) && doAlphaChannelPostProcessing) {
						for(int i = 0; i < rc.Width * rc.Height; ++i, ++pBits) {
							if(pBits->rgbReserved == 0x00) {
								ZeroMemory(pBits, sizeof(RGBQUAD));
							}
						}
					}
				}
			} else {
				ATLASSERT(FALSE && "Bitmap scaler failed");
			}
		}
	}
	return hr;
}


BOOL ExplorerTreeView::IsComctl32Version610OrNewer(void)
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hr = ATL::AtlGetCommCtrlVersion(&major, &minor);
	if(SUCCEEDED(hr)) {
		return (((major == 6) && (minor >= 10)) || (major > 6));
	}
	return FALSE;
}