// DispIDs.h: Defines a DISPID for each COM property and method we're providing.

// IExplorerTreeView
// properties
#define DISPID_EXTVW_ALLOWDRAGDROP											1
#define DISPID_EXTVW_ALLOWLABELEDITING									2
#define DISPID_EXTVW_ALWAYSSHOWSELECTION								3
#define DISPID_EXTVW_APPEARANCE													5
#define DISPID_EXTVW_BACKCOLOR													6
#define DISPID_EXTVW_BLENDSELECTEDITEMSICONS						13
#define DISPID_EXTVW_BORDERSTYLE												14
#define DISPID_EXTVW_CARETCHANGEDDELAYTIME							15
#define DISPID_EXTVW_CARETITEM													DISPID_VALUE
#define DISPID_EXTVW_DISABLEDEVENTS											17
#define DISPID_EXTVW_DONTREDRAW													18
#define DISPID_EXTVW_DRAGEXPANDTIME											19
#define DISPID_EXTVW_DRAGGEDITEMS												20
#define DISPID_EXTVW_DRAGSCROLLTIMEBASE									21
#define DISPID_EXTVW_DROPHILITEDITEM										22
#define DISPID_EXTVW_EDITBACKCOLOR											23
#define DISPID_EXTVW_EDITFORECOLOR											24
#define DISPID_EXTVW_EDITHOVERTIME											25
#define DISPID_EXTVW_EDITIMEMODE												26
#define DISPID_EXTVW_EDITTEXT														27
#define DISPID_EXTVW_ENABLED														DISPID_ENABLED
#define DISPID_EXTVW_FAVORITESSTYLE											28
#define DISPID_EXTVW_FIRSTROOTITEM											29
#define DISPID_EXTVW_FIRSTVISIBLEITEM										30
#define DISPID_EXTVW_FONT																31
#define DISPID_EXTVW_FORECOLOR													32
#define DISPID_EXTVW_FULLROWSELECT											33
#define DISPID_EXTVW_GROUPBOXCOLOR											34
#define DISPID_EXTVW_HDRAGIMAGELIST											35
#define DISPID_EXTVW_HIMAGELIST													36
#define DISPID_EXTVW_HOTTRACKING												37
#define DISPID_EXTVW_HOVERTIME													38
#define DISPID_EXTVW_HWND																39
#define DISPID_EXTVW_HWNDEDIT														40
#define DISPID_EXTVW_HWNDTOOLTIP												41
#define DISPID_EXTVW_IMEMODE														42
#define DISPID_EXTVW_INCREMENTALSEARCHSTRING						43
#define DISPID_EXTVW_INDENT															44
#define DISPID_EXTVW_INSERTMARKCOLOR										45
#define DISPID_EXTVW_ITEMBOUNDINGBOXDEFINITION					46
#define DISPID_EXTVW_ITEMHEIGHT													47
#define DISPID_EXTVW_ITEMXBORDER												48
#define DISPID_EXTVW_ITEMYBORDER												49
#define DISPID_EXTVW_LASTROOTITEM												50
#define DISPID_EXTVW_LASTVISIBLEITEM										51
#define DISPID_EXTVW_LINECOLOR													52
#define DISPID_EXTVW_LINESTYLE													53
#define DISPID_EXTVW_MAXSCROLLTIME											54
#define DISPID_EXTVW_MOUSEICON													55
#define DISPID_EXTVW_MOUSEPOINTER												56
#define DISPID_EXTVW_PROCESSCONTEXTMENUKEYS							57
#define DISPID_EXTVW_REGISTERFOROLEDRAGDROP							58
#define DISPID_EXTVW_RIGHTTOLEFT												59
#define DISPID_EXTVW_ROOTITEMS													60
#define DISPID_EXTVW_SCROLLBARS													61
#define DISPID_EXTVW_SHOWDRAGIMAGE											62
#define DISPID_EXTVW_SHOWSTATEIMAGES										63
#define DISPID_EXTVW_SHOWTOOLTIPS												64
#define DISPID_EXTVW_SINGLEEXPAND												65
#define DISPID_EXTVW_SORTORDER													66
#define DISPID_EXTVW_SUPPORTOLEDRAGIMAGES								67
#define DISPID_EXTVW_TREEITEMS													68
#define DISPID_EXTVW_TREEVIEWSTYLE											69
#define DISPID_EXTVW_USESYSTEMFONT											70
#define DISPID_EXTVW_VERSION														71
#define DISPID_EXTVW_AUTOHSCROLL												84
#define DISPID_EXTVW_AUTOHSCROLLPIXELSPERSECOND					85
#define DISPID_EXTVW_AUTOHSCROLLREDRAWINTERVAL					86
#define DISPID_EXTVW_BUILTINSTATEIMAGES									87
#define DISPID_EXTVW_DRAWIMAGESASYNCHRONOUSLY						88
#define DISPID_EXTVW_FADEEXPANDOS												89
#define DISPID_EXTVW_INDENTSTATEIMAGES									90
#define DISPID_EXTVW_RICHTOOLTIPS												91
#define DISPID_EXTVW_OLEDRAGIMAGESTYLE									92
#define DISPID_EXTVW_LOCALE															93
#define DISPID_EXTVW_TEXTPARSINGFLAGS										94
// fingerprint
#define DISPID_EXTVW_APPID															500
#define DISPID_EXTVW_APPNAME														501
#define DISPID_EXTVW_APPSHORTNAME												502
#define DISPID_EXTVW_BUILD															503
#define DISPID_EXTVW_CHARSET														504
#define DISPID_EXTVW_ISRELEASE													505
#define DISPID_EXTVW_PROGRAMMER													506
#define DISPID_EXTVW_TESTER															507
// methods
#define DISPID_EXTVW_ABOUT															DISPID_ABOUTBOX
#define DISPID_EXTVW_BEGINDRAG													72
#define DISPID_EXTVW_COUNTVISIBLE												73
#define DISPID_EXTVW_CREATEITEMCONTAINER								74
#define DISPID_EXTVW_ENDDRAG														75
#define DISPID_EXTVW_ENDLABELEDIT												76
#define DISPID_EXTVW_GETINSERTMARKPOSITION							77
#define DISPID_EXTVW_HITTEST														78
#define DISPID_EXTVW_LOADSETTINGSFROMFILE								79
#define DISPID_EXTVW_OLEDRAG														80
#define DISPID_EXTVW_REFRESH														DISPID_REFRESH
#define DISPID_EXTVW_SAVESETTINGSTOFILE									81
#define DISPID_EXTVW_SETINSERTMARKPOSITION							82
#define DISPID_EXTVW_SORTITEMS													83
#define DISPID_EXTVW_FINISHOLEDRAGDROP									95


// _IExplorerTreeViewEvents
// methods
#define DISPID_EXTVWE_ABORTEDDRAG												1
#define DISPID_EXTVWE_CARETCHANGED											DISPID_VALUE
#define DISPID_EXTVWE_CARETCHANGING											4
#define DISPID_EXTVWE_CLICK															5
#define DISPID_EXTVWE_COLLAPSEDITEM											6
#define DISPID_EXTVWE_COLLAPSINGITEM										7
#define DISPID_EXTVWE_COMPAREITEMS											8
#define DISPID_EXTVWE_CONTEXTMENU												9
#define DISPID_EXTVWE_CREATEDEDITCONTROLWINDOW					10
#define DISPID_EXTVWE_CUSTOMDRAW												11
#define DISPID_EXTVWE_DBLCLICK													12
#define DISPID_EXTVWE_DESTROYEDCONTROLWINDOW						13
#define DISPID_EXTVWE_DESTROYEDEDITCONTROLWINDOW				14
#define DISPID_EXTVWE_DRAGMOUSEMOVE											15
#define DISPID_EXTVWE_DROP															16
#define DISPID_EXTVWE_EDITCHANGE												17
#define DISPID_EXTVWE_EDITCLICK													18
#define DISPID_EXTVWE_EDITCONTEXTMENU										19
#define DISPID_EXTVWE_EDITDBLCLICK											20
#define DISPID_EXTVWE_EDITGOTFOCUS											21
#define DISPID_EXTVWE_EDITKEYDOWN												22
#define DISPID_EXTVWE_EDITKEYPRESS											23
#define DISPID_EXTVWE_EDITKEYUP													24
#define DISPID_EXTVWE_EDITLOSTFOCUS											25
#define DISPID_EXTVWE_EDITMCLICK												26
#define DISPID_EXTVWE_EDITMDBLCLICK											27
#define DISPID_EXTVWE_EDITMOUSEDOWN											28
#define DISPID_EXTVWE_EDITMOUSEENTER										29
#define DISPID_EXTVWE_EDITMOUSEHOVER										30
#define DISPID_EXTVWE_EDITMOUSELEAVE										31
#define DISPID_EXTVWE_EDITMOUSEMOVE											32
#define DISPID_EXTVWE_EDITMOUSEUP												33
#define DISPID_EXTVWE_EDITRCLICK												34
#define DISPID_EXTVWE_EDITRDBLCLICK											35
#define DISPID_EXTVWE_EXPANDEDITEM											36
#define DISPID_EXTVWE_EXPANDINGITEM											37
#define DISPID_EXTVWE_FREEITEMDATA											38
#define DISPID_EXTVWE_INCREMENTALSEARCHSTRINGCHANGING		39
#define DISPID_EXTVWE_INSERTEDITEM											40
#define DISPID_EXTVWE_INSERTINGITEM											41
#define DISPID_EXTVWE_ITEMBEGINDRAG											42
#define DISPID_EXTVWE_ITEMBEGINRDRAG										43
#define DISPID_EXTVWE_ITEMGETDISPLAYINFO								44
#define DISPID_EXTVWE_ITEMGETINFOTIPTEXT								45
#define DISPID_EXTVWE_ITEMMOUSEENTER										46
#define DISPID_EXTVWE_ITEMMOUSELEAVE										47
#define DISPID_EXTVWE_ITEMSELECTIONCHANGED							48
#define DISPID_EXTVWE_ITEMSETTEXT												49
#define DISPID_EXTVWE_ITEMSTATEIMAGECHANGED							50
#define DISPID_EXTVWE_ITEMSTATEIMAGECHANGING						51
#define DISPID_EXTVWE_KEYDOWN														52
#define DISPID_EXTVWE_KEYPRESS													53
#define DISPID_EXTVWE_KEYUP															54
#define DISPID_EXTVWE_MCLICK														55
#define DISPID_EXTVWE_MDBLCLICK													56
#define DISPID_EXTVWE_MOUSEDOWN													57
#define DISPID_EXTVWE_MOUSEENTER												58
#define DISPID_EXTVWE_MOUSEHOVER												59
#define DISPID_EXTVWE_MOUSELEAVE												60
#define DISPID_EXTVWE_MOUSEMOVE													61
#define DISPID_EXTVWE_MOUSEUP														62
#define DISPID_EXTVWE_OLECOMPLETEDRAG										63
#define DISPID_EXTVWE_OLEDRAGDROP												64
#define DISPID_EXTVWE_OLEDRAGENTER											65
#define DISPID_EXTVWE_OLEDRAGLEAVE											66
#define DISPID_EXTVWE_OLEDRAGMOUSEMOVE									67
#define DISPID_EXTVWE_OLEGIVEFEEDBACK										68
#define DISPID_EXTVWE_OLEQUERYCONTINUEDRAG							69
#define DISPID_EXTVWE_OLESETDATA												70
#define DISPID_EXTVWE_OLESTARTDRAG											71
#define DISPID_EXTVWE_RCLICK														72
#define DISPID_EXTVWE_RDBLCLICK													73
#define DISPID_EXTVWE_RECREATEDCONTROLWINDOW						74
#define DISPID_EXTVWE_REMOVEDITEM												75
#define DISPID_EXTVWE_REMOVINGITEM											76
#define DISPID_EXTVWE_RENAMEDITEM												77
#define DISPID_EXTVWE_RENAMINGITEM											78
#define DISPID_EXTVWE_RESIZEDCONTROLWINDOW							79
#define DISPID_EXTVWE_SINGLEEXPANDING										80
#define DISPID_EXTVWE_STARTINGLABELEDITING							81
#define DISPID_EXTVWE_ITEMASYNCHRONOUSDRAWFAILED				82
#define DISPID_EXTVWE_ITEMSELECTIONCHANGING							83
#define DISPID_EXTVWE_OLEDRAGENTERPOTENTIALTARGET				84
#define DISPID_EXTVWE_OLEDRAGLEAVEPOTENTIALTARGET				85
#define DISPID_EXTVWE_OLERECEIVEDNEWDATA								86
#define DISPID_EXTVWE_CHANGEDSORTORDER									87
#define DISPID_EXTVWE_CHANGINGSORTORDER									88
#define DISPID_EXTVWE_EDITMOUSEWHEEL										89
#define DISPID_EXTVWE_EDITXCLICK												90
#define DISPID_EXTVWE_EDITXDBLCLICK											91
#define DISPID_EXTVWE_MOUSEWHEEL												92
#define DISPID_EXTVWE_XCLICK														93
#define DISPID_EXTVWE_XDBLCLICK													94


// ITreeViewItem
// properties
#define DISPID_TVI_ACCESSIBILITYID											1
#define DISPID_TVI_BOLD																	2
#define DISPID_TVI_CARET																3
#define DISPID_TVI_DROPHILITED													4
#define DISPID_TVI_EXPANDEDICONINDEX										5
#define DISPID_TVI_EXPANSIONSTATE												6
#define DISPID_TVI_FIRSTSUBITEM													7
#define DISPID_TVI_GHOSTED															8
#define DISPID_TVI_GRAYED																9
#define DISPID_TVI_HANDLE																10
#define DISPID_TVI_HASEXPANDO														11
#define DISPID_TVI_HEIGHTINCREMENT											12
#define DISPID_TVI_ICONINDEX														13
#define DISPID_TVI_ID																		14
#define DISPID_TVI_ITEMDATA															15
#define DISPID_TVI_LASTSUBITEM													16
#define DISPID_TVI_LEVEL																17
#define DISPID_TVI_NEXTSIBLINGITEM											18
#define DISPID_TVI_NEXTVISIBLEITEM											19
#define DISPID_TVI_OVERLAYINDEX													20
//TODO: #define DISPID_TVI_PADDING															21
#define DISPID_TVI_PARENTITEM														22
#define DISPID_TVI_PREVIOUSSIBLINGITEM									23
#define DISPID_TVI_PREVIOUSVISIBLEITEM									24
#define DISPID_TVI_SELECTED															25
#define DISPID_TVI_SELECTEDICONINDEX										26
#ifdef INCLUDESHELLBROWSERINTERFACE
	#define DISPID_TVI_SHELLTREEVIEWITEMOBJECT							27
#endif
#define DISPID_TVI_STATEIMAGEINDEX											28
#define DISPID_TVI_SUBITEMS															29
#define DISPID_TVI_TEXT																	DISPID_VALUE
#define DISPID_TVI_VIRTUAL															30
// methods
#define DISPID_TVI_COLLAPSE															31
#define DISPID_TVI_COULDMOVETO													32
#define DISPID_TVI_CREATEDRAGIMAGE											33
#define DISPID_TVI_DISPLAYINFOTIP												34
#define DISPID_TVI_ENSUREVISIBLE												35
#define DISPID_TVI_EXPAND																36
#define DISPID_TVI_GETPATH															37
#define DISPID_TVI_GETRECTANGLE													38
#define DISPID_TVI_ISTRUNCATED													39
#define DISPID_TVI_MOVE																	40
#define DISPID_TVI_SORTSUBITEMS													41
#define DISPID_TVI_STARTLABELEDITING										42


// ITreeViewItems
// properties
#define DISPID_TVIS_CASESENSITIVEFILTERS								1
#define DISPID_TVIS_COMPARISONFUNCTION									2
#define DISPID_TVIS_FILTER															3
#define DISPID_TVIS_FILTERTYPE													4
#define DISPID_TVIS_ITEM																DISPID_VALUE
#define DISPID_TVIS__NEWENUM														DISPID_NEWENUM
#define DISPID_TVIS_PARENTITEM													5
#define DISPID_TVIS_SINGLELEVEL													6
// methods
#define DISPID_TVIS_ADD																	7
#define DISPID_TVIS_CONTAINS														8
#define DISPID_TVIS_COUNT																9
#define DISPID_TVIS_REMOVE															10
#define DISPID_TVIS_REMOVEALL														11
#define DISPID_TVIS_FINDBYPATH													12


// ITreeViewItemContainer
// properties
#define DISPID_TVIC_ITEM																DISPID_VALUE
#define DISPID_TVIC__NEWENUM														DISPID_NEWENUM
// methods
#define DISPID_TVIC_ADD																	1
#define DISPID_TVIC_CLONE																2
#define DISPID_TVIC_COUNT																3
#define DISPID_TVIC_CREATEDRAGIMAGE											4
#define DISPID_TVIC_REMOVE															5
#define DISPID_TVIC_REMOVEALL														6


// IVirtualTreeViewItem
// properties
#define DISPID_VTVI_BOLD																1
#define DISPID_VTVI_DROPHILITED													2
#define DISPID_VTVI_EXPANSIONSTATE											3
#define DISPID_VTVI_GHOSTED															4
#define DISPID_VTVI_HASEXPANDO													5
#define DISPID_VTVI_HEIGHTINCREMENT											6
#define DISPID_VTVI_ICONINDEX														7
#define DISPID_VTVI_ITEMDATA														8
#define DISPID_VTVI_LEVEL																9
#define DISPID_VTVI_NEXTSIBLINGITEM											10
#define DISPID_VTVI_OVERLAYINDEX												11
#define DISPID_VTVI_PARENTITEM													12
#define DISPID_VTVI_PREVIOUSSIBLINGITEM									13
#define DISPID_VTVI_SELECTED														14
#define DISPID_VTVI_SELECTEDICONINDEX										15
#define DISPID_VTVI_STATEIMAGEINDEX											16
#define DISPID_VTVI_TEXT																DISPID_VALUE
#define DISPID_VTVI_EXPANDEDICONINDEX										18
#define DISPID_VTVI_GRAYED															19
#define DISPID_VTVI_VIRTUAL															20
//TODO: #define DISPID_VTVI_PADDING															21
// methods
#define DISPID_VTVI_GETPATH															17


// IOLEDataObject
// properties
// methods
#define DISPID_ODO_CLEAR																1
#define DISPID_ODO_GETCANONICALFORMAT										2
#define DISPID_ODO_GETDATA															3
#define DISPID_ODO_GETFORMAT														4
#define DISPID_ODO_SETDATA															5
#define DISPID_ODO_GETDROPDESCRIPTION										6
#define DISPID_ODO_SETDROPDESCRIPTION										7
