; ======================================================================================================================
; Namespace:      LV_Colors
; Function:       Individual row and cell coloring for AHK ListView controls.
; Testted with:   AHK 1.1.20.03 (A32/U32/U64)
; Tested on:      Win 8.1 (x64)
; Changelog:
;     1.0.00.00/2015-03-23/just me - new version using new AHK 1.1.20+ features
;     0.5.00.00/2014-08-13/just me - changed 'static mode' handling
;     0.4.01.00/2013-12-30/just me - minor bug fix
;     0.4.00.00/2013-12-30/just me - added static mode
;     0.3.00.00/2013-06-15/just me - added "Critical, 100" to avoid drawing issues
;     0.2.00.00/2013-01-12/just me - bugfixes and minor changes
;     0.1.00.00/2012-10-27/just me - initial release
; ======================================================================================================================
; CLASS LV_Colors
;
; The class provides six public methods to set individual colors for rows and/or cells, to clear all colors, to
; prevent/allow sorting and rezising of columns dynamically, and to remove/add a message handler for WM_NOTIFY messages.
;
; The message handler for WM_NOTIFY messages will be activated for the specified ListView whenever a new instance is
; created. If you want to use an own message handler, set the OnMessage parameter to False when creating the new
; instance or call MyNewInstance.OnMessage(False) after the new instance has been created.
; ======================================================================================================================
Class LV_Colors {
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; META FUNCTIONS ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; -------------------------------------------------------------------------------------------------------------------
   ; __New()         Create a new LV_Colors instance for the given ListView
   ; Parameters:     HWND        -  ListView's HWND.
   ;                 Optional ------------------------------------------------------------------------------------------
   ;                 OnMessage   -  Add a message handler for WM_NOTIFY messages for this ListView.
   ;                                Values:  True/False
   ;                                Default: True
   ;                 StaticMode  -  Static color assignment, i.e. the colors will be assigned permanently to a row
   ;                                rather than to a row number.
   ;                                Values:  True/False
   ;                                Default: False
   ;                 NoSort      -  Prevent sorting by click on a header item.
   ;                                Values:  True/False
   ;                                Default: True
   ;                 NoSizing    -  Prevent resizing of columns.
   ;                                Values:  True/False
   ;                                Default: True
   ; -------------------------------------------------------------------------------------------------------------------
   __New(HWND, OnMessage := True, StaticMode := False, NoSort := True, NoSizing := True) {
   ; Attach(HWND, StaticMode := False, NoSort := True, NoSizing := True) {
      If (This.Base.Base.__Class) ; do not instantiate instances
         Return False
      If This.Attached[HWND] ; HWND is already attached
         Return False
      If !DllCall("IsWindow", "Ptr", HWND) ; invalid HWND
         Return False
      VarSetCapacity(Class, 512, 0)
      DllCall("GetClassName", "Ptr", HWND, "Str", Class, "Int", 256)
      If (Class <> "SysListView32") ; HWND doesn't belong to a ListView
         Return False
      ; ----------------------------------------------------------------------------------------------------------------
      ; Set LVS_EX_DOUBLEBUFFER (0x010000) style to avoid drawing issues.
      SendMessage, 0x1036, 0x010000, 0x010000, , % "ahk_id " . HWND ; LVM_SETEXTENDEDLISTVIEWSTYLE
      ; Get the default colors
      SendMessage, 0x1000, 0, 0, , % "ahk_id " . HWND ; LVM_GETBKCOLOR
      This.BkClr := ErrorLevel
      SendMessage, 0x1025, 0, 0, , % "ahk_id " . HWND ; LVM_GETTEXTBKCOLOR
      This.TxBkClr := ErrorLevel
      SendMessage, 0x1023, 0, 0, , % "ahk_id " . HWND ; LVM_GETTEXTCOLOR
      This.TxClr := ErrorLevel
      ; Get the header control
      SendMessage, 0x101F, 0, 0, , % "ahk_id " . HWND ; LVM_GETHEADER
      This.Header := ErrorLevel
      ; Set other properties
      This.HWND := HWND
      This.IsStatic := !!StaticMode
      If (NoSort)
         This.NoSort()
      If (NoSizing)
         This.NoSizing()
      This.OnMessage(!!OnMessage)
      This.Attached[HWND] := True
   }
   ; -------------------------------------------------------------------------------------------------------------------
   __Delete() {
      This.Attached.Remove(HWND, "")
      This.OnMessage(False)
      WinSet, Redraw, , % "ahk_id " . This.HWND
   }
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PRIVATE PROPERTIES  +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   Static Attached := {}
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PRIVATE METHODS +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   On_WM_NOTIFY(W, L, M, H) {
      ; Notifications: NM_CUSTOMDRAW = -12, LVN_COLUMNCLICK = -108, HDN_BEGINTRACKA = -306, HDN_BEGINTRACKW = -326
      Critical, % LV_Colors.Critical
      If ((HCTL := NumGet(L + 0, 0, "UPtr")) = This.HWND) || (HCTL = This.Header) {
         Code := NumGet(L + (A_PtrSize * 2), 0, "Int")
         If (Code = -12)
            Return This.On_NM_CUSTOMDRAW(H, L)
         If This.NoSort && (Code = -108)
            Return 0
         If This.NoSizing && ((Code = -306) || (Code = -326))
            Return True
      }
   }
   ; -------------------------------------------------------------------------------------------------------------------
   On_NM_CUSTOMDRAW(H, L) {
      ; Return values: 0x00 (CDRF_DODEFAULT), 0x20 (CDRF_NOTIFYITEMDRAW / CDRF_NOTIFYSUBITEMDRAW)
      ; Size off NMHDR structure
      Static SizeNMHDR := A_PtrSize * 3
      ; Size of NMCUSTOMDRAW structure
      Static SizeNCD := SizeNMHDR + 16 + (A_PtrSize * 5)
      ; Offset of dwItemSpec (NMCUSTOMDRAW)
      Static OffItem := SizeNMHDR + 16 + (A_PtrSize * 2)
      ; Offset of clrText (NMLVCUSTOMDRAW)
      Static OffCT :=  SizeNCD
      ; Offset of clrTextBk (NMLVCUSTOMDRAW)
      Static OffCTB := OffCT + 4
      ; Offset of iSubItem (NMLVCUSTOMDRAW)
      Static OffSubItem := OffCTB + 4
      ; Offset of clrFace (NMLVCUSTOMDRAW)
      Static OffCB := OffSubItem + 8
      ; ----------------------------------------------------------------------------------------------------------------
      DrawStage := NumGet(L + SizeNMHDR, 0, "UInt")
      , Row := NumGet(L + OffItem, 0, "UPtr") + 1
      , Col := NumGet(L + OffSubItem, 0, "Int") + 1
      If This.IsStatic
         Row := This.MapIndexToID(H, Row)
      ; CDDS_SUBITEMPREPAINT = 0x030001 --------------------------------------------------------------------------------
      If (DrawStage = 0x030001) {
         NumPut(((ClrTx := This["Cells", Row, Col, "T"]) <> "") ? ClrTx : This.RowCT, L + OffCT, 0, "UInt")
         , NumPut(((ClrBk := This["Cells", Row, Col, "B"]) <> "") ? ClrBk : This.RowCB, L + OffCB, 0, "UInt")
         , NumPut((ClrBk <> "") ? ClrBk : This.RowCTB, L + OffCTB, 0, "UInt")
         Return (!This.HasKey(Row) && (Col > This["Cells", Row].MaxIndex())) ? 0x00 : 0x20
      }
      ; CDDS_ITEMPREPAINT = 0x010001 -----------------------------------------------------------------------------------
      If (DrawStage = 0x010001) {
         This.RowCT := This.TxClr, This.RowCTB := This.TxBkClr, This.RowCB := This.BkClr
         If ((ClrTx := This["Rows", Row, "T"]) <> "")
            NumPut((This.RowCT := ClrTx), L + OffCT, 0, "UInt")
         If ((ClrBk := This["Rows", Row, "B"]) <> "")
            NumPut((This.RowCTB := ClrBk), L + OffCTB, 0, "UInt"), NumPut((This.RowCB := ClrBk), L + OffCB, 0, "UInt")
         Return This["Cells"].HasKey(Row) ? 0x20 : 0x00
      }
      ; CDDS_PREPAINT = 0x000001 ---------------------------------------------------------------------------------------
      Return (DrawStage = 0x000001) ? 0x20 : 0x00
   }
   ; -------------------------------------------------------------------------------------------------------------------
   MapIndexToID(Row) {
      SendMessage, 0x10B4, % (Row - 1), 0, , % "ahk_id " . This.HWND ; LVM_MAPINDEXTOID
      Return ErrorLevel
   }
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PUBLIC PROPERTIES  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   Static Critical := 100
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PUBLIC METHODS ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ===================================================================================================================
   ; Clear()         Clears all colors and removes the message handler for this ListView.
   ; Return Value:   Always True.
   ; ===================================================================================================================
   Clear() {
      This.Remove("Rows")
      This.Remove("Cells")
      This.OnMessage(False)
      Return True
   }
   ; ===================================================================================================================
   ; Row()           Setd background and/or text color for the specified row.
   ; Parameters:     Row         -  Row number
   ;                 Optional ------------------------------------------------------------------------------------------
   ;                 BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red)
   ;                                Default: Empty -> default background color
   ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red)
   ;                                Default: Empty -> default text color
   ; Return Value:   True on success, otherwise false.
   ; ===================================================================================================================
   Row(Row, BkColor := "", TxColor := "") {
      If !(This.HWND)
         Return False
      If This.IsStatic
         Row := This.MapIndexToID(Row)
      If (BkColor = "") && (TxColor = "") {
         This["Rows"].Remove(Row, "")
         Return True
      }
      BkBGR := ((BkColor & 0xFF0000) >> 16) | (BkColor & 0x00FF00) | ((BkColor & 0x0000FF) << 16)
      TxBGR := ((TxColor & 0xFF0000) >> 16) | (TxColor & 0x00FF00) | ((TxColor & 0x0000FF) << 16)
      If (BkBGR = "") && (TxBGR = "")
         Return False
      If !This.HasKey("Rows")
         This.Rows := {}
      If !This.Rows.HasKey(Row)
         This.Rows[Row] := {}
      If (BkBGR <> "")
         This.Rows[Row, "B"] := BkBGR
      If (TxBGR <> "")
         This.Rows[Row, "T"] := TxBGR
      Return True
   }
   ; ===================================================================================================================
   ; Cell()          Sets background and/or text color for the specified cell.
   ; Parameters:     Row         -  Row number
   ;                 Col         -  Column number
   ;                 Optional ------------------------------------------------------------------------------------------
   ;                 BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red)
   ;                                Default: Empty -> default background color
   ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red)
   ;                                Default: Empty -> default text color
   ; Return Value:   True on success, otherwise false.
   ; ===================================================================================================================
   Cell(Row, Col, BkColor := "", TxColor := "") {
      If !(This.HWND)
         Return False
      If ThisIsStatic
         Row := This.MapIndexToID(Row)
      If (BkColor = "") && (TxColor = "") {
         This["Cells"].Remove(Row, "")
         Return True
      }
      BkBGR := ((BkColor & 0xFF0000) >> 16) | (BkColor & 0x00FF00) | ((BkColor & 0x0000FF) << 16)
      TxBGR := ((TxColor & 0xFF0000) >> 16) | (TxColor & 0x00FF00) | ((TxColor & 0x0000FF) << 16)
      If (BkBGR = "") && (TxBGR = "")
         Return False
      If !This.HasKey("Cells")
         This.Cells := {}
      If !This.Cells.HasKey(Row)
         This.Cells[Row] := {}
      This.Cells[Row, Col] := {}
      If (BkBGR <> "")
         This.Cells[Row, Col, "B"] := BkBGR
      If (TxBGR <> "")
         This.Cells[Row, Col, "T"] := TxBGR
      Return True
   }
   ; ===================================================================================================================
   ; NoSort()        Prevents/allows sorting by click on a header item for this ListView.
   ; Parameters:     Apply       -  True/False
   ;                                Default: True
   ; Return Value:   True on success, otherwise false.
   ; ===================================================================================================================
   NoSort(Apply := True) {
      If !(This.HWND)
         Return False
      If (Apply)
         This.NoSort := True
      Else
         This.NoSort := False
      Return True
   }
   ; ===================================================================================================================
   ; NoSizing()      Prevents/allows resizing of columns for this ListView.
   ; Parameters:     Apply       -  True/False
   ;                                Default: True
   ; Return Value:   True on success, otherwise false.
   ; ===================================================================================================================
   NoSizing(Apply := True) {
      Static OSVersion := DllCall("GetVersion", "UChar")
      If !(This.Header)
         Return False
      If (Apply) {
         If (OSVersion > 5)
            Control, Style, +0x0800, , % "ahk_id " . This.Header ; HDS_NOSIZING = 0x0800
         This.NoSizing := True
      }
      Else {
         If (OSVersion > 5)
            Control, Style, -0x0800, , % "ahk_id " . This.Header ; HDS_NOSIZING
         This.NoSizing := False
      }
      Return True
   }
   ; ===================================================================================================================
   ; OnMessage()     Adds/removes a message handler for WM_NOTIFY messages for this ListView.
   ; Parameters:     Apply       -  True/False
   ;                                Default: True
   ; Return Value:   Always True
   ; ===================================================================================================================
   OnMessage(Apply := True) {
      If (Apply) && !This.HasKey("OnMessageFunc") {
         This.OnMessageFunc := ObjBindMethod(This, "On_WM_Notify")
         OnMessage(0x004E, This.OnMessageFunc) ; add the WM_NOTIFY message handler
      }
      Else If !(Apply) && This.HasKey("OnMessageFunc") {
         OnMessage(0x004E, This.OnMessageFunc, 0) ; remove the WM_NOTIFY message handler
         This.OnMessageFunc := ""
         This.Remove("OnMessageFunc")
      }
      WinSet, Redraw, , % "ahk_id " . This.HWND
      Return True
   }
}