﻿; ======================================================================================================================
; Namespace:      LV_Colors
; Function:       Helper object and functions for ListView row and cell coloring
; Testted with:   AHK 1.1.13.01 (A32/U32/U64)
; Tested on:      Win 7 (x64)
; Changelog:
;     0.4.01.00/2013-12-30/just me - minor bug fix
;     0.4.00.00/2013-12-30/just me - added static mode
;     0.3.00.00/2013-06-15/just me - added "Critical, 100" to avoid drawing issues
;     0.2.00.00/2013-01-12/just me - bugfixes and minor changes
;     0.1.00.00/2012-10-27/just me - initial release
; ======================================================================================================================
; CLASS LV_Colors
;
; The class provides seven public methods to register / unregister coloring for ListView controls, to set individual
; colors for rows and/or cells, to prevent/allow sorting and rezising dynamically, and to register / unregister the
; included message handler function for WM_NOTIFY -> NM_CUSTOMDRAW messages.
;
; If you want to use the included message handler you must call LV_Colors.OnMessage() once.
; Otherwise you should integrate the code within LV_Colors_WM_NOTIFY into your own notification handler.
; Without notification handling coloring won't work.
; ======================================================================================================================
Class LV_Colors {
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PRIVATE PROPERTIES ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   Static MessageHandler := "LV_Colors_WM_NOTIFY"
   Static WM_NOTIFY := 0x4E
   Static SubclassProc := RegisterCallback("LV_Colors_SubclassProc")
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PUBLIC PROPERTIES  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   Static Critical := 100
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; META FUNCTIONS ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   __New(P*) {
      Return False   ; There is no reason to instantiate this class!
   }
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PRIVATE METHODS +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   On_NM_CUSTOMDRAW(H, L) {
      Static CDDS_PREPAINT          := 0x00000001
      Static CDDS_ITEMPREPAINT      := 0x00010001
      Static CDDS_SUBITEMPREPAINT   := 0x00030001
      Static CDRF_DODEFAULT         := 0x00000000
      Static CDRF_NEWFONT           := 0x00000002
      Static CDRF_NOTIFYITEMDRAW    := 0x00000020
      Static CDRF_NOTIFYSUBITEMDRAW := 0x00000020
      Static CLRDEFAULT             := 0xFF000000
      ; Size off NMHDR structure
      Static NMHDRSize := (2 * A_PtrSize) + 4 + (A_PtrSize - 4)
      ; Offset of dwItemSpec (NMCUSTOMDRAW)
      Static ItemSpecP := NMHDRSize + (5 * 4) + A_PtrSize + (A_PtrSize - 4)
      ; Size of NMCUSTOMDRAW structure
      Static NCDSize  := NMHDRSize + (6 * 4) + (3 * A_PtrSize) + (2 * (A_PtrSize - 4))
      ; Offset of clrText (NMLVCUSTOMDRAW)
      Static ClrTxP   :=  NCDSize
      ; Offset of clrTextBk (NMLVCUSTOMDRAW)
      Static ClrTxBkP := ClrTxP + 4
      ; Offset of iSubItem (NMLVCUSTOMDRAW)
      Static SubItemP := ClrTxBkP + 4
      ; Offset of clrFace (NMLVCUSTOMDRAW)
      Static ClrBkP   := SubItemP + 8
      DrawStage := NumGet(L + NMHDRSize, 0, "UInt")
      , Row := NumGet(L + ItemSpecP, 0, "UPtr") + 1
      , Col := NumGet(L + SubItemP, 0, "Int") + 1
      If This[H].IsStatic
         Row := This.GetItemParam(H, Row)
      ; SubItemPrepaint ------------------------------------------------------------------------------------------------
      If (DrawStage = CDDS_SUBITEMPREPAINT) {
         NumPut(This[H].CurTX, L + ClrTxP, 0, "UInt"), NumPut(This[H].CurTB, L + ClrTxBkP, 0, "UInt")
         , NumPut(This[H].CurBK, L + ClrBkP, 0, "UInt")
         ClrTx := This[H].Cells[Row][Col].T, ClrBk := This[H].Cells[Row][Col].B
         If (ClrTx <> "")
            NumPut(ClrTX, L + ClrTxP, 0, "UInt")
         If (ClrBk <> "")
            NumPut(ClrBk, L + ClrTxBkP, 0, "UInt"), NumPut(ClrBk, L + ClrBkP, 0, "UInt")
         If (Col > This[H].Cells[Row].MaxIndex()) && !This[H].HasKey(Row)
            Return CDRF_DODEFAULT
         Return CDRF_NOTIFYSUBITEMDRAW
      }
      ; ItemPrepaint ---------------------------------------------------------------------------------------------------
      If (DrawStage = CDDS_ITEMPREPAINT) {
         This[H].CurTX := This[H].TX, This[H].CurTB := This[H].TB, This[H].CurBK := This[H].BK
         ClrTx := ClrBk := ""
         If This[H].Rows.HasKey(Row)
            ClrTx := This[H].Rows[Row].T, ClrBk := This[H].Rows[Row].B
         If (ClrTx <> "")
            NumPut(ClrTx, L + ClrTxP, 0, "UInt"), This[H].CurTX := ClrTx
         If (ClrBk <> "")
            NumPut(ClrBk, L + ClrTxBkP, 0, "UInt") , NumPut(ClrBk, L + ClrBkP, 0, "UInt")
            , This[H].CurTB := ClrBk, This[H].CurBk := ClrBk
         If This[H].Cells.HasKey(Row)
            Return CDRF_NOTIFYSUBITEMDRAW
         Return CDRF_DODEFAULT
      }
      ; Prepaint -------------------------------------------------------------------------------------------------------
      If (DrawStage = CDDS_PREPAINT) {
         Return CDRF_NOTIFYITEMDRAW
      }
      ; Others ---------------------------------------------------------------------------------------------------------
      Return CDRF_DODEFAULT
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetItemParam(HWND, Row) {
      Static LVM_GETITEM := 0x1005 ; it's save to use LVM_GETITEMA
      Static LVIF_PARAM := 0x00000004
      Static LVITEMsize := 72 ; size of 64-bit structure, it doesn't matter in this case
      Static OffParam := 24 + (A_PtrSize * 2) ; offset of the lParam member
      VarSetCapacity(LVITEM, LVITEMsize, 0)
      , NumPut(LVIF_PARAM, LVITEM, 0, "UInt")
      , NumPut(Row - 1, LVITEM, 4, "Int")
      SendMessage, % LVM_GETITEM, 0, % &LVITEM, , % "ahk_id " . HWND
      Return NumGet(LVITEM, OffParam, "UPtr")
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetItemParam(HWND, Row, Param) {
      Static LVM_SETITEM := 0x1006 ; it's save to use LVM_SETITEMA
      Static LVIF_PARAM := 0x00000004
      Static LVITEMsize := 72 ; size of 64-bit structure, it doesn't matter in this case
      Static OffParam := 24 + (A_PtrSize * 2) ; offset of the lParam member
      VarSetCapacity(LVITEM, LVITEMsize, 0)
      , NumPut(LVIF_PARAM, LVITEM, 0, "UInt")
      , NumPut(Row - 1, LVITEM, 4, "Int")
      , NumPut(Param, LVITEM, OffParam, "Ptr")
      SendMessage, % LVM_SETITEM, 0, % &LVITEM, , % "ahk_id " . HWND
      Return ErrorLevel
   }
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PUBLIC METHODS ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ===================================================================================================================
   ; Attach()        Register ListView control for coloring
   ; Parameters:     HWND        -  ListView's HWND.
   ;                 Optional ------------------------------------------------------------------------------------------
   ;                 StaticMode  -  Static color assignment, i.e. the colors will be assigned permanently to a row
   ;                                rather than to a row number.
   ;                                Values:  True / False
   ;                                Default: False
   ;                 NoSort      -  Prevent sorting by click on a header item.
   ;                                Values:  True / False
   ;                                Default: True
   ;                 NoSizing    -  Prevent resizing of columns.
   ;                                Values:  True / False
   ;                                Default: True
   ; Return Values:  True on success, otherwise false.
   ; ===================================================================================================================
   Attach(HWND, StaticMode := False, NoSort := True, NoSizing := True) {
      Static LVM_GETBKCOLOR     := 0x1000
      Static LVM_GETHEADER      := 0x101F
      Static LVM_GETTEXTBKCOLOR := 0x1025
      Static LVM_GETTEXTCOLOR   := 0x1023
      Static LVM_SETEXTENDEDLISTVIEWSTYLE := 0x1036
      Static LVS_EX_DOUBLEBUFFER := 0x00010000
      If !DllCall("User32.dll\IsWindow", "Ptr", HWND, "UInt")
         Return False
      If This.HasKey(HWND)
         Return False
      ; Set LVS_EX_DOUBLEBUFFER style to avoid drawing issues, if it isn't set as yet.
      SendMessage, % LVM_SETEXTENDEDLISTVIEWSTYLE, % LVS_EX_DOUBLEBUFFER, % LVS_EX_DOUBLEBUFFER, , % "ahk_id " . HWND
      If (ErrorLevel = "FAIL")
         Return False
      ; Get the default colors
      SendMessage, % LVM_GETBKCOLOR, 0, 0, , % "ahk_id " . HWND
      BkClr := ErrorLevel
      SendMessage, % LVM_GETTEXTBKCOLOR, 0, 0, , % "ahk_id " . HWND
      TBClr := ErrorLevel
      SendMessage, % LVM_GETTEXTCOLOR, 0, 0, , % "ahk_id " . HWND
      TxClr := ErrorLevel
      ; Get the header control
      SendMessage, % LVM_GETHEADER, 0, 0, , % "ahk_id " . HWND
      Header := ErrorLevel
      ; Store the values in a new object
      This[HWND] := {BK: BkClr, TB: TBClr, TX: TxClr, Header: Header, IsStatic: !!StaticMode}
      If (NoSort)
         This.NoSort(HWND)
      If (NoSizing)
         This.NoSizing(HWND)
      Return True
   }
   ; ===================================================================================================================
   ; Detach()        Unregister ListView control
   ; Parameters:     HWND        -  ListView's HWND
   ; Return Value:   Always True
   ; ===================================================================================================================
   Detach(HWND) {
      ; Remove the subclass, if any
      Static LVM_GETITEMCOUNT := 0x1004
      If (This[HWND].SC)
         DllCall("Comctl32.dll\RemoveWindowSubclass", "Ptr", HWND, "Ptr", This.SubclassProc, "Ptr", HWND)
      If This[HWND].IsStatic {
         SendMessage, % LVM_GETITEMCOUNT, 0, 0, , % "ahk_id " . HWND
         ItemCount := ErrorLevel
         Loop, % ItemCount
            This.SetItemParam(HWND, A_Index, 0)
      }
      This.Remove(HWND, "")
      WinSet, Redraw, , % "ahk_id " . HWND
      Return True
   }
   ; ===================================================================================================================
   ; Row()           Set background and/or text color for the specified row
   ; Parameters:     HWND        -  ListView's HWND
   ;                 Row         -  Row number
   ;                 Optional ------------------------------------------------------------------------------------------
   ;                 BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red)
   ;                                Default: Empty -> default background color
   ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red)
   ;                                Default: Empty -> default text color
   ; Return Value:   True on success, otherwise false.
   ; ===================================================================================================================
   Row(HWND, Row, BkColor := "", TxColor := "") {
      If !This.HasKey(HWND)
         Return False
      If (BkColor = "") && (TxColor = "") {
         This[HWND].Rows.Remove(Row, "")
         Return True
      }
      BkBGR := TxBGR := ""
      If BkColor Is Integer
         BkBGR := ((BkColor & 0xFF0000) >> 16) | (BkColor & 0x00FF00) | ((BkColor & 0x0000FF) << 16)
      If TxColor Is Integer
         TxBGR := ((TxColor & 0xFF0000) >> 16) | (TxColor & 0x00FF00) | ((TxColor & 0x0000FF) << 16)
      If (BkBGR = "") && (TxBGR = "")
         Return False
      If !This[HWND].HasKey("Rows")
         This[HWND].Rows := {}
      If !This[HWND].Rows.HasKey(Row)
         This[HWND].Rows[Row] := {}
      If (BkBGR <> "")
         This[HWND].Rows[Row].Insert("B", BkBGR)
      If (TxBGR <> "")
         This[HWND].Rows[Row].Insert("T", TxBGR)
      If This[HWND].IsStatic
         This.SetItemParam(HWND, Row, Row)
      Return True
   }
   ; ===================================================================================================================
   ; Cell()          Set background and/or text color for the specified cell
   ; Parameters:     HWND        -  ListView's HWND
   ;                 Row         -  Row number
   ;                 Col         -  Column number
   ;                 Optional ------------------------------------------------------------------------------------------
   ;                 BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red)
   ;                                Default: Empty -> default background color
   ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red)
   ;                                Default: Empty -> default text color
   ; Return Value:   True on success, otherwise false.
   ; ===================================================================================================================
   Cell(HWND, Row, Col, BkColor := "", TxColor := "") {
      If !This.HasKey(HWND)
         Return False
      If (BkColor = "") && (TxColor = "") {
         This[HWND].Cells.Remove(Row, "")
         Return True
      }
      BkBGR := TxBGR := ""
      If BkColor Is Integer
         BkBGR := ((BkColor & 0xFF0000) >> 16) | (BkColor & 0x00FF00) | ((BkColor & 0x0000FF) << 16)
      If TxColor Is Integer
         TxBGR := ((TxColor & 0xFF0000) >> 16) | (TxColor & 0x00FF00) | ((TxColor & 0x0000FF) << 16)
      If (BkBGR = "") && (TxBGR = "")
         Return False
      If !This[HWND].HasKey("Cells")
         This[HWND].Cells := {}
      If !This[HWND].Cells.HasKey(Row)
         This[HWND].Cells[Row] := {}
      This[HWND].Cells[Row, Col] := {}
      If (BkBGR <> "")
         This[HWND].Cells[Row, Col].Insert("B", BkBGR)
      If (TxBGR <> "")
         This[HWND].Cells[Row, Col].Insert("T", TxBGR)
      If This[HWND].IsStatic
         This.SetItemParam(HWND, Row, Row)
      Return True
   }
   ; ===================================================================================================================
   ; NoSort()        Prevent / allow sorting by click on a header item dynamically.
   ; Parameters:     HWND        -  ListView's HWND
   ;                 Optional ------------------------------------------------------------------------------------------
   ;                 Apply       -  True / False
   ;                                Default: True
   ; Return Value:   True on success, otherwise false.
   ; ===================================================================================================================
   NoSort(HWND, Apply := True) {
      Static HDM_GETITEMCOUNT := 0x1200
      If !This.HasKey(HWND)
         Return False
      If (Apply)
         This[HWND].NS := True
      Else
         This[HWND].Remove("NS")
      Return True
   }
   ; ===================================================================================================================
   ; NoSizing()      Prevent / allow resizing of columns dynamically.
   ; Parameters:     HWND        -  ListView's HWND
   ;                 Optional ------------------------------------------------------------------------------------------
   ;                 Apply       -  True / False
   ;                                Default: True
   ; Return Value:   True on success, otherwise false.
   ; ===================================================================================================================
   NoSizing(HWND, Apply := True) {
      Static OSVersion := DllCall("Kernel32.dll\GetVersion", "UChar")
      Static HDS_NOSIZING := 0x0800
      If !This.HasKey(HWND)
         Return False
      HHEADER := This[HWND].Header
      If (Apply) {
         If (OSVersion < 6) {
            If !(This[HWND].SC) {
               DllCall("Comctl32.dll\SetWindowSubclass", "Ptr", HWND, "Ptr", This.SubclassProc, "Ptr", HWND, "Ptr", 0)
               This[HWND].SC := True
            } Else {
               Return True
            }
         } Else {
            Control, Style, +%HDS_NOSIZING%, , ahk_id %HHEADER%
         }
      } Else {
         If (OSVersion < 6) {
            If (This[HWND].SC) {
               DllCall("Comctl32.dll\RemoveWindowSubclass", "Ptr", HWND, "Ptr", This.SubclassProc, "Ptr", HWND)
               This[HWND].Remove("SC")
            } Else {
               Return True
            }
         } Else {
            Control, Style, -%HDS_NOSIZING%, , ahk_id %HHEADER%
         }
      }
      Return True
   }
   ; ===================================================================================================================
   ; OnMessage()     Register / unregister LV_Colors message handler for WM_NOTIFY -> NM_CUSTOMDRAW messages
   ; Parameters:     Apply       -  True / False
   ;                                Default: True
   ; Return Value:   Always True
   ; ===================================================================================================================
   OnMessage(Apply := True) {
      If (Apply)
         OnMessage(This.WM_NOTIFY, This.MessageHandler)
      Else If (This.MessageHandler = OnMessage(This.WM_NOTIFY))
         OnMessage(This.WM_NOTIFY, "")
      Return True
   }
}
; ======================================================================================================================
; PRIVATE FUNCTION LV_Colors_WM_NOTIFY() - message handler for WM_NOTIFY -> NM_CUSTOMDRAW notifications
; ======================================================================================================================
LV_Colors_WM_NOTIFY(W, L) {
   Static NM_CUSTOMDRAW := -12
   Static LVN_COLUMNCLICK := -108
   Critical, % LV_Colors.Critical
   If LV_Colors.HasKey(H := NumGet(L + 0, 0, "UPtr")) {
      M := NumGet(L + (A_PtrSize * 2), 0, "Int")
      ; NM_CUSTOMDRAW --------------------------------------------------------------------------------------------------
      If (M = NM_CUSTOMDRAW)
         Return LV_Colors.On_NM_CUSTOMDRAW(H, L)
      ; LVN_COLUMNCLICK ------------------------------------------------------------------------------------------------
      If (LV_Colors[H].NS && (M = LVN_COLUMNCLICK))
         Return 0
   }
}
; ======================================================================================================================
; PRIVATE FUNCTION LV_Colors_SubclassProc() - subclass for WM_NOTIFY -> HDN_BEGINTRACK notifications (Win XP)
; ======================================================================================================================
LV_Colors_SubclassProc(H, M, W, L, S, R) {
   Static HDN_BEGINTRACKA := -306
   Static HDN_BEGINTRACKW := -326
   Static WM_NOTIFY := 0x4E
   Critical, % LV_Colors.Critical
   If (M = WM_NOTIFY) {
      ; HDN_BEGINTRACK -------------------------------------------------------------------------------------------------
      C := NumGet(L + (A_PtrSize * 2), 0, "Int")
      If (C = HDN_BEGINTRACKA) || (C = HDN_BEGINTRACKW) {
         Return True
      }
   }
   Return DllCall("Comctl32.dll\DefSubclassProc", "Ptr", H, "UInt", M, "Ptr", W, "Ptr", L, "UInt")
}
; ======================================================================================================================