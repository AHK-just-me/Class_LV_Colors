#NoEnv
#Include Class_LV_Colors.ahk
SetBatchLines, -1
Gui, Margin, 20, 20
Gui, Add, ListView, w600 r15 Grid -ReadOnly vVLV hwndHLV
   , Column 1|Column 2|Column 3|Column 4|Column 5|Column6
Loop, 256
   LV_Add("", "Value " . A_Index, "Value " . A_Index, "Value " . A_Index, "Value " . A_Index, "Value "
        . A_Index, "Value " . A_Index)
Loop, % LV_GetCount("Column")
   LV_ModifyCol(A_Index, "AutoHdr")
; Create a new instance of LV_Colors
CLV := New LV_Colors(HLV)
If !IsObject(CLV) {
   MsgBox, 0, ERROR, Couldn't create a new LV_Colors object!
   ExitApp
}
; Set the colors
Loop, 256 {
   If (A_Index & 1) {
      If (Mod(A_Index, 3) = 0)
         CLV.Row(A_Index, 0xFF0000, 0xFFFF00)
      CLV.Cell(A_Index, 1, 0x00FF00, 0x000080)
      CLV.Cell(A_Index, 3, 0x00FF00, 0x000080)
      CLV.Cell(A_Index, 5, 0x00FF00, 0x000080)
   }
   Else
      CLV.Row(A_Index, 0x000080, 0x00FF00)
}
Gui, Add, Button, wp gSubColors vBtnColors, Colors off!
Gui, Show, , ListView & Colors
; Redraw the ListView after the first Gui, Show command to show the colors.
WinSet, Redraw, , ahk_id %HLV%
Return
; ----------------------------------------------------------------------------------------------------------------------
GuiClose:
GuiEscape:
ExitApp
; ----------------------------------------------------------------------------------------------------------------------
SetColors:
GuiControl, -Redraw, %HLV%
GuiControl, +Redraw, %HLV%
Return
; ----------------------------------------------------------------------------------------------------------------------
SubColors:
   GuiControlGet, BtnColors
   If (BtnColors = "Colors on!") {
      CLV.OnMessage()
      GuiControl, , BtnColors, Colors off!
   }
   Else {
      CLV.OnMessage(False)
      GuiControl, , BtnColors, Colors on!
   }
   GuiControl, Focus, %HLV%
Return