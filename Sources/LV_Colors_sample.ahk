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
Gui, Add, Button, wp gSubColors vBtnColors, Colors on!
Gui, Show, , ListView & Colors
LV_Colors.OnMessage()
Return
; ----------------------------------------------------------------------------------------------------------------------
GuiClose:
GuiEscape:
ExitApp
; ----------------------------------------------------------------------------------------------------------------------
SubColors:
   GuiControlGet, BtnColors
   GuiControl, -Redraw, %HLV%
   If (BtnColors = "Colors on!") {
      If !LV_Colors.Attach(HLV, 1, 0, 0) {
         GuiControl, +Redraw, %HLV%
         Return
      }
      Sleep, 10
      Loop, 256 {
         If (A_Index & 1) {
            If (Mod(A_Index, 3) = 0)
               LV_Colors.Row(HLV, A_Index, 0xFF0000, 0xFFFF00)
            LV_Colors.Cell(HLV, A_Index, 1, 0x00FF00, 0x000080)
            LV_Colors.Cell(HLV, A_Index, 3, 0x00FF00, 0x000080)
            LV_Colors.Cell(HLV, A_Index, 5, 0x00FF00, 0x000080)
         } Else {
            LV_Colors.Row(HLV, A_Index, 0x000080, 0x00FF00)
         }
      GuiControl, , BtnColors, Colors off!
      }
   } Else {
      LV_Colors.Detach(HLV)
      GuiControl, , BtnColors, Colors on!
   }
   GuiControl, +Redraw, %HLV%
Return