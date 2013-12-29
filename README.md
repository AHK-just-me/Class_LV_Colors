# Class_LV_Colors #

The class supports individually colored rows and cells for AHK ListView controls.

### How to use: ###
- First call `LV_Colors.Attach(HLV)` passing the HWND of your ListView.
- If you want to use the included message handler you must call `LV_Colors.OnMessage()` once. Otherwise you have to integrate the code within `LV_Colors_WM_NOTIFY(`) into your own notification handler.
- Then call `LV_Colors.Cell()` or `LV_Colors.Row()` to setup colors for individual cells and/or rows.
- That's all you have to do for coloring.
- If you finally don't want the colors to be shown any more, call LV_Colors.Detach(HLV) to restore the ListView's default behaviour.  

For more detailed informations look at the inline documentation, please.