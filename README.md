# Class_LV_Colors #

The class supports individually colored rows and cells for AHK ListView controls.

### How to use: ###
- Create a new instance with  `MyInstance := New LV_Colors(HLV)` passing the HWND of your ListView.
- Then call `MyInstance.Cell()` or `MyInstance.Row()` to setup colors for individual cells and/or rows.
- That's all you have to do for coloring.
- If you finally don't want the colors to be shown any more, use `MyInstance := ""` to restore the ListView's default behaviour.  
For more detailed informations look at the inline documentation, please.
