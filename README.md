Friendly.FarPoint
============================

This library is a layer on top of
Friendly, so you must learn that first.
But it is very easy to learn.

https://github.com/Codeer-Software/Friendly.Windows 

## Getting Started
Install Friendly.FarPoint from NuGet

    Install-Package Friendly.FarPoint
https://github.com/Codeer-Software/Friendly.FarPoint

***
Friendly.FarPoint defines the following classes.   
They can operate the control easily from a separate process.  

* FpSpreadDriver

***
```cs  
//sample  
var process = Process.GetProcessesByName("Target")[0];  
using (var app = new WindowsAppFriend(process))  
{  
    dynamic main = app.Type(typeof(Application)).OpenForms[0];  
    var spread = new FpSpreadDriver(main._grid);
    
    //sheets.
    int count = spread.Sheets.Count;
    int activeSheetIndex = spread.ActiveSheetIndex;
    spread.EmulateChangeActiveSheet(1);
    var sheet = spread.Sheets[1];
    sheet = spread.ActiveSheet;
    
    //cell.
    var cell = sheet.Cells[0. 3];
    cell = sheet.ActiveCell;
    string text = cell.Text;
    int rowIndex = cell.Row.Index;
    int rowIndex2 = cell.Row.Index2;
    int rowIndex = cell.Column.Index;
    int rowIndex2 = cell.Column.Index2;
    sheet.EmulateChangeActiveCell(3, 5, true);
    sheet.EmulateAddSelection(1, 2, 3, 5);
    sheet.EmulateRemoveSelection(1, 2, 3, 5);
    sheet.EmulateClearSelection();
    
    //edit.
    sheet.EmulateChangeActiveCell(0, 1, true);
    spread.EmualteEditText("abc");
    
    sheet.EmulateChangeActiveCell(0, 2, true);
    spread.EmualteEditValue(2);
    
    sheet.EmulateChangeActiveCell(0, 3, true);
    spread.EmualteEditSelectedIndex(1);
    
    sheet.EmulateChangeActiveCell(0, 4, true);
    spread.EmualteEditCheckState(CheckState.Checked);
}
```
### More samples.
https://github.com/Codeer-Software/Friendly.FarPoint/tree/master/Project/Test

***
For other GUI types, use the following libraries:

* For Win32.  
https://www.nuget.org/packages/Codeer.Friendly.Windows.NativeStandardControls/  

* For WinForms.  
https://www.nuget.org/packages/Ong.Friendly.FormsStandardControls/  

* For WPF.  
https://www.nuget.org/packages/RM.Friendly.WPFStandardControls/

* For getting the target window.  
https://www.nuget.org/packages/Codeer.Friendly.Windows.Grasp/  

* For 3d party.  
https://www.nuget.org/packages/Friendly.XamControls/  
https://www.nuget.org/packages/Friendly.C1.Win/  
https://www.nuget.org/packages/Friendly.FarPoint/  

***
If you use PinInterface, you map control simple.  
https://www.nuget.org/packages/VSHTC.Friendly.PinInterface/

