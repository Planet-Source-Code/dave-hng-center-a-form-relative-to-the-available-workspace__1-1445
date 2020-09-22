<div align="center">

## Center a form, relative to the available workspace


</div>

### Description

Centers a form, relative to the available workspace. This means that if your users have high, or wide taskbars, or other apps which restrict the workspace, your forms will still center properly.
 
### More Info
 
General usage:

Stick this in the form's show event:

Center Me

It's a sub, so no returns.

If the workspace is smaller than the form, it still centers, so part of the form will be off the visible area (of course this is a problem with all form centering code).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dave Hng](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dave-hng.md)
**Level**          |Unknown
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dave-hng-center-a-form-relative-to-the-available-workspace__1-1445/archive/master.zip)

### API Declarations

```
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Global Const SPI_GETWORKAREA As Long = 48
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
```


### Source Code

```
Public Sub Center(ByRef frm As Form)
'Centers a form, relative to the available workspace
Dim rt As RECT, result As Long
Dim X As Single, Y As Single
Dim oldScaleMode As Integer
result = SystemParametersInfo(SPI_GETWORKAREA, 0&, rt, 0&)
X = rt.Right - rt.Left
Y = rt.Bottom - rt.Top
X = X * Screen.TwipsPerPixelX
Y = Y * Screen.TwipsPerPixelY
X = X \ 2 - (frm.Width \ 2)
Y = Y \ 2 - (frm.Height \ 2)
oldScaleMode = frm.ScaleMode
frm.ScaleMode = vbTwips
frm.Move X, Y
frm.ScaleMode = oldScaleMode
End Sub
```

