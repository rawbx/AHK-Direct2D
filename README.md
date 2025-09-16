# AHK-Direct2D

A simple Direct2D wrapper for ahk v2 

## Acknowledgements

- [Direct2D - Win32 apps  Microsoft Learn](https://learn.microsoft.com/windows/win32/Direct2D/direct2d-portal)
- [Direct2D Win32 API](https://learn.microsoft.com/zh-cn/windows/win32/api/_direct2d/)

## Usage

```autoit
ui := Gui("-DPIScale")
d2d := Direct2D(ui.Hwnd, 512, 512)
ui.Show("W512 H512")
d2d.BeginDraw()
d2d.DrawSvg('<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" ...</svg>', 10, 10, 128, 128)
d2d.EndDraw()
```

## Credit

Some code are adapted from this project: [Spawnova/ShinsOverlayClass - MIT License](https://github.com/Spawnova/ShinsOverlayClass)
