/*******************************************************************************
 * @description
 * A simple Direct2D wrapper for ahk v2
 *
 * Some code are adapted from this project: [Spawnova/ShinsOverlayClass - MIT License](https://github.com/Spawnova/ShinsOverlayClass)
 *
 * @link https://github.com/rawbx/AHK-Direct2D
 * @file Direct2D.ahk
 * @license MIT
 * @author rawbx
 * @version 0.1.3
 * @example
 * ui := Gui("-DPIScale")
 * d2d := Direct2D(ui.Hwnd, 512, 512)
 * ui.Show("W512 H512")
 * d2d.BeginDraw()
 * d2d.DrawSvg('<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" ...</svg>', 10, 10, 128, 128)
 * d2d.EndDraw()
 ******************************************************************************/
class Direct2D {
    __New(target, w?, h?) {
        this.target := target
        this.hwnd := target is Gui ? target.Hwnd : (IsInteger(target) && target != 0) ? target : 0
        this.x := 0, this.y := 0, this.width := w ?? 512, this.height := h ?? 512
        this.isDrawing := 0

        ; to get image bitmap info
        this.gdipToken := this.GdipStartUp()

        this.textFormats := Map()
        this.solidBrushes := Map()
        this.strokeStyles := Map()
        this.d2dBitmaps := Map()

        this.ID2D1WICBitmapRenderTarget := Direct2D.ID2D1WicBitmapRenderTarget(this.width, this.height)
        this.ID2D1RenderTarget := this.hwnd ? Direct2D.ID2D1RenderTarget(this.hwnd, this.width, this.height)
            : this.ID2D1WICBitmapRenderTarget

        this.ID2D1RenderTarget.SetAntiAliasMode(D2D1_ANTIALIAS_MODE_PER_PRIMITIVE := 0)
        this.ID2D1RenderTarget.SetTextAntiAliasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE := 1)
        this.Clear()
    }

    __Delete() {
        for _, s in this.solidBrushes
            Direct2D.release(s)
        for _, t in this.textFormats
            Direct2D.release(t)
        for _, st in this.strokeStyles
            Direct2D.release(st)

        Direct2D.release(Direct2D.IDWriteFactory.Get())
        Direct2D.release(Direct2D.ID2D1Factory.Get())
        this.GdipShutdown(this.gdipToken)
    }

    static isX64 := A_PtrSize == 8
    static vTable(p, i) => NumGet(NumGet(p, 0, "ptr"), i * A_PtrSize, "ptr")
    static release(p) => DllCall(this.vTable(p, 2), "ptr", p)
    static getCLSID(guid, &clsid) => DllCall("ole32\CLSIDFromString", "WStr", guid, "Ptr", clsid := buffer(16, 0))

    class ID2D1Factory {
        static __New() {
            D2D1CreateFactory := DllCall("GetProcAddress", "ptr", DllCall("LoadLibrary", "str", "d2d1.dll", "ptr"), "astr", "D2D1CreateFactory", "ptr")
            if !D2D1CreateFactory
                throw Error("Failed to load D2D1CreateFactory")

            IID_ID2DFactory := "{06152247-6f50-465a-9245-118bfd3b6007}"
            Direct2D.getCLSID(IID_ID2DFactory, &clsidDF := 0)
            if DllCall(D2D1CreateFactory,
                "uint", 1,
                "Ptr", clsidDF,
                "uint*", 0,
                "Ptr*", &pFactory := 0
            ) != 0
                throw Error("D2D1CreateFactory failed")

            this.pF := pFactory
            this.VT_GetDesktopDpi := Direct2D.vTable(pFactory, 4)
            this.VT_CreateStrokeStyle := Direct2D.vTable(pFactory, 11)
            this.VT_CreateWicBitmapRenderTarget := Direct2D.vTable(pFactory, 13)
            this.VT_CreateHwndRenderTarget := Direct2D.vTable(pFactory, 14)
            this.VT_CreateDCRenderTarget := Direct2D.vTable(pFactory, 16)
        }

        static Get() => this.pF

        static GetDesktopDpi() =>
            (DllCall(this.VT_GetDesktopDpi, "ptr", this.pF, 'float*', &dpiX := 0, 'float*', &dpiY := 0, 'uint'), dpiX)

        static CreateStrokeStyle(styleProps) =>
            (DllCall(this.VT_CreateStrokeStyle, "ptr", this.pF, "ptr", styleProps, "ptr", 0, "uint", 0, "ptr*", &pStrokeStyle := 0), pStrokeStyle)

        static CreateWicBitmapRenderTarget(pWICBitmap, rtProps) =>
            (DllCall(this.VT_CreateWicBitmapRenderTarget, "Ptr", this.pF, "Ptr", pWICBitmap, "ptr", rtProps, "Ptr*", &pRenderTarget := 0), pRenderTarget)

        static CreateHwndRenderTarget(rtProps, hRtProps) =>
            (DllCall(this.VT_CreateHwndRenderTarget, "Ptr", this.pF, "Ptr", rtProps, "ptr", hRtProps, "Ptr*", &pRenderTarget := 0), pRenderTarget)

        static CreateDCRenderTarget(rtProps) =>
            (DllCall(this.VT_CreateDCRenderTarget, "Ptr", this.pF, "Ptr", rtProps, "Ptr*", &pRenderTarget := 0), pRenderTarget)
    }

    class IDWriteFactory {
        static __New() {
            DWriteCreateFactory := DllCall("GetProcAddress", "ptr", DllCall("LoadLibrary", "str", "dwrite.dll", "ptr"), "astr", "DWriteCreateFactory", "ptr")
            if !DWriteCreateFactory
                throw Error("Failed to load DWriteCreateFactory")

            IID_IDWriteFactory := "{B859EE5A-D838-4B5B-A2E8-1ADC7D93DB48}"
            Direct2D.getCLSID(IID_IDWriteFactory, &clsidWF := 0)
            if DllCall(DWriteCreateFactory,
                "uint", 0,
                "ptr", clsidWF,
                "ptr*", &pWFactory := 0
            ) != 0
                throw Error("DWriteCreateFactory failed")

            this.pWF := pWFactory
            this.VT_CreateTextFormat := Direct2D.vTable(pWFactory, 15)
            this.VT_CreateTextLayout := Direct2D.vTable(pWFactory, 18)
            return pWFactory
        }

        static Get() => this.pWF

        static CreateTextFormat(fontName, fontSize) =>
            (DllCall(this.VT_CreateTextFormat, "ptr", this.pWF,
                "wstr", fontName,
                "ptr", 0,
                "uint", 400, ; DWRITE_FONT_WEIGHT_NORMAL
                "uint", 0, ; DWRITE_FONT_STYLE_NORMAL
                "uint", 5, ; DWRITE_FONT_STRETCH_NORMAL
                "float", fontSize,
                "wstr", "en-us",
                "Ptr*", &pTextFormat := 0
            ), pTextFormat)

        static CreateTextLayout(text, pTextFormat) =>
            (DllCall(this.VT_CreateTextLayout, "ptr", this.pWF,
                "wstr", text,
                "uint", StrLen(text),
                "ptr", pTextFormat,
                "float", A_ScreenWidth,  ; maxWidth
                "float", A_ScreenHeight,  ; maxHeight
                "ptr*", &pTextLayout := 0,
                "uint"
            ), pTextLayout)
    }

    ; GetMetrics for the formatted text.
    ; https://learn.microsoft.com/windows/win32/api/dwrite/ns-dwrite-dwrite_text_metrics
    GetMetrics(text, fontName := "Segoe UI", fontSize := 16) {
        pTextFormat := this.GetSavedTextFormat(fontName, fontSize)
        pTextLayout := Direct2D.IDWriteFactory.CreateTextLayout(text, pTextFormat)
        ; struct DWRITE_TEXT_METRICS {
        ;     FLOAT left;
        ;     FLOAT top;
        ;     FLOAT width;
        ;     FLOAT widthIncludingTrailingWhitespace;
        ;     FLOAT height;
        ;     FLOAT layoutWidth;
        ;     FLOAT layoutHeight;
        ;     UINT32 maxBidiReorderingDepth;
        ;     UINT32 lineCount;
        ; };
        textMetrics := Buffer(4 * 9)
        if DllCall(Direct2D.vTable(pTextLayout, 60), ;IDWriteTextLayout::GetMetrics
            "ptr", pTextLayout,
            "ptr", textMetrics,
            "uint"
        ) != 0
            throw Error("GetMetrics failed")

        width := NumGet(textMetrics, 8, "float")
        height := NumGet(textMetrics, 16, "float")

        Direct2D.release(pTextLayout)
        ; Direct2D.release(pTextFormat) ; pTextFormat will release in map
        return { w: width, h: height }
    }

    GetDesktopDpiScale() {
        dpiX := Direct2D.ID2D1Factory.GetDesktopDpi()
        return dpiX / 96
    }

    class ID2D1RenderTarget {
        __New(target, width, height) {
            this.pRT := 0
            this.width := width, this.height := height
            rtProps := Buffer(64, 0)
            NumPut("uint", 1, rtProps, 8) ; AlphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED
            NumPut("float", 96, rtProps, 12) ; dpiX
            NumPut("float", 96, rtProps, 16) ; dpiY
            hRtProps := Buffer(64, 0)
            NumPut("Uptr", target, hRtProps, 0)
            NumPut("uint", width, hRtProps, A_PtrSize)
            NumPut("uint", height, hRtProps, A_PtrSize + 4)
            NumPut("uint", 2, hRtProps, A_PtrSize + 8)

            ; set window visible
            DllCall("SetLayeredWindowAttributes", "Uptr", target, "Uint", ColorKey := 0, "char", Alpha := 255, "uint", LWA_ALPHA := 2)

            margins := Buffer(16, 0)
            NumPut("int", -1, margins, 0), NumPut("int", -1, margins, 4)
            NumPut("int", -1, margins, 8), NumPut("int", -1, margins, 12)
            DllCall("dwmapi\DwmExtendFrameIntoClientArea", "Uptr", target, "ptr", margins, "uint")

            this.pRT := Direct2D.ID2D1Factory.CreateHwndRenderTarget(rtProps, hRtProps)
            if !this.pRT {
                MsgBox("ID2D1RenderTarget init failed")
                return 0
            }
            this.__InitCommonVT()

            ; ID2D1HwndRenderTarget
            this.VT_CheckWindowState := Direct2D.vTable(this.pRT, 57)
            this.VT_Resize := Direct2D.vTable(this.pRT, 58)
            this.VT_GetHwnd := Direct2D.vTable(this.pRT, 59)

            return this.pRT
        }

        __Delete() {
            if this.pRT
                Direct2D.release(this.pRT)
        }

        __InitCommonVT() { ; ID2D1RenderTarget
            this.VT_CreateBitmap := Direct2D.vTable(this.pRT, 4)
            this.VT_CreateBitmapFromWicBitmap := Direct2D.vTable(this.pRT, 5)
            this.VT_CreateSolidBrush := Direct2D.vTable(this.pRT, 8)
            this.VT_DrawLine := Direct2D.vTable(this.pRT, 14)
            this.VT_DrawRectangle := Direct2D.vTable(this.pRT, 16)
            this.VT_FillRectangle := Direct2D.vTable(this.pRT, 17)
            this.VT_DrawRoundedRectangle := Direct2D.vTable(this.pRT, 18)
            this.VT_FillRoundedRectangle := Direct2D.vTable(this.pRT, 19)
            this.VT_DrawEllipse := Direct2D.vTable(this.pRT, 20)
            this.VT_FillEllipse := Direct2D.vTable(this.pRT, 21)
            this.VT_DrawBitmap := Direct2D.vTable(this.pRT, 26)
            this.VT_DrawText := Direct2D.vTable(this.pRT, 27)
            this.VT_DrawTextLayout := Direct2D.vTable(this.pRT, 28)
            this.VT_SetAntialiasMode := Direct2D.vTable(this.pRT, 32)
            this.VT_SetTextAntialiasMode := Direct2D.vTable(this.pRT, 34)
            this.VT_Clear := Direct2D.vTable(this.pRT, 47)
            this.VT_BeginDraw := Direct2D.vTable(this.pRT, 48)
            this.VT_EndDraw := Direct2D.vTable(this.pRT, 49)
        }

        CreateBitmap(w, h, srcData, pitch, bmpProps) {
            pBitmap := 0
            if Direct2D.isX64 {
                bf := Buffer(64)
                NumPut("uint", w, bf, 0)
                NumPut("uint", h, bf, 4)
                DllCall(this.VT_CreateBitmap, "Ptr", this.pRT, "int64", NumGet(bf, 0, "int64"), "Ptr", srcData, "uint", pitch, "Ptr", bmpProps, "Ptr*", &pBitmap)
            }
            else {
                DllCall(this.VT_CreateBitmap, "Ptr", this.pRT, "uint", w, "uint", h, "Ptr", srcData, "uint", pitch, "Ptr", bmpProps, "Ptr*", &pBitmap)
            }
            return pBitmap
        }

        CreateBitmapFromWicBitmap(pWicBitmapSource, bmpProps) =>
            (DllCall(this.VT_CreateBitmapFromWicBitmap, "Ptr", this.pRT, "Ptr", pWicBitmapSource, "Ptr", bmpProps, "Ptr*", &pBitmap := 0), pBitmap)

        CreateSolidBrush(sColor, brushProps := 0) =>
            (DllCall(this.VT_CreateSolidBrush, "Ptr", this.pRT, "Ptr", sColor, "Ptr", brushProps, "Ptr*", &pBrush := 0), pBrush)

        ; D2D1_POINT_2F is different in 32 and 64 bits
        ; reference https://github.com/Spawnova/ShinsOverlayClass/blob/main/AHK%20V2/ShinsOverlayClass.ahk#L626
        DrawLine(pointStart, pointEnd, pBrush, strokeWidth, pStrokeStyle) {
            if Direct2D.isX64 {
                points := Buffer(64)
                NumPut("float", pointStart[1], points, 0)  ;Special thanks to teadrinker for helping me
                NumPut("float", pointStart[2], points, 4)  ;with these params!
                NumPut("float", pointEnd[1], points, 8)
                NumPut("float", pointEnd[2], points, 12)
                DllCall(this.VT_DrawLine, "Ptr", this.pRT, "Double", NumGet(points, 0, "double"), "Double", NumGet(points, 8, "double"), "ptr", pBrush, "float", strokeWidth, "ptr", pStrokeStyle)
            } else {
                DllCall(this.VT_DrawLine, "Ptr", this.pRT, "float", pointStart[1], "float", pointStart[2], "float", pointEnd[1], "float", pointEnd[2], "ptr", pBrush, "float", strokeWidth, "ptr", pStrokeStyle)
            }
        }

        DrawRectangle(rect, pBrush, strokeWidth, pStrokeStyle) =>
            DllCall(this.VT_DrawRectangle, "Ptr", this.pRT, "Ptr", rect, "ptr", pBrush, "float", strokeWidth, "ptr", pStrokeStyle)

        FillRectangle(rect, pBrush) =>
            DllCall(this.VT_FillRectangle, "Ptr", this.pRT, "Ptr", rect, "ptr", pBrush)

        DrawRoundedRectangle(roundedRect, pBrush, strokeWidth, pStrokeStyle) =>
            DllCall(this.VT_DrawRoundedRectangle, "Ptr", this.pRT, "Ptr", roundedRect, "ptr", pBrush, "float", strokeWidth, "ptr", pStrokeStyle)

        FillRoundedRectangle(roundedRect, pBrush) =>
            DllCall(this.VT_FillRoundedRectangle, "Ptr", this.pRT, "Ptr", roundedRect, "ptr", pBrush)

        DrawEllipse(sEllipse, pBrush, strokeWidth, pStrokeStyle) =>
            DllCall(this.VT_DrawEllipse, "Ptr", this.pRT, "Ptr", sEllipse, "ptr", pBrush, "float", strokeWidth, "ptr", pStrokeStyle)

        FillEllipse(sEllipse, pBrush) =>
            DllCall(this.VT_FillEllipse, "Ptr", this.pRT, "Ptr", sEllipse, "ptr", pBrush)

        DrawBitmap(pBitmap, destRect := 0, opacity := 1, interpolationMode := 1, srcRect := 0) =>
            DllCall(this.VT_DrawBitmap, "ptr", this.pRT,
                "ptr", pBitmap,
                "ptr", destRect,
                "float", opacity,
                "uint", interpolationMode,
                "ptr", srcRect)

        DrawText(text, pTextFormat, rect, pBrush, drawOpt) =>
            DllCall(this.VT_DrawText, "ptr", this.pRT, "wstr", text, "uint", strlen(text), "ptr", pTextFormat, "ptr", rect, "ptr", pBrush, "uint", drawOpt, "uint", 0)

        DrawTextLayout(point, pTextLayout, pBrush, drawOpt) {
            if Direct2D.isX64 {
                topLeftPt := Buffer(8)
                NumPut('float', point[1], topLeftPt, 0)
                NumPut('float', point[2], topLeftPt, 4)
                DllCall(this.VT_DrawTextLayout, "ptr", this.pRT, 'double', NumGet(topLeftPt, 0, 'double'), "ptr", pTextLayout, 'ptr', pBrush, "uint", drawOpt)
            } else {
                DllCall(this.VT_DrawTextLayout, "ptr", this.pRT, "float", point[1], "float", point[2], "ptr", pTextLayout, 'ptr', pBrush, "uint", drawOpt)
            }
        }

        SetAntialiasMode(mode) => DllCall(this.VT_SetAntialiasMode, "Ptr", this.pRT, "Uint", mode)

        SetTextAntialiasMode(mode) => DllCall(this.VT_SetTextAntialiasMode, "Ptr", this.pRT, "Uint", mode)

        Clear() => DllCall(this.VT_Clear, "Ptr", this.pRT, "Ptr", 0)

        BeginDraw() => DllCall(this.VT_BeginDraw, "Ptr", this.pRT)

        EndDraw() => DllCall(this.VT_EndDraw, "Ptr", this.pRT, "Ptr*", 0, "Ptr*", 0)

        Resize(size) => DllCall(this.VT_Resize, "Ptr", this.pRT, "ptr", size)
    }

    class ID2D1WicBitmapRenderTarget extends Direct2D.ID2D1RenderTarget {
        __New(width, height) {
            rtProps := Buffer(64, 0)
            NumPut("uint", 1, rtProps, 8) ; AlphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED
            NumPut("uint", 1, rtProps, 0) ; D2D1_RENDER_TARGET_TYPE_SOFTWARE
            NumPut("uint", 87, rtProps, 4) ; DXGI_FORMAT_B8G8R8A8_UNORM
            ; Initialize Windows Imaging Component.
            CLSID_WICImagingFactory := "{CACAF262-9370-4615-A13B-9F5539DA4C0A}"
            IID_IWICImagingFactory := "{EC5EC8A9-C395-4314-9C77-54D7A935FF70}"
            pWICImagingFactory := ComObject(CLSID_WICImagingFactory, IID_IWICImagingFactory)
            Direct2D.getCLSID(WICPixelFormat32bppPBGRA := "{6fddc324-4e03-4bfe-b185-3d77768dc910}", &clsidPBGRA)
            ComCall(17, pWICImagingFactory, ; IWICImagingFactory::CreateBitmap in memory
                "uint", width, "uint", height,
                "ptr", this.clsidPBGRA := clsidPBGRA,
                "uint", WICBitmapCacheOnDemand := 1,
                "ptr*", &pWICBitmap := 0)
            this.pWICBitmap := pWICBitmap
            this.pRT := Direct2D.ID2D1Factory.CreateWicBitmapRenderTarget(pWicBitmap, rtProps)
            if !this.pRT {
                MsgBox("ID2D1WicBitmapRenderTarget init failed")
                return 0
            }

            super.__InitCommonVT()
            this.VT_CreateSvgDocument := Direct2D.vTable(this.pRT, 115)
            this.VT_DrawSvgDocument := Direct2D.vTable(this.pRT, 116)

            return this.pRT
        }

        __Delete() {
            if this.pRT
                Direct2D.release(this.pRT)
            if this.pWICBitmap
                Direct2D.release(this.pWICBitmap)
        }

        GetWICBitmap() => this.pWICBitmap

        CreateSvgDocument(fs, sizeF) => (DllCall(this.VT_CreateSvgDocument, "Ptr", this.pRT, "ptr", fs, "uint64", sizeF, "ptr*", &pSvgDocument := 0), pSvgDocument)

        DrawSvgDocument(pSvgDocument) => DllCall(this.VT_DrawSvgDocument, "Ptr", this.pRT, "Ptr", pSvgDocument)

        CreateSvgHBitmap(fs, width, height) {
            D2D1_SIZE_F := Buffer(8)
            NumPut("float", width, D2D1_SIZE_F, 0x0)
            NumPut("float", height, D2D1_SIZE_F, 0x4)
            pSvgDocument := this.CreateSvgDocument(fs, NumGet(D2D1_SIZE_F, "uint64")) ; ID2D1SvgDocument

            this.BeginDraw()
            this.DrawSvgDocument(pSvgDocument)
            this.EndDraw()

            hBitmap := this.GetHBitmapFromWICBitmap(width, height)
            Direct2D.release(pSvgDocument)
            return hBitmap
        }

        GetHBitmapFromWICBitmap(width, height) {
            if !this.pWICBitmap
                return

            stride := 4 * width
            pData := Buffer(stride * height)
            ComCall(CopyPixels := 7, this.pWICBitmap, "ptr", 0, "uint", stride, "uint", pData.Size, "ptr", pData)
            hBitmap := DllCall("gdi32\CreateBitmap", "Int", width, "Int", height, "Uint", 1, "Uint", 32, "Ptr", pData, "Ptr")

            Direct2D.release(this.pWICBitmap), this.pWICBitmap := 0
            return hBitmap
        }
    }

    class ID2D1DCRenderTarget extends Direct2D.ID2D1RenderTarget {
        __New() {
            rtProps := Buffer(64, 0)
            NumPut("uint", 87, rtProps, 4) ; DXGI_FORMAT_B8G8R8A8_UNORM
            this.pRT := Direct2D.ID2D1Factory.CreateDCRenderTarget(rtProps)
            if !this.pRT {
                MsgBox("ID2D1DCRenderTarget init failed")
                return 0
            }

            super.__InitCommonVT()
            this.VT_BindDC := Direct2D.vTable(this.pRT, 57)

            return this.pRT
        }

        BindDC(rect) => DllCall(this.VT_BindDC, "Ptr", this.pRT, "ptr", rect)
    }

    Clear() {
        this.ID2D1RenderTarget.BeginDraw()
        this.ID2D1RenderTarget.Clear()
        this.ID2D1RenderTarget.EndDraw()
    }

    BeginDraw() {
        this.ID2D1RenderTarget.BeginDraw()
        this.ID2D1RenderTarget.Clear()
        return this.isDrawing := 1
    }

    EndDraw() {
        if (this.isDrawing)
            this.ID2D1RenderTarget.EndDraw()
    }

    DrawText(text, x, y, fontSize, color, fontName, w?, h?, drawOpt := 4) {
        pTextFormat := this.GetSavedTextFormat(fontName, fontSize)
        pBrush := this.GetSavedSolidBrush(color)
        sRect := Buffer(64)
        NumPut("float", x, sRect, 0)
        NumPut("float", y, sRect, 4)
        NumPut("float", x + (w ?? this.width), sRect, 8)
        NumPut("float", y + (h ?? this.height), sRect, 12)

        ; https://learn.microsoft.com/windows/win32/api/d2d1/ne-d2d1-d2d1_draw_text_options
        ; typedef enum D2D1_DRAW_TEXT_OPTIONS {
        ;   D2D1_DRAW_TEXT_OPTIONS_NO_SNAP = 0x00000001,
        ;   D2D1_DRAW_TEXT_OPTIONS_CLIP = 0x00000002,
        ;   D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT = 0x00000004,
        ;   D2D1_DRAW_TEXT_OPTIONS_DISABLE_COLOR_BITMAP_SNAPPING = 0x00000008,
        ;   D2D1_DRAW_TEXT_OPTIONS_NONE = 0x00000000,
        ;   D2D1_DRAW_TEXT_OPTIONS_FORCE_DWORD = 0xffffffff
        ; }
        this.ID2D1RenderTarget.DrawText(text, pTextFormat, sRect, pBrush, drawOpt)
    }

    DrawTextLayout(text, x, y, color, pTextLayout, drawOpt := 4) {
        if !text
            return

        pBrush := this.GetSavedSolidBrush(color)
        this.ID2D1RenderTarget.DrawTextLayout([x, y], pTextLayout, pBrush, drawOpt)

        Direct2D.release(pTextLayout)
    }

    /**
     * @param {Array}  pointStart [x, y]
     * @param {Array} pointEnd [x, y]
     * @param {Integer} color abgr 0xFFFFFFFF
     * @param {Integer} strokeWidth thickness for stroke
     * @param {Integer} strokeCapStyle flat(0) square(1) round(2) triangle(3)
     * @param {Integer} strokeShapeStyle solid(0) dash(1) dot(2) dash_dot(3)
     */
    DrawLine(pointStart, pointEnd, color := 0xFFFFFFFF, strokeWidth := 2, strokeCapStyle := 2, strokeShapeStyle := 0) {
        pBrush := this.GetSavedSolidBrush(color)
        pStrokeStyle := this.GetSavedStrokeStyle(strokeCapStyle, strokeShapeStyle)
        this.ID2D1RenderTarget.DrawLine(pointStart, pointEnd, pBrush, strokeWidth, pStrokeStyle)
    }

    DrawRectangle(x, y, w, h, color, strokeWidth := 2, strokeCapStyle := 0, strokeShapeStyle := 0) {
        pBrush := this.GetSavedSolidBrush(color)
        pStrokeStyle := this.GetSavedStrokeStyle(strokeCapStyle, strokeShapeStyle)
        sRect := Buffer(64)
        NumPut("float", x, sRect, 0)
        NumPut("float", y, sRect, 4)
        NumPut("float", x + w, sRect, 8)
        NumPut("float", y + h, sRect, 12)
        this.ID2D1RenderTarget.DrawRectangle(sRect, pBrush, strokeWidth, pStrokeStyle)
    }

    FillRectangle(x, y, w, h, color) {
        pBrush := this.GetSavedSolidBrush(color)
        sRect := Buffer(64)
        NumPut("float", x, sRect, 0)
        NumPut("float", y, sRect, 4)
        NumPut("float", x + w, sRect, 8)
        NumPut("float", y + h, sRect, 12)
        this.ID2D1RenderTarget.FillRectangle(sRect, pBrush)
    }

    DrawRoundedRectangle(x, y, w, h, radiusX, radiusY, color, strokeWidth := 2, strokeCapStyle := 2, strokeShapeStyle := 0) {
        pBrush := this.GetSavedSolidBrush(color)
        pStrokeStyle := this.GetSavedStrokeStyle(strokeCapStyle, strokeShapeStyle)
        roundedRect := Buffer(64)
        NumPut("float", x, roundedRect, 0)
        NumPut("float", y, roundedRect, 4)
        NumPut("float", x + w, roundedRect, 8)
        NumPut("float", y + h, roundedRect, 12)
        NumPut("float", radiusX, roundedRect, 16)
        NumPut("float", radiusY, roundedRect, 20)
        this.ID2D1RenderTarget.DrawRoundedRectangle(roundedRect, pBrush, strokeWidth, pStrokeStyle)
    }

    FillRoundedRectangle(x, y, w, h, radiusX, radiusY, color) {
        pBrush := this.GetSavedSolidBrush(color)
        roundedRect := Buffer(64)
        NumPut("float", x, roundedRect, 0)
        NumPut("float", y, roundedRect, 4)
        NumPut("float", x + w, roundedRect, 8)
        NumPut("float", y + h, roundedRect, 12)
        NumPut("float", radiusX, roundedRect, 16)
        NumPut("float", radiusY, roundedRect, 20)
        this.ID2D1RenderTarget.FillRoundedRectangle(roundedRect, pBrush)
    }

    DrawEllipse(x, y, w, h, color, strokeWidth := 2, strokeCapStyle := 2, strokeShapeStyle := 0) {
        pBrush := this.GetSavedSolidBrush(color)
        pStrokeStyle := this.GetSavedStrokeStyle(strokeCapStyle, strokeShapeStyle)
        ellipse := Buffer(64)
        NumPut("float", x, ellipse, 0)
        NumPut("float", y, ellipse, 4)
        NumPut("float", w, ellipse, 8)
        NumPut("float", h, ellipse, 12)
        this.ID2D1RenderTarget.DrawEllipse(ellipse, pBrush, strokeWidth, pStrokeStyle)
    }

    FillEllipse(x, y, w, h, color) {
        pBrush := this.GetSavedSolidBrush(color)
        ellipse := Buffer(64)
        NumPut("float", x, ellipse, 0)
        NumPut("float", y, ellipse, 4)
        NumPut("float", w, ellipse, 8)
        NumPut("float", h, ellipse, 12)
        this.ID2D1RenderTarget.FillEllipse(ellipse, pBrush)
    }

    DrawCircle(x, y, radius, color, strokeWidth := 2, strokeCapStyle := 2, strokeShapeStyle := 0) {
        pBrush := this.GetSavedSolidBrush(color)
        pStrokeStyle := this.GetSavedStrokeStyle(strokeCapStyle, strokeShapeStyle)
        ellipse := Buffer(64)
        NumPut("float", x, ellipse, 0)
        NumPut("float", y, ellipse, 4)
        NumPut("float", radius, ellipse, 8)
        NumPut("float", radius, ellipse, 12)
        this.ID2D1RenderTarget.DrawEllipse(ellipse, pBrush, strokeWidth, pStrokeStyle)
    }

    FillCircle(x, y, radius, color) {
        pBrush := this.GetSavedSolidBrush(color)
        ellipse := Buffer(64)
        NumPut("float", x, ellipse, 0)
        NumPut("float", y, ellipse, 4)
        NumPut("float", radius, ellipse, 8)
        NumPut("float", radius, ellipse, 12)
        this.ID2D1RenderTarget.FillEllipse(ellipse, pBrush)
    }

    DrawSvg(svg, x, y, w, h) {
        local IStream := 0
        if FileExist(svg) && RegExReplace(svg, ".*\.(.*)$", "$1") = "svg" {
            DllCall("shlwapi\SHCreateStreamOnFileW", "WStr", svg, "uint", 0, "ptr*", &IStream)
        } else {
            bin := Buffer(StrPut(svg, "UTF-8"))
            cbSize := StrPut(svg, bin, "UTF-8") - 1
            hMem := DllCall("GlobalAlloc", "uint", 0x2, "uptr", cbSize, "ptr")
            pBuf := DllCall("GlobalLock", "ptr", hMem, "ptr")
            ; copy string to mem
            DllCall("RtlMoveMemory", "ptr", pBuf, "ptr", bin.Ptr, "uptr", cbSize)
            DllCall("GlobalUnlock", "ptr", hMem)
            ; convert to IStream
            DllCall("ole32\CreateStreamOnHGlobal", "ptr", hMem, "int", fDeleteOnRelease := 1, "ptr*", &IStream)
        }

        D2D1_SIZE_F := Buffer(8)
        NumPut("float", w, D2D1_SIZE_F, 0x0)
        NumPut("float", h, D2D1_SIZE_F, 0x4)
        pSvgDocument := this.ID2D1WICBitmapRenderTarget.CreateSvgDocument(IStream, NumGet(D2D1_SIZE_F, "uint64"))
        this.ID2D1WICBitmapRenderTarget.BeginDraw()
        this.ID2D1WICBitmapRenderTarget.DrawSvgDocument(pSvgDocument)
        this.ID2D1WICBitmapRenderTarget.EndDraw()

        DllCall("windowscodecs\WICConvertBitmapSource", "ptr", this.ID2D1WICBitmapRenderTarget.clsidPBGRA, "ptr", this.ID2D1WICBitmapRenderTarget.pWICBitmap, "ptr*", &pWICBitmapSource := 0, "hresult")
        d2dBmpProps := Buffer(64, 0)
        NumPut("uint", 87, d2dBmpProps, 0) ; DXGI_FORMAT_B8G8R8A8_UNORM
        NumPut("uint", 1, d2dBmpProps, 4)  ; D2D1_ALPHA_MODE_PREMULTIPLIED
        pD2DBitmap := this.ID2D1RenderTarget.CreateBitmapFromWicBitmap(pWICBitmapSource, d2dBmpProps)
        ObjRelease(IStream)
        ObjRelease(this.ID2D1WICBitmapRenderTarget.pWICBitmap)
        ObjRelease(pSvgDocument)

        dstRect := Buffer(16, 0)
        NumPut("float", x, dstRect, 0)
        NumPut("float", y, dstRect, 4)
        NumPut("float", x + w, dstRect, 8)
        NumPut("float", y + h, dstRect, 12)
        srcROI := Buffer(16, 0)
        NumPut("float", 0, srcROI, 0)
        NumPut("float", 0, srcROI, 4)
        NumPut("float", w, srcROI, 8)
        NumPut("float", h, srcROI, 12)
        this.ID2D1RenderTarget.DrawBitmap(pD2DBitmap, dstRect, opacity := 1, linear := 1, srcROI)
    }

    DrawImage(imgPath, x := 0, y := 0, w := 0, h := 0, opacity := 1) {
        dstRect := Buffer(16, 0)
        NumPut("float", x, dstRect, 0)
        NumPut("float", y, dstRect, 4)
        NumPut("float", x + w, dstRect, 8)
        NumPut("float", y + h, dstRect, 12)
        srcROI := Buffer(16, 0)
        NumPut("float", 0, srcROI, 0)
        NumPut("float", 0, srcROI, 4)
        NumPut("float", w, srcROI, 8)
        NumPut("float", h, srcROI, 12)
        if pBitmap := this.GetSavedBitmapFromWicBitmap(imgPath)
            this.ID2D1RenderTarget.DrawBitmap(pBitmap, dstRect, opacity, linear := 1, srcROI)
    }

    GetSavedBitmapFromWicBitmap(imgPath) {
        if (this.d2dBitmaps.has(imgPath))
            return this.d2dBitmaps[imgPath]

        if (!FileExist(imgPath)) {
            MsgBox(Format("{} does not exist!", imgPath))
            return 0
        }
        DllCall("gdiplus\GdipCreateBitmapFromFile", "Str", imgPath, "Ptr*", &pGdiBitmap := 0)
        DllCall("gdiplus\GdipGetImageWidth", "ptr", pGdiBitmap, "uint*", &imgW := 0)
        DllCall("gdiplus\GdipGetImageHeight", "ptr", pGdiBitmap, "uint*", &imgH := 0)
        pWICImagingFactory := ComObject("{CACAF262-9370-4615-A13B-9F5539DA4C0A}", "{EC5EC8A9-C395-4314-9C77-54D7A935FF70}")
        Direct2D.getCLSID("{6FDDC324-4E03-4BFE-B185-3D77768DC90F}", &clsidBGRA)
        ComCall(CreateBitmap := 17, pWICImagingFactory, "uint", imgW, "uint", imgH, "ptr", clsidBGRA, "int", 1, "ptr*", &pWICBitmap := 0)

        ; convert GdiBitmap to WICBitmap
        rect := Buffer(16, 0)
        NumPut("uint", imgW, rect, 8)
        NumPut("uint", imgH, rect, 12)
        ComCall(Lock := 8, pWICBitmap, "ptr", rect, "uint", 0x2, "ptr*", &pWICBitmapLock := 0)
        ComCall(GetDataPointer := 5, pWICBitmapLock, "uint*", &size := 0, "ptr*", &pWICBmpData := 0)
        ; copy GdiBitmap data to WICBitmap data
        pBitmapData := Buffer(16 + 2 * A_PtrSize, 0)
        NumPut("int", 4 * imgW, pBitmapData, 8) ; stride
        NumPut("ptr", pWICBmpData, pBitmapData, 16) ; scan0
        DllCall("gdiplus\GdipBitmapLockBits", "ptr", pGdiBitmap, "ptr", rect,
            "uint", 5, ; UserInputBuffer | ReadOnly
            "int", 0xE200B, ; Format32bppPArgb
            "ptr", pBitmapData)
        DllCall("gdiplus\GdipBitmapUnlockBits", "ptr", pGdiBitmap, "ptr", pBitmapData)
        ObjRelease(pWICBitmapLock)

        ; convert WICBitmap to WICBitmapSource for CreateBitmapFromWicBitmap
        DllCall("windowscodecs\WICConvertBitmapSource", "ptr", clsidBGRA, "ptr", pWICBitmap, "ptr*", &pWICBitmapSource := 0, "hresult")
        d2dBmpProps := Buffer(64, 0)
        NumPut("uint", 87, d2dBmpProps, 0) ; DXGI_FORMAT_B8G8R8A8_UNORM
        NumPut("uint", 1, d2dBmpProps, 4)  ; D2D1_ALPHA_MODE_PREMULTIPLIED
        return this.d2dBitmaps[imgPath] := this.ID2D1RenderTarget.CreateBitmapFromWicBitmap(pWICBitmapSource, d2dBmpProps)
    }

    GetSavedTextFormat(fontName, fontSize) {
        fK := Format("{}_{}", fontName, fontSize)
        if this.textFormats.Has(fK)
            return this.textFormats[fK]

        return this.textFormats[fK] := Direct2D.IDWriteFactory.CreateTextFormat(fontName, fontSize)
    }

    GetSavedSolidBrush(c) {
        bK := Format("{}", c)
        if this.solidBrushes.Has(bK)
            return this.solidBrushes[bK]

        if c <= 0xFFFFFF
            c := c | 0xFF000000
        sColor := Buffer(16, 0)
        NumPut("Float", ((c & 0xFF0000) >> 16) / 255, sColor, 0)  ; R
        NumPut("Float", ((c & 0xFF00) >> 8) / 255, sColor, 4)  ; G
        NumPut("Float", ((c & 0xFF)) / 255, sColor, 8)  ; B
        NumPut("Float", (c > 0xFFFFFF ? ((c & 0xFF000000) >> 24) / 255 : 1), sColor, 12) ; A
        pBrush := this.ID2D1RenderTarget.CreateSolidBrush(sColor)
        return this.solidBrushes[bK] := pBrush
    }

    /**
     * @see https://learn.microsoft.com/windows/win32/api/d2d1/ns-d2d1-d2d1_stroke_style_properties
     * @param {Integer} capStyle flat(0) square(1) round(2) triangle(3)
     * @param {Integer} shapeStyle solid(0) dash(1) dot(2) dash_dot(3)
     * @returns {Integer} pStrokeStyle
     */
    GetSavedStrokeStyle(capStyle, shapeStyle := 0) {
        sK := Format("{}_{}", capStyle, shapeStyle)
        if this.strokeStyles.Has(sK)
            return this.strokeStyles[sK]

        styleProps := Buffer(A_PtrSize * 7, 0)
        NumPut("UInt", capStyle, styleProps, 0) ; startCap: D2D1_CAP_STYLE_ROUND(2)
        NumPut("UInt", capStyle, styleProps, 4) ; endCap: D2D1_CAP_STYLE_ROUND(2)
        NumPut("UInt", capStyle, styleProps, 8) ; dashCap: D2D1_CAP_STYLE_ROUND(2)
        NumPut("UInt", capStyle, styleProps, 12) ; lineJoin: D2D1_LINE_JOIN_ROUND(2)
        NumPut("Float", 10.0, styleProps, 16) ; miterLimit
        NumPut("UInt", shapeStyle, styleProps, 20) ; dashStyle: SOLID(0) DASH(1) DOT(2) DASHDOT(3) DASHDOTDOT(4)
        NumPut("Float", -1.0, styleProps, 24) ; dashOffset
        return this.strokeStyles[sK] := Direct2D.ID2D1Factory.CreateStrokeStyle(styleProps)
    }

    SetPosition(x, y, w := 0, h := 0) {
        if !this.hwnd
            return

        this.x := x, this.y := y
        if w != 0 and h != 0 {
            newSize := Buffer(16, 0)
            NumPut("uint", this.width := w, newSize, 0)
            NumPut("uint", this.height := h, newSize, 4)
            this.ID2D1RenderTarget.Resize(newSize)
        }
        DllCall("MoveWindow", "Uptr", this.hwnd, "int", x, "int", y, "int", this.width, "int", this.height, "char", 1)
    }

    GdipStartUp() {
        if !DllCall("GetModuleHandleW", "Wstr", "gdiplus", "UPtr")
            DllCall("LoadLibraryW", "Wstr", "gdiplus", "UPtr")

        gdipStartupInput := Buffer(A_PtrSize == 8 ? 24 : 16, 0), NumPut("UInt", Gdip := 1, gdipStartupInput, 0)
        DllCall("gdiplus\GdiplusStartup", "Ptr*", &pToken := 0, "Ptr", gdipStartupInput, "Ptr", 0)
        return pToken
    }

    GdipShutdown(pToken) {
        DllCall("gdiplus\GdiplusShutdown", "Ptr", pToken)
        if hModule := DllCall("GetModuleHandleW", "Wstr", "gdiplus", "UPtr")
            DllCall("FreeLibrary", "Ptr", hModule)
    }
}
