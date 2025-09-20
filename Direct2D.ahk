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
 * @version 0.1.4
 * @example
 * ui := Gui("-DPIScale")
 * d2d := Direct2D()
 * d2d.SetTargetRender(ui.Hwnd, 512, 512)
 * ui.Show("W512 H512")
 * d2d.BeginDraw()
 * d2d.DrawSvg('<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" ...</svg>', 10, 10, 128, 128)
 * d2d.EndDraw()
 ******************************************************************************/
class Direct2D {
    __New(target?) {
        this.gdipToken := this.GdipStartUp() ; to get image bitmap info

        this.textFormats := Map()
        this.solidBrushes := Map()
        this.strokeStyles := Map()
        this.d2dBitmaps := Map()
        this.drawBounds := Buffer(24, 0) ; for D2D1_RECT D2D1_ROUNDED_RECT D2D1_ELLIPSE
        this.isDrawing := 0

        if isSet(target)
            this.SetRenderTarget(target)
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

        OnMessage(WM_ERASEBKGND := 0x14, (wParam, lParam, msg, hwnd) => this.hwnd == hwnd ? 0 : "", 0)
    }

    static isX64 := A_PtrSize == 8
    static vTable(p, i) => (v := NumGet(p + 0, 0, "ptr")) ? NumGet(v + 0, i * A_PtrSize, "Ptr") : 0
    static release(p) => (r := this.vTable(p, 2)) ? DllCall(r, "ptr", p) : 0
    static getCLSID(guid, &clsid) => DllCall("ole32\CLSIDFromString", "WStr", guid, "Ptr", clsid := Buffer(16, 0))
    static fnv1aHash(str) {
        hash := 0x811C9DC5
        prime := 0x01000193
        buf := StrPtr(str)
        len := StrLen(str)
        Loop len {
            c := NumGet(buf, (A_Index - 1) * 2, "UShort")
            hash := (hash ^ c) * prime & 0xFFFFFFFF
        }
        return Format("{:08x}", hash)
    }

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

    GetDesktopDpiScale() {
        dpiX := Direct2D.ID2D1Factory.GetDesktopDpi()
        return dpiX / 96
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
        pTextFormat := this.GetSavedOrCreateTextFormat(fontName, fontSize)
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
        static textMetrics := Buffer(4 * 9)
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

    class ID2D1RenderTarget {
        __New(target, width, height) {
            this.pRT := 0
            this.width := width, this.height := height
            static rtProps := Buffer(64, 0)
            NumPut("uint", 1, rtProps, 8) ; AlphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED
            NumPut("float", 96, rtProps, 12) ; dpiX
            NumPut("float", 96, rtProps, 16) ; dpiY
            static hRtProps := Buffer(64, 0)
            NumPut("Uptr", target, hRtProps, 0)
            NumPut("uint", width, hRtProps, A_PtrSize)
            NumPut("uint", height, hRtProps, A_PtrSize + 4)
            NumPut("uint", 2, hRtProps, A_PtrSize + 8)

            ; set window visible
            DllCall("SetLayeredWindowAttributes", "Uptr", target, "Uint", ColorKey := 0, "char", Alpha := 255, "uint", LWA_ALPHA := 2)

            static margins := Buffer(16, 0)
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
                Direct2D.release(this.pRT), this.pRT := 0
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
                static D2D1_SIZE_U := Buffer(16)
                NumPut("uint", w, D2D1_SIZE_U, 0)
                NumPut("uint", h, D2D1_SIZE_U, 4)
                DllCall(this.VT_CreateBitmap, "Ptr", this.pRT, "int64", NumGet(D2D1_SIZE_U, 0, "int64"), "Ptr", srcData, "uint", pitch, "Ptr", bmpProps, "Ptr*", &pBitmap)
            } else {
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
                static points := Buffer(16)
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
                static D2D1_POINT_2F := Buffer(8)
                NumPut('float', point[1], D2D1_POINT_2F, 0)
                NumPut('float', point[2], D2D1_POINT_2F, 4)
                DllCall(this.VT_DrawTextLayout, "ptr", this.pRT, 'double', NumGet(D2D1_POINT_2F, 0, 'double'), "ptr", pTextLayout, 'ptr', pBrush, "uint", drawOpt)
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
        ; WicBitmapRenderTarget needs pixelFormat32bppPBGRA
        __New(imgW, imgH, pixelFormatGUID := "{6fddc324-4e03-4bfe-b185-3d77768dc910}") {
            this.width := imgW, this.height := imgH
            static rtProps := Buffer(64, 0)
            NumPut("uint", 1, rtProps, 0) ; D2D1_RENDER_TARGET_TYPE_SOFTWARE
            NumPut("uint", 87, rtProps, 4) ; DXGI_FORMAT_B8G8R8A8_UNORM
            NumPut("uint", 1, rtProps, 8) ; AlphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED
            ; Initialize Windows Imaging Component.
            CLSID_WICImagingFactory := "{CACAF262-9370-4615-A13B-9F5539DA4C0A}"
            IID_IWICImagingFactory := "{EC5EC8A9-C395-4314-9C77-54D7A935FF70}"
            pWICImagingFactory := ComObject(CLSID_WICImagingFactory, IID_IWICImagingFactory)
            Direct2D.getCLSID(pixelFormatGUID, &clsid)
            ComCall(CreateBitmap := 17, pWICImagingFactory, ; IWICImagingFactory::CreateBitmap in memory
                "uint", this.width, "uint", this.height,
                "ptr", this.pixelFormatCLSID := clsid,
                "uint", WICBitmapCacheOnDemand := 1,
                "ptr*", &pWICBitmap := 0)
            this.pWICBitmap := pWICBitmap

            if pixelFormatGUID == "{6FDDC324-4E03-4BFE-B185-3D77768DC90F}" ; pixelFormat32bppBGRA
                return 0

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
            ; pRT will be released in ID2D1RenderTarget

            if this.pWICBitmap
                Direct2D.release(this.pWICBitmap)
        }

        GetWICBitmap() => this.pWICBitmap

        WICBitmapToWICBitmapSource() => (DllCall("windowscodecs\WICConvertBitmapSource", "ptr", this.pixelFormatCLSID, "ptr", this.pWICBitmap, "ptr*", &pWICBitmapSource := 0), pWICBitmapSource)

        CreateSvgDocument(fs, w, h) {
            local pSvgDocument := 0
            if Direct2D.isX64 {
                static D2D1_SIZE_F := Buffer(8)
                NumPut("float", w, D2D1_SIZE_F, 0)
                NumPut("float", h, D2D1_SIZE_F, 4)
                DllCall(this.VT_CreateSvgDocument, "Ptr", this.pRT, "ptr", fs, "uint64", NumGet(D2D1_SIZE_F, "uint64"), "ptr*", &pSvgDocument)
            } else {
                DllCall(this.VT_CreateSvgDocument, "Ptr", this.pRT, "ptr", fs, "float", w, "float", h, "ptr*", &pSvgDocument)
            }
            return pSvgDocument
        }

        DrawSvgDocument(pSvgDocument) => DllCall(this.VT_DrawSvgDocument, "Ptr", this.pRT, "Ptr", pSvgDocument)

        DrawSvgWICBitmap(fs) {
            pSvgDocument := this.CreateSvgDocument(fs, this.width, this.height) ; ID2D1SvgDocument

            this.BeginDraw()
            this.DrawSvgDocument(pSvgDocument)
            this.EndDraw()

            ObjRelease(fs)
            Direct2D.release(pSvgDocument)
        }

        GetSvgHBitmap(fs) {
            this.DrawSvgWICBitmap(fs)
            return this.GetHBitmapFromWICBitmap()
        }

        GetSvgWICBitmapSource(fs) {
            this.DrawSvgWICBitmap(fs)
            return this.WICBitmapToWICBitmapSource()
        }

        GetHBitmapFromWICBitmap() {
            stride := 4 * this.width
            pData := Buffer(stride * this.height)
            ComCall(CopyPixels := 7, this.pWICBitmap, "ptr", 0, "uint", stride, "uint", pData.Size, "ptr", pData)
            hBitmap := DllCall("gdi32\CreateBitmap", "Int", this.width, "Int", this.height, "Uint", 1, "Uint", 32, "Ptr", pData, "Ptr")

            Direct2D.release(this.pWICBitmap), this.pWICBitmap := 0
            return hBitmap
        }

        GdiBitmapToWICBitmapSource(pGdiBitmap, gdiPixelFormat) {
            if !pGdiBitmap
                return 0

            ; convert GdiBitmap to WICBitmap
            static sRect := Buffer(16, 0)
            NumPut("uint", this.width, sRect, 8)
            NumPut("uint", this.height, sRect, 12)
            ComCall(Lock := 8, this.pWICBitmap, "ptr", sRect, "uint", 0x2, "ptr*", &pWICBitmapLock := 0)
            ComCall(GetDataPointer := 5, pWICBitmapLock, "uint*", &size := 0, "ptr*", &pWICBmpData := 0)
            ; copy GdiBitmap data to WICBitmap data
            static pBitmapData := Buffer(16 + 2 * A_PtrSize, 0)
            NumPut("int", 4 * this.width, pBitmapData, 8) ; stride
            NumPut("ptr", pWICBmpData, pBitmapData, 16) ; scan0
            DllCall("gdiplus\GdipBitmapLockBits", "ptr", pGdiBitmap, "ptr", sRect,
                "uint", 5, ; UserInputBuffer | ReadOnly
                "int", gdiPixelFormat, ; Format32bppPArgb:0xE200B
                "ptr", pBitmapData)
            DllCall("gdiplus\GdipBitmapUnlockBits", "ptr", pGdiBitmap, "ptr", pBitmapData)
            ObjRelease(pWICBitmapLock)

            return this.WICBitmapToWICBitmapSource()
        }
    }

    class ID2D1DCRenderTarget extends Direct2D.ID2D1RenderTarget {
        __New() {
            static rtProps := Buffer(64, 0)
            NumPut("uint", 87, rtProps, 4) ; DXGI_FORMAT_B8G8R8A8_UNORM
            NumPut("uint", 1, rtProps, 8) ; AlphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED
            this.pRT := Direct2D.ID2D1Factory.CreateDCRenderTarget(rtProps)
            if !this.pRT {
                MsgBox("ID2D1DCRenderTarget init failed")
                return 0
            }

            super.__InitCommonVT()
            this.VT_BindDC := Direct2D.vTable(this.pRT, 57)

            return this.pRT
        }

        BindDC(hDC, rect) => DllCall(this.VT_BindDC, "Ptr", this.pRT, "ptr", hDC, "ptr", rect)
    }

    /**
     * @description ID2D1DCRenderTarget test
     * @example
     * d2d := Direct2D()
     * dcRT := d2d.SetRenderTarget("dc", 500, 500)
     * dcGui := Direct2D.DCGui(dcRT, 500, 500)
     * d2d.BeginDraw()
     * d2d.FillRoundedRectangle(0, 0, 500, 500, 6, 6, 0xFFFF2222)
     * d2d.EndDraw()
     * dcGui.Show(100, 100)
     */
    class DCGui extends Gui {
        __New(dcRt, w, h) {
            super.__New("-DPIScale -Caption +E0x80000 +E0x08000000")
            this.x := 0, this.y := 0, this.width := w, this.height := h
            this.hDC := DllCall('CreateCompatibleDC', 'ptr', 0, 'ptr')
            rect := Buffer(16), NumPut("uint", w, rect, 8), NumPut("uint", h, rect, 12)
            dcRt.BindDC(this.hDC, rect)
            bitmapInfo := Buffer(40)
            NumPut('uint', 40, 'uint', w, 'uint', h, 'ushort', 1, 'ushort', 32, 'uint', 0, bitmapInfo)
            this.hBitmap := DllCall('CreateDIBSection', 'ptr', this.hDC, 'ptr', bitmapInfo, 'uint', 0, 'ptr*', &ppvBits := 0, 'ptr', 0, 'uint', 0, 'ptr')
            this.oBitmap := DllCall('SelectObject', 'ptr', this.hDC, 'ptr', this.hBitmap, 'ptr')
        }

        __Delete() {
            DllCall('SelectObject', 'ptr', this.hDC, 'ptr', this.oBitmap, 'ptr')
            DllCall('DeleteObject', 'ptr', this.hBitmap)
            DllCall('DeleteDC', 'ptr', this.hDC)
        }

        Show(x := 0, y := 0) {
            super.Show(Format("x{} y{} W{} H{}", this.x := x, this.y := y, this.width, this.height))
            this.UpdateLayeredWindow(this.x, this.y, this.width, this.height, alpha := 255)
        }

        UpdateLayeredWindow(x, y, w, h, alpha?) {
            point := Buffer(8)
            NumPut("UInt", x, "UInt", y, point)
            DllCall("UpdateLayeredWindow",
                "UPtr", this.Hwnd, "UPtr", 0,
                "UPtr", (!x && !y) ? 0 : point.Ptr,
                "Int64*", w | h << 32,
                "UPtr", this.hDC, "Int64*", 0, "UInt", 0, "UInt*", Alpha << 16 | 1 << 24, "UInt", 2)
        }
    }

    /**
     * @param {Integer | String} target GuiObj or GuiHwnd for hwndTraget, "wic" for WicBitmapTarget
     * @param {Integer} w GuiClient or WicBitmap width
     * @param {Integer} h GuiClient or WicBitmap height
     * @returns {Direct2D.ID2D1RenderTarget | Direct2D.ID2D1WicBitmapRenderTarget}
     */
    SetRenderTarget(target, w := 0, h := 0) {
        this.target := target, this.hwnd := 0, this.attachHwnd := 0
        this.x := 0, this.y := 0, this.width := w, this.height := h
        this.winRect := []
        if target && IsInteger(target) {
            this.hwnd := target
            this.ID2D1RenderTarget := Direct2D.ID2D1RenderTarget(this.hwnd, this.width, this.height)
        } else if target is String {
            if target = "wic" {
                if w == 0 || h == 0 {
                    MsgBox("WicBitmapRenderTarget needs width and height for a image!")
                    return 0
                }
                this.ID2D1RenderTarget := Direct2D.ID2D1WicBitmapRenderTarget(this.width, this.height)
            } else if target = "dc" { ; gdip DC
                this.ID2D1RenderTarget := Direct2D.ID2D1DCRenderTarget()
            } else { ; attach to a window
                this.lastSize := 0, this.lastPos := 0
                if this.attachHwnd := WinExist(target) {
                    this.gui := Gui("-DPIScale -Caption +E0x80800A8")
                    this.hwnd := this.gui.Hwnd
                    this.ID2D1RenderTarget := Direct2D.ID2D1RenderTarget(this.hwnd, this.width, this.height)
                } else {
                    MsgBox(Format('WinTitle "{}" dose not exist', target))
                    return 0
                }
            }
        } else if target is Gui {
            this.hwnd := target.Hwnd
            this.ID2D1RenderTarget := Direct2D.ID2D1RenderTarget(this.hwnd, this.width, this.height)
        } else {
            MsgBox("Unsupported target!")
            return 0
        }

        OnMessage(WM_ERASEBKGND := 0x14, (wParam, lParam, msg, hwnd) => this.hwnd == hwnd ? 0 : "")

        this.ID2D1RenderTarget.SetAntiAliasMode(D2D1_ANTIALIAS_MODE_PER_PRIMITIVE := 0)
        this.ID2D1RenderTarget.SetTextAntiAliasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE := 1)
        this.Clear()

        return this.ID2D1RenderTarget
    }

    UpdatePosition() {
        static sRect := Buffer(64, 0)
        if this.attachHwnd {
            hasWinRect := DllCall("GetWindowInfo", "Uptr", this.attachHWND, "ptr", sRect)
            onActiveWin := DllCall("GetForegroundWindow", "cdecl Ptr") == this.attachHwnd
            if !hasWinRect or !onActiveWin {
                return attachFailed()
            } else if (!this.isDrawing) {
                DllCall("ShowWindow", "Ptr", this.hwnd, "Int", 4)
            }
        } else {
            if !DllCall("GetWindowInfo", "Uptr", this.hwnd, "ptr", sRect) {
                return attachFailed()
            }
        }
        x := NumGet(sRect, 4, "int"), y := NumGet(sRect, 8, "int")
        right := NumGet(sRect, 12, "int"), bottom := NumGet(sRect, 16, "int")
        w := right - x, h := bottom - y
        this.winRect := [x, y, right, bottom]
        cxWindowBorders := NumGet(sRect, 48, "int"), cyWindowBorders := NumGet(sRect, 52, "int")
        sizeCode := (w << 16) + h, posCode := (x << 16) + y
        if (sizeCode != this.lastSize) {
            this.SetPosition(x + cyWindowBorders, y, w - cyWindowBorders * 2, h - cxWindowBorders)
            this.lastSize := sizeCode, this.lastPos := posCode
        } else if (posCode != this.lastPos) {
            this.SetPosition(x + cyWindowBorders, y)
            this.lastPos := posCode
        }
        return 1

        attachFailed() {
            if (this.isDrawing) {
                this.Clear()
                this.isDrawing := 0
                DllCall("ShowWindow", "Ptr", this.hwnd, "Int", 0)
            }
            return 0
        }
    }

    BeginDraw() {
        if this.attachHwnd {
            if !this.UpdatePosition()
                return 0
        }

        this.ID2D1RenderTarget.BeginDraw()
        this.ID2D1RenderTarget.Clear()
        return this.isDrawing := 1
    }

    EndDraw() {
        if (this.isDrawing)
            this.ID2D1RenderTarget.EndDraw()
    }

    Clear() {
        this.ID2D1RenderTarget.BeginDraw()
        this.ID2D1RenderTarget.Clear()
        this.ID2D1RenderTarget.EndDraw()
    }

    ResizeRenderTarget(w, h) {
        static D2D1_SIZE_U := Buffer(16, 0)
        NumPut("uint", w, D2D1_SIZE_U, 0)
        NumPut("uint", h, D2D1_SIZE_U, 4)
        this.ID2D1RenderTarget.Resize(D2D1_SIZE_U)
    }

    DrawText(text, x, y, fontSize, color, fontName, w?, h?, drawOpt := 4) {
        pTextFormat := this.GetSavedOrCreateTextFormat(fontName, fontSize)
        pBrush := this.GetSavedOrCreateSolidBrush(color)
        NumPut("float", x, this.drawBounds, 0)
        NumPut("float", y, this.drawBounds, 4)
        NumPut("float", x + (w ?? this.width), this.drawBounds, 8)
        NumPut("float", y + (h ?? this.height), this.drawBounds, 12)

        ; https://learn.microsoft.com/windows/win32/api/d2d1/ne-d2d1-d2d1_draw_text_options
        ; typedef enum D2D1_DRAW_TEXT_OPTIONS {
        ;   D2D1_DRAW_TEXT_OPTIONS_NO_SNAP = 0x00000001,
        ;   D2D1_DRAW_TEXT_OPTIONS_CLIP = 0x00000002,
        ;   D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT = 0x00000004,
        ;   D2D1_DRAW_TEXT_OPTIONS_DISABLE_COLOR_BITMAP_SNAPPING = 0x00000008,
        ;   D2D1_DRAW_TEXT_OPTIONS_NONE = 0x00000000,
        ;   D2D1_DRAW_TEXT_OPTIONS_FORCE_DWORD = 0xffffffff
        ; }
        this.ID2D1RenderTarget.DrawText(text, pTextFormat, this.drawBounds, pBrush, drawOpt)
    }

    DrawTextLayout(text, x, y, color, pTextLayout, drawOpt := 4) {
        if !text
            return

        pBrush := this.GetSavedOrCreateSolidBrush(color)
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
        pBrush := this.GetSavedOrCreateSolidBrush(color)
        pStrokeStyle := this.GetSavedOrCreateStrokeStyle(strokeCapStyle, strokeShapeStyle)
        this.ID2D1RenderTarget.DrawLine(pointStart, pointEnd, pBrush, strokeWidth, pStrokeStyle)
    }

    DrawRectangle(x, y, w, h, color, strokeWidth := 2, strokeCapStyle := 0, strokeShapeStyle := 0) {
        pBrush := this.GetSavedOrCreateSolidBrush(color)
        pStrokeStyle := this.GetSavedOrCreateStrokeStyle(strokeCapStyle, strokeShapeStyle)
        NumPut("float", x, this.drawBounds, 0)
        NumPut("float", y, this.drawBounds, 4)
        NumPut("float", x + w, this.drawBounds, 8)
        NumPut("float", y + h, this.drawBounds, 12)
        this.ID2D1RenderTarget.DrawRectangle(this.drawBounds, pBrush, strokeWidth, pStrokeStyle)
    }

    FillRectangle(x, y, w, h, color) {
        pBrush := this.GetSavedOrCreateSolidBrush(color)
        NumPut("float", x, this.drawBounds, 0)
        NumPut("float", y, this.drawBounds, 4)
        NumPut("float", x + w, this.drawBounds, 8)
        NumPut("float", y + h, this.drawBounds, 12)
        this.ID2D1RenderTarget.FillRectangle(this.drawBounds, pBrush)
    }

    DrawRoundedRectangle(x, y, w, h, radiusX, radiusY, color, strokeWidth := 2, strokeCapStyle := 2, strokeShapeStyle := 0) {
        pBrush := this.GetSavedOrCreateSolidBrush(color)
        pStrokeStyle := this.GetSavedOrCreateStrokeStyle(strokeCapStyle, strokeShapeStyle)
        NumPut("float", x, this.drawBounds, 0)
        NumPut("float", y, this.drawBounds, 4)
        NumPut("float", x + w, this.drawBounds, 8)
        NumPut("float", y + h, this.drawBounds, 12)
        NumPut("float", radiusX, this.drawBounds, 16)
        NumPut("float", radiusY, this.drawBounds, 20)
        this.ID2D1RenderTarget.DrawRoundedRectangle(this.drawBounds, pBrush, strokeWidth, pStrokeStyle)
    }

    FillRoundedRectangle(x, y, w, h, radiusX, radiusY, color) {
        pBrush := this.GetSavedOrCreateSolidBrush(color)
        NumPut("float", x, this.drawBounds, 0)
        NumPut("float", y, this.drawBounds, 4)
        NumPut("float", x + w, this.drawBounds, 8)
        NumPut("float", y + h, this.drawBounds, 12)
        NumPut("float", radiusX, this.drawBounds, 16)
        NumPut("float", radiusY, this.drawBounds, 20)
        this.ID2D1RenderTarget.FillRoundedRectangle(this.drawBounds, pBrush)
    }

    DrawEllipse(x, y, w, h, color, strokeWidth := 2, strokeCapStyle := 2, strokeShapeStyle := 0) {
        pBrush := this.GetSavedOrCreateSolidBrush(color)
        pStrokeStyle := this.GetSavedOrCreateStrokeStyle(strokeCapStyle, strokeShapeStyle)
        NumPut("float", x, this.drawBounds, 0)
        NumPut("float", y, this.drawBounds, 4)
        NumPut("float", w, this.drawBounds, 8)
        NumPut("float", h, this.drawBounds, 12)
        this.ID2D1RenderTarget.DrawEllipse(this.drawBounds, pBrush, strokeWidth, pStrokeStyle)
    }

    FillEllipse(x, y, w, h, color) {
        pBrush := this.GetSavedOrCreateSolidBrush(color)
        NumPut("float", x, this.drawBounds, 0)
        NumPut("float", y, this.drawBounds, 4)
        NumPut("float", w, this.drawBounds, 8)
        NumPut("float", h, this.drawBounds, 12)
        this.ID2D1RenderTarget.FillEllipse(this.drawBounds, pBrush)
    }

    DrawCircle(x, y, radius, color, strokeWidth := 2, strokeCapStyle := 2, strokeShapeStyle := 0) {
        pBrush := this.GetSavedOrCreateSolidBrush(color)
        pStrokeStyle := this.GetSavedOrCreateStrokeStyle(strokeCapStyle, strokeShapeStyle)
        NumPut("float", x, this.drawBounds, 0)
        NumPut("float", y, this.drawBounds, 4)
        NumPut("float", radius, this.drawBounds, 8)
        NumPut("float", radius, this.drawBounds, 12)
        this.ID2D1RenderTarget.DrawEllipse(this.drawBounds, pBrush, strokeWidth, pStrokeStyle)
    }

    FillCircle(x, y, radius, color) {
        pBrush := this.GetSavedOrCreateSolidBrush(color)
        NumPut("float", x, this.drawBounds, 0)
        NumPut("float", y, this.drawBounds, 4)
        NumPut("float", radius, this.drawBounds, 8)
        NumPut("float", radius, this.drawBounds, 12)
        this.ID2D1RenderTarget.FillEllipse(this.drawBounds, pBrush)
    }

    /**
     * @param {String} svg svg content or svg file path
     * @param {Integer} x svg x of leftTop point
     * @param {Integer} y svg y of leftTop point
     * @param {Integer} w svg width
     * @param {Integer} h svg height
     */
    DrawSvg(svg, x, y, w, h) {
        if FileExist(svg) && RegExReplace(svg, ".*\.(.*)$", "$1") = "svg" {
            DllCall("shlwapi\SHCreateStreamOnFileW", "WStr", svg, "uint", 0, "ptr*", &fs := 0)
            ; convert stream to svgStr for changing width and height
            DllCall("shlwapi\IStream_Size", "ptr", fs, "uint64*", &size := 0)
            DllCall("shlwapi\IStream_Reset", "ptr", fs)
            DllCall("shlwapi\IStream_Read", "ptr", fs, "ptr", buf := Buffer(size), "uint", size, "hresult")
            svg := StrGet(buf, size, "UTF-8")
            ObjRelease(fs)
        }

        if RegExMatch(svg, 'i)^\s*<svg\b[^>]*>') {
            ; replace or insert svg new width and height
            svg := RegExReplace(svg, 'i)(<svg\b[^>]*?)\bwidth\s*=\s*"(?:[^"]*)"', '${1} width="' . w . '"')
            svg := RegExReplace(svg, 'i)(<svg\b[^>]*?)\bheight\s*=\s*"(?:[^"]*)"', '${1} height="' . h . '"')
            if !RegExMatch(svg, 'i)<svg\b[^>]*\bwidth\s*=')
                svg := RegExReplace(svg, 'i)<svg\b', '<svg width="' . w . '"')
            if !RegExMatch(svg, 'i)<svg\b[^>]*\bheight\s*=')
                svg := RegExReplace(svg, 'i)(<svg\b[^>]*?)\b', '${1} height="' . h . '"')
        } else {
            MsgBox("Invailed svg content")
            return 0
        }

        static dstRect := Buffer(16, 0)
        NumPut("float", x, dstRect, 0)
        NumPut("float", y, dstRect, 4)
        NumPut("float", x + w, dstRect, 8)
        NumPut("float", y + h, dstRect, 12)
        static srcROI := Buffer(16, 0)
        NumPut("float", 0, srcROI, 0)
        NumPut("float", 0, srcROI, 4)
        NumPut("float", w, srcROI, 8)
        NumPut("float", h, srcROI, 12)
        if pD2dBitmap := this.GetSavedOrCreateSvgBitmap(svg, w, h)
            this.ID2D1RenderTarget.DrawBitmap(pD2dBitmap, dstRect, opacity := 1, linear := 1, srcROI)
    }

    DrawImage(imgPath, x := 0, y := 0, w := 0, h := 0, opacity := 1) {
        static dstRect := Buffer(16, 0)
        NumPut("float", x, dstRect, 0)
        NumPut("float", y, dstRect, 4)
        NumPut("float", x + w, dstRect, 8)
        NumPut("float", y + h, dstRect, 12)
        static srcROI := Buffer(16, 0)
        NumPut("float", 0, srcROI, 0)
        NumPut("float", 0, srcROI, 4)
        NumPut("float", w, srcROI, 8)
        NumPut("float", h, srcROI, 12)
        if pD2dBitmap := this.GetSavedOrCreateImgBitmap(imgPath)
            this.ID2D1RenderTarget.DrawBitmap(pD2dBitmap, dstRect, opacity, linear := 1, srcROI)
    }

    GetSavedOrCreateSvgBitmap(svgStr, w, h) {
        svgId := Direct2D.fnv1aHash(svgStr)
        if (this.d2dBitmaps.has(svgId))
            return this.d2dBitmaps[svgId]

        ; put newDim svg to stream
        bin := Buffer(StrPut(svgStr, "UTF-8"))
        cbSize := StrPut(svgStr, bin, "UTF-8") - 1
        hMem := DllCall("GlobalAlloc", "uint", 0x2, "uptr", cbSize, "ptr")
        pBuf := DllCall("GlobalLock", "ptr", hMem, "ptr")
        DllCall("RtlMoveMemory", "ptr", pBuf, "ptr", bin.Ptr, "uptr", cbSize)
        DllCall("GlobalUnlock", "ptr", hMem)
        DllCall("ole32\CreateStreamOnHGlobal", "ptr", hMem, "int", fDeleteOnRelease := 1, "ptr*", &svgStream := 0)

        wicBitmapRT := Direct2D.ID2D1WicBitmapRenderTarget(w, h)
        pWICBitmapSource := wicBitmapRT.GetSvgWICBitmapSource(svgStream)
        static d2dBmpProps := Buffer(64, 0)
        NumPut("uint", 87, d2dBmpProps, 0) ; DXGI_FORMAT_B8G8R8A8_UNORM
        NumPut("uint", 1, d2dBmpProps, 4)  ; D2D1_ALPHA_MODE_PREMULTIPLIED
        return this.d2dBitmaps[svgId] := this.ID2D1RenderTarget.CreateBitmapFromWicBitmap(pWICBitmapSource, d2dBmpProps)
    }

    GetSavedOrCreateImgBitmap(imgPath) {
        if (this.d2dBitmaps.has(imgPath))
            return this.d2dBitmaps[imgPath]

        if (!FileExist(imgPath)) {
            MsgBox(Format("{} does not exist!", imgPath))
            return 0
        }
        DllCall("gdiplus\GdipCreateBitmapFromFile", "Str", imgPath, "Ptr*", &pGdiBitmap := 0)
        DllCall("gdiplus\GdipGetImageWidth", "ptr", pGdiBitmap, "uint*", &imgW := 0)
        DllCall("gdiplus\GdipGetImageHeight", "ptr", pGdiBitmap, "uint*", &imgH := 0)
        wicBitmap := Direct2D.ID2D1WicBitmapRenderTarget(imgW, imgH, pixelFormat32bppBGRA := "{6FDDC324-4E03-4BFE-B185-3D77768DC90F}")
        pWICBitmapSource := wicBitmap.GdiBitmapToWICBitmapSource(pGdiBitmap, Format32bppPArgb := 0xE200B)
        static d2dBmpProps := Buffer(64, 0)
        NumPut("uint", 87, d2dBmpProps, 0) ; DXGI_FORMAT_B8G8R8A8_UNORM
        NumPut("uint", 1, d2dBmpProps, 4)  ; D2D1_ALPHA_MODE_PREMULTIPLIED
        pD2DBitmap := this.ID2D1RenderTarget.CreateBitmapFromWicBitmap(pWICBitmapSource, d2dBmpProps)
        Direct2D.release(wicBitmap.pWICBitmap), wicBitmap.pWICBitmap := 0
        return this.d2dBitmaps[imgPath] := pD2dBitmap
    }

    GetSavedOrCreateTextFormat(fontName, fontSize) {
        fK := Format("{}_{}", fontName, fontSize)
        if this.textFormats.Has(fK)
            return this.textFormats[fK]

        return this.textFormats[fK] := Direct2D.IDWriteFactory.CreateTextFormat(fontName, fontSize)
    }

    GetSavedOrCreateSolidBrush(c) {
        bK := Format("{}", c)
        if this.solidBrushes.Has(bK)
            return this.solidBrushes[bK]

        if c <= 0xFFFFFF
            c := c | 0xFF000000
        static sColor := Buffer(16, 0)
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
    GetSavedOrCreateStrokeStyle(capStyle, shapeStyle := 0) {
        sK := Format("{}_{}", capStyle, shapeStyle)
        if this.strokeStyles.Has(sK)
            return this.strokeStyles[sK]

        static styleProps := Buffer(4 * 7, 0)
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
        if w != 0 and h != 0
            this.ResizeRenderTarget(this.width := w, this.height := h)
        this.winRect := [x, y, x + this.width, y + this.height]
        DllCall("MoveWindow", "Uptr", this.hwnd, "int", x, "int", y, "int", this.width, "int", this.height, "char", 1)
    }

    GetMousePosRefAttachWin(&x, &y) {
        DllCall("GetCursorPos", "int64*", &cursorPoint := 0)
        x := cursorPoint & 0xffffffff, y := cursorPoint >> 32
        if this.winRect.Length {
            insideWin := x > this.winRect[1] - 1 && x < this.winRect[3] + 1
                && y > this.winRect[2] - 1 && y < this.winRect[4] + 1
            x -= this.winRect[1], y -= this.winRect[2]
            return insideWin
        }
        return 0
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
