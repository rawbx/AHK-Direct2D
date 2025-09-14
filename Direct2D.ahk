class Direct2D {
	__New(bindHwnd) {
		this.hwnd := bindHwnd
		this.x := 0, this.y := 0
		this.width := 0, this.height := 0
		this.isDrawing := 0

		this.textFormats := Map()
		this.solidBrushes := Map()
		this.strokeStyles := Map()

		; init ID2D1HwndRenderTarget
		Direct2D.ID2D1RenderTarget.init(this.width, this.height, this.hwnd)

		; set window visible
		DllCall("ShowWindow", "Uptr", this.hwnd, "uint", SW_SHOWNOACTIVATE := 4)
		DllCall("SetLayeredWindowAttributes", "Uptr", this.hwnd, "Uint", ColorKey := 0, "char", Alpha := 255, "uint", LWA_ALPHA := 2)

		margins := Buffer(16, 0)
		NumPut("int", -1, margins, 0), NumPut("int", -1, margins, 4)
		NumPut("int", -1, margins, 8), NumPut("int", -1, margins, 12)
		DllCall("dwmapi\DwmExtendFrameIntoClientArea", "Uptr", this.hwnd, "ptr", margins, "uint")

		Direct2D.ID2D1RenderTarget.SetAntiAliasMode(D2D1_ANTIALIAS_MODE_PER_PRIMITIVE := 0)
		Direct2D.ID2D1RenderTarget.SetTextAntiAliasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE := 1)
		this.SetPosition(0, 0)
		this.Clear()
	}

	__Delete() {
		for _, s in this.solidBrushes
			Direct2D.release(s)
		for _, t in this.textFormats
			Direct2D.release(t)
		for _, st in this.strokeStyles
			Direct2D.release(st)
		Direct2D.release(Direct2D.ID2D1RenderTarget.Get())
		Direct2D.release(Direct2D.IDWriteFactory.Get())
		Direct2D.release(Direct2D.ID2D1Factory.Get())
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

		static __Delete() => Direct2D.release(this.pF)

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

	; IDWriteTextLayout::GetMetrics
	GetMetrics(text, fontName := "Segoe UI", fontSize := 16) {
		pTextFormat := this.GetSavedTextFormat(fontName, fontSize)
		pTextLayout := Direct2D.IDWriteFactory.CreateTextLayout(text, pTextFormat)
		; https://learn.microsoft.com/zh-cn/windows/win32/api/dwrite/ns-dwrite-dwrite_text_metrics
		; GetMetrics for the formatted text.
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
		if DllCall(Direct2D.vTable(pTextLayout, 60),
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
		static init(width, height, target) {
			rtProps := Buffer(64, 0)
			NumPut("uint", 1, rtProps, 8) ; AlphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED
			if IsInteger(target) { ; Hwnd
				NumPut("float", 96, rtProps, 12) ; dpiX
				NumPut("float", 96, rtProps, 16) ; dpiY
				hRtProps := Buffer(64, 0)
				NumPut("Uptr", target, hRtProps, 0)
				NumPut("uint", width, hRtProps, A_PtrSize) ; width
				NumPut("uint", height, hRtProps, A_PtrSize + 4) ; height
				NumPut("uint", 2, hRtProps, A_PtrSize + 8)
				this.pRT := Direct2D.ID2D1Factory.CreateHwndRenderTarget(rtProps, hRtProps)
			} else {
				NumPut("int", 87, rtProps, 4) ; DXGI_FORMAT_B8G8R8A8_UNORM
				if target == "wic" {
					CLSID_WICImagingFactory := "{cacaf262-9370-4615-a13b-9f5539da4c0a}"
					IID_IWICImagingFactory := "{ec5ec8a9-c395-4314-9c77-54d7a935ff70}"
					pWIC := ComObject(CLSID_WICImagingFactory, IID_IWICImagingFactory)
					WICPixelFormat32bppPBGRA := "{6fddc324-4e03-4bfe-b185-3d77768dc90a}"
					Direct2D.getCLSID(WICPixelFormat32bppPBGRA, &clsidWICPF)
					IWICImagingFactory_CreateBitmap := Direct2D.vTable(pWIC, 17)
					DllCall(IWICImagingFactory_CreateBitmap, "ptr", pWIC,
						"uint", width, "uint", height,
						"ptr", clsidWICPF,
						"uint", 1, ; WICBitmapCacheOnLoad
						"ptr*", &pWICBitmap := 0)
					this.pRT := Direct2D.ID2D1Factory.CreateWicBitmapRenderTarget(pWicBitmap, rtProps)
				} else {
					this.pRT := Direct2D.ID2D1Factory.CreateDCRenderTarget(rtProps)
				}
			}

			; ID2D1RenderTarget
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
			; ID2D1HwndRenderTarget
			this.VT_CheckWindowState := Direct2D.vTable(this.pRT, 57)
			this.VT_Resize := Direct2D.vTable(this.pRT, 58)
			this.VT_GetHwnd := Direct2D.vTable(this.pRT, 59)
			; ID2D1DCRenderTarget
			this.VT_BindDC := Direct2D.vTable(this.pRT, 57)
		}

		static Get() => this.pRT

		static CreateSolidBrush(sColor, brushProps := 0) =>
			(DllCall(this.VT_CreateSolidBrush, "Ptr", this.pRT, "Ptr", sColor, "Ptr", brushProps, "Ptr*", &pBrush := 0), pBrush)

		; D2D1_POINT_2F is different in 32 and 64 bits
		; reference https://github.com/Spawnova/ShinsOverlayClass/blob/main/AHK%20V2/ShinsOverlayClass.ahk#L626
		static DrawLine(pointStart, pointEnd, pBrush, strokeWidth, pStrokeStyle) {
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

		static DrawRectangle(rect, pBrush, strokeWidth, pStrokeStyle) =>
			DllCall(this.VT_DrawRectangle, "Ptr", this.pRT, "Ptr", rect, "ptr", pBrush, "float", strokeWidth, "ptr", pStrokeStyle)

		static FillRectangle(rect, pBrush) =>
			DllCall(this.VT_FillRectangle, "Ptr", this.pRT, "Ptr", rect, "ptr", pBrush)

		static DrawRoundedRectangle(roundedRect, pBrush, strokeWidth, pStrokeStyle) =>
			DllCall(this.VT_DrawRoundedRectangle, "Ptr", this.pRT, "Ptr", roundedRect, "ptr", pBrush, "float", strokeWidth, "ptr", pStrokeStyle)

		static FillRoundedRectangle(roundedRect, pBrush) =>
			DllCall(this.VT_FillRoundedRectangle, "Ptr", this.pRT, "Ptr", roundedRect, "ptr", pBrush)

		static DrawEllipse(sEllipse, pBrush, strokeWidth, pStrokeStyle) =>
			DllCall(this.VT_DrawEllipse, "Ptr", this.pRT, "Ptr", sEllipse, "ptr", pBrush, "float", strokeWidth, "ptr", pStrokeStyle)

		static FillEllipse(sEllipse, pBrush) =>
			DllCall(this.VT_FillEllipse, "Ptr", this.pRT, "Ptr", sEllipse, "ptr", pBrush)

		static DrawBitmap(bitmap, destRect := 0, opacity := 1, interpolationMode := 1, srcRect := 0) =>
			DllCall(this.VT_DrawBitmap, "ptr", this.pRT, "ptr", bitmap, "ptr", destRect, "float", opacity, "uint", interpolationMode, "ptr", srcRect)

		static DrawText(text, pTextFormat, rect, pBrush, drawOpt) =>
			DllCall(this.VT_DrawText, "ptr", this.pRT, "wstr", text, "uint", strlen(text), "ptr", pTextFormat, "ptr", rect, "ptr", pBrush, "uint", drawOpt, "uint", 0)

		static DrawTextLayout(point, pTextLayout, pBrush, drawOpt) {
			if Direct2D.isX64 {
				topLeftPt := Buffer(8)
				NumPut('float', point[1], topLeftPt, 0)
				NumPut('float', point[2], topLeftPt, 4)
				DllCall(this.VT_DrawTextLayout, "ptr", this.pRT, 'double', NumGet(topLeftPt, 0, 'double'), "ptr", pTextLayout, 'ptr', pBrush, "uint", drawOpt)
			} else {
				DllCall(this.VT_DrawTextLayout, "ptr", this.pRT, "float", point[1], "float", point[2], "ptr", pTextLayout, 'ptr', pBrush, "uint", drawOpt)
			}
		}

		static SetAntialiasMode(mode) => DllCall(this.VT_SetAntialiasMode, "Ptr", this.pRT, "Uint", mode)

		static SetTextAntialiasMode(mode) => DllCall(this.VT_SetTextAntialiasMode, "Ptr", this.pRT, "Uint", mode)

		static Clear() => DllCall(this.VT_Clear, "Ptr", this.pRT, "Ptr", 0)

		static BeginDraw() => DllCall(this.VT_BeginDraw, "Ptr", this.pRT)

		static EndDraw() => DllCall(this.VT_EndDraw, "Ptr", this.pRT, "Ptr*", 0, "Ptr*", 0)

		static Resize(size) => DllCall(this.VT_Resize, "Ptr", this.pRT, "ptr", size)

		static BindDC(rect) => DllCall(this.VT_BindDC, "Ptr", this.pRT, "ptr", rect)
	}

	Clear() {
		Direct2D.ID2D1RenderTarget.BeginDraw()
		Direct2D.ID2D1RenderTarget.Clear()
		Direct2D.ID2D1RenderTarget.EndDraw()
	}

	BeginDraw() {
		Direct2D.ID2D1RenderTarget.BeginDraw()
		Direct2D.ID2D1RenderTarget.Clear()
		return this.isDrawing := 1
	}

	EndDraw() {
		if (this.isDrawing)
			Direct2D.ID2D1RenderTarget.EndDraw()
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
		Direct2D.ID2D1RenderTarget.DrawText(text, pTextFormat, sRect, pBrush, drawOpt)
	}

	DrawTextLayout(text, x, y, color, pTextLayout, drawOpt := 4) {
		if !text
			return

		pBrush := this.GetSavedSolidBrush(color)
		Direct2D.ID2D1RenderTarget.DrawTextLayout([x, y], pTextLayout, pBrush, drawOpt)

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
		Direct2D.ID2D1RenderTarget.DrawLine(pointStart, pointEnd, pBrush, strokeWidth, pStrokeStyle)
	}

	DrawRectangle(x, y, w, h, color, strokeWidth := 2, strokeCapStyle := 0, strokeShapeStyle := 0) {
		pBrush := this.GetSavedSolidBrush(color)
		pStrokeStyle := this.GetSavedStrokeStyle(strokeCapStyle, strokeShapeStyle)
		sRect := Buffer(64)
		NumPut("float", x, sRect, 0)
		NumPut("float", y, sRect, 4)
		NumPut("float", x + w, sRect, 8)
		NumPut("float", y + h, sRect, 12)
		Direct2D.ID2D1RenderTarget.DrawRectangle(sRect, pBrush, strokeWidth, pStrokeStyle)
	}

	FillRectangle(x, y, w, h, color) {
		pBrush := this.GetSavedSolidBrush(color)
		sRect := Buffer(64)
		NumPut("float", x, sRect, 0)
		NumPut("float", y, sRect, 4)
		NumPut("float", x + w, sRect, 8)
		NumPut("float", y + h, sRect, 12)
		Direct2D.ID2D1RenderTarget.FillRectangle(sRect, pBrush)
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
		Direct2D.ID2D1RenderTarget.DrawRoundedRectangle(roundedRect, pBrush, strokeWidth, pStrokeStyle)
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
		Direct2D.ID2D1RenderTarget.FillRoundedRectangle(roundedRect, pBrush)
	}

	DrawEllipse(x, y, w, h, color, strokeWidth := 2, strokeCapStyle := 2, strokeShapeStyle := 0) {
		pBrush := this.GetSavedSolidBrush(color)
		pStrokeStyle := this.GetSavedStrokeStyle(strokeCapStyle, strokeShapeStyle)
		ellipse := Buffer(64)
		NumPut("float", x, ellipse, 0)
		NumPut("float", y, ellipse, 4)
		NumPut("float", w, ellipse, 8)
		NumPut("float", h, ellipse, 12)
		Direct2D.ID2D1RenderTarget.DrawEllipse(ellipse, pBrush, strokeWidth, pStrokeStyle)
	}

	FillEllipse(x, y, w, h, color) {
		pBrush := this.GetSavedSolidBrush(color)
		ellipse := Buffer(64)
		NumPut("float", x, ellipse, 0)
		NumPut("float", y, ellipse, 4)
		NumPut("float", w, ellipse, 8)
		NumPut("float", h, ellipse, 12)
		Direct2D.ID2D1RenderTarget.FillEllipse(ellipse, pBrush)
	}

	DrawCircle(x, y, radius, color, strokeWidth := 2, strokeCapStyle := 2, strokeShapeStyle := 0) {
		pBrush := this.GetSavedSolidBrush(color)
		pStrokeStyle := this.GetSavedStrokeStyle(strokeCapStyle, strokeShapeStyle)
		ellipse := Buffer(64)
		NumPut("float", x, ellipse, 0)
		NumPut("float", y, ellipse, 4)
		NumPut("float", radius, ellipse, 8)
		NumPut("float", radius, ellipse, 12)
		Direct2D.ID2D1RenderTarget.DrawEllipse(ellipse, pBrush, strokeWidth, pStrokeStyle)
	}

	FillCircle(x, y, radius, color) {
		pBrush := this.GetSavedSolidBrush(color)
		ellipse := Buffer(64)
		NumPut("float", x, ellipse, 0)
		NumPut("float", y, ellipse, 4)
		NumPut("float", radius, ellipse, 8)
		NumPut("float", radius, ellipse, 12)
		Direct2D.ID2D1RenderTarget.FillEllipse(ellipse, pBrush)
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
		pBrush := Direct2D.ID2D1RenderTarget.CreateSolidBrush(sColor)
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
		this.x := x, this.y := y
		if w != 0 and h != 0 {
			newSize := Buffer(16, 0)
			NumPut("uint", this.width := w, newSize, 0)
			NumPut("uint", this.height := h, newSize, 4)
			Direct2D.ID2D1RenderTarget.Resize(newSize)
		}
		DllCall("MoveWindow", "Uptr", this.hwnd, "int", x, "int", y, "int", this.width, "int", this.height, "char", 1)
	}
}
