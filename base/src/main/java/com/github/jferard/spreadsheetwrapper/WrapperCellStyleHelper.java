package com.github.jferard.spreadsheetwrapper;

import java.util.HashMap;
import java.util.Map;

//import jxl.format.Colour;
//
//import org.apache.poi.ss.usermodel.Color;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class WrapperCellStyleHelper {
	private final Map<String, WrapperColor> colorByHex;
//	private final Map<Color, WrapperColor> colorByHssf;
//	private final Map<Colour, WrapperColor> colorByJxl;

	public WrapperCellStyleHelper() {
		this.colorByHex = new HashMap<String, WrapperColor>();
//		this.colorByHssf = new HashMap<Color, WrapperColor>();
//		this.colorByJxl = new HashMap<Colour, WrapperColor>();

		for (final WrapperColor wrapperColor : WrapperColor.values()) {
			this.colorByHex.put(wrapperColor.toHex(), wrapperColor);
//			this.colorByHssf.put(wrapperColor.getHssfColor(), wrapperColor);
//			this.colorByJxl.put(wrapperColor.getJxlColor(), wrapperColor);
		}
	}

//	public WrapperColor getColor(final/*@Nullable*/Color hssfColor) {
//		WrapperColor wrapperColor;
//		if (this.colorByHssf.containsKey(hssfColor))
//			wrapperColor = this.colorByHssf.get(hssfColor);
//		else
//			wrapperColor = WrapperColor.AUTOMATIC;
//		return wrapperColor;
//	}
//
//	public WrapperColor getColor(final/*@Nullable*/Colour jxlColour) {
//		WrapperColor wrapperColor;
//		if (this.colorByJxl.containsKey(jxlColour))
//			wrapperColor = this.colorByJxl.get(jxlColour);
//		else
//			wrapperColor = WrapperColor.AUTOMATIC;
//		return wrapperColor;
//	}

	public WrapperColor getColor(final String hex) {
		WrapperColor wrapperColor;
		if (this.colorByHex.containsKey(hex))
			wrapperColor = this.colorByHex.get(hex);
		else
			wrapperColor = WrapperColor.AUTOMATIC;
		return wrapperColor;
	}
}
