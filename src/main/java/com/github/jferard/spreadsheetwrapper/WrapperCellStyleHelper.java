package com.github.jferard.spreadsheetwrapper;

import java.util.HashMap;
import java.util.Map;

import jxl.format.Colour;

import org.apache.poi.ss.usermodel.Color;

public class WrapperCellStyleHelper {
	private final Map<String, WrapperColor> colorByHex;
	private final Map<Color, WrapperColor> colorByHssf;
	private final Map<Colour, WrapperColor> colorByJxl;

	public WrapperCellStyleHelper() {
		this.colorByHex = new HashMap<String, WrapperColor>();
		this.colorByHssf = new HashMap<Color, WrapperColor>();
		this.colorByJxl = new HashMap<Colour, WrapperColor>();

		for (final WrapperColor wrapperColor : WrapperColor.values()) {
			this.colorByHex.put(wrapperColor.toHex(), wrapperColor);
			this.colorByHssf.put(wrapperColor.getHssfColor(), wrapperColor);
			this.colorByJxl.put(wrapperColor.getJxlColor(), wrapperColor);
		}
	}

	public WrapperColor getColor(final Color hssfColor) {
		return this.colorByHssf.get(hssfColor);
	}

	public WrapperColor getColor(final Colour jxlColour) {
		return this.colorByJxl.get(jxlColour);
	}

	public WrapperColor getColor(final String hex) {
		return this.colorByHex.get(hex);
	}
}
