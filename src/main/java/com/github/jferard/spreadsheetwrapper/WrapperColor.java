package com.github.jferard.spreadsheetwrapper;

import jxl.format.Colour;

import org.apache.poi.hssf.util.HSSFColor;

public enum WrapperColor {
	AQUA(Colour.AQUA, new HSSFColor.AQUA()), AUTOMATIC(Colour.AUTOMATIC,
			new HSSFColor.AUTOMATIC()), BLACK(Colour.BLACK,
			new HSSFColor.BLACK()), BLUE(Colour.BLUE, new HSSFColor.BLUE());

	private String hex;
	private HSSFColor hssfColor;
	private Colour jxlColor;

	WrapperColor(final Colour jxlColor, final HSSFColor hssfColor) {
		this.jxlColor = jxlColor;
		this.hssfColor = hssfColor;
		final short[] triplet = this.hssfColor.getTriplet();
		final StringBuilder sb = new StringBuilder("#");
		for (final short l : triplet)
			sb.append(Integer.toHexString(l));
		this.hex = sb.toString();
	}

	public HSSFColor getHssfColor() {
		return this.hssfColor;
	}

	public Colour getJxlColor() {
		return this.jxlColor;
	}

	public String toHex() {
		return this.hex;
	}

}
