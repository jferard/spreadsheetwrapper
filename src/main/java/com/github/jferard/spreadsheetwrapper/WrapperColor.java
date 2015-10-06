package com.github.jferard.spreadsheetwrapper;

import jxl.format.Colour;

import org.apache.poi.hssf.util.HSSFColor;

public enum WrapperColor {
	BLACK("BLACK", "#000000"), WHITE("WHITE", "#ffffff"), DEFAULT_BACKGROUND1(
			"DEFAULT_BACKGROUND1", "#ffffff"), DEFAULT_BACKGROUND(
			"DEFAULT_BACKGROUND", "#ffffff"), PALETTE_BLACK("PALETTE_BLACK",
			"#010000"), RED("RED", "#ff0000"), BRIGHT_GREEN("BRIGHT_GREEN",
			"#00ff00"), BLUE("BLUE", "#0000ff"), YELLOW("YELLOW", "#ffff00"), PINK(
			"PINK", "#ff00ff"), TURQUOISE("TURQUOISE", "#00ffff"), DARK_RED(
			"DARK_RED", "#800000"), GREEN("GREEN", "#008000"), DARK_BLUE(
			"DARK_BLUE", "#000080"), DARK_YELLOW("DARK_YELLOW", "#808000"), VIOLET(
			"VIOLET", "#808000"), TEAL("TEAL", "#008080"), GREY_25_PERCENT(
			"GREY_25_PERCENT", "#c0c0c0"), GREY_50_PERCENT("GREY_50_PERCENT",
			"#808080"), PERIWINKLE("PERIWINKLE", "#9999ff"), PLUM2("PLUM2",
			"#993366"), IVORY("IVORY", "#ffffcc"), LIGHT_TURQUOISE2(
			"LIGHT_TURQUOISE2", "#ccffff"), DARK_PURPLE("DARK_PURPLE",
			"#660066"), CORAL("CORAL", "#ff8080"), OCEAN_BLUE("OCEAN_BLUE",
			"#0066cc"), ICE_BLUE("ICE_BLUE", "#ccccff"), DARK_BLUE2(
			"DARK_BLUE2", "#000080"), PINK2("PINK2", "#ff00ff"), YELLOW2(
			"YELLOW2", "#ffff00"), TURQOISE2("TURQOISE2", "#00ffff"), VIOLET2(
			"VIOLET2", "#800080"), DARK_RED2("DARK_RED2", "#800000"), TEAL2(
			"TEAL2", "#008080"), BLUE2("BLUE2", "#0000ff"), SKY_BLUE(
			"SKY_BLUE", "#00ccff"), LIGHT_TURQUOISE("LIGHT_TURQUOISE",
			"#ccffff"), LIGHT_GREEN("LIGHT_GREEN", "#ccffcc"), VERY_LIGHT_YELLOW(
			"VERY_LIGHT_YELLOW", "#ffff99"), PALE_BLUE("PALE_BLUE", "#99ccff"), ROSE(
			"ROSE", "#ff99cc"), LAVENDER("LAVENDER", "#cc99ff"), TAN("TAN",
			"#ffcc99"), LIGHT_BLUE("LIGHT_BLUE", "#3366ff"), AQUA("AQUA",
			"#33cccc"), LIME("LIME", "#99cc00"), GOLD("GOLD", "#ffcc00"), LIGHT_ORANGE(
			"LIGHT_ORANGE", "#ff9900"), ORANGE("ORANGE", "#ff6600"), BLUE_GREY(
			"BLUE_GREY", "#6666cc"), GREY_40_PERCENT("GREY_40_PERCENT",
			"#969696"), DARK_TEAL("DARK_TEAL", "#003366"), SEA_GREEN(
			"SEA_GREEN", "#339966"), DARK_GREEN("DARK_GREEN", "#003300"), OLIVE_GREEN(
			"OLIVE_GREEN", "#333300"), BROWN("BROWN", "#993300"), PLUM("PLUM",
			"#993366"), INDIGO("INDIGO", "#333399"), GREY_80_PERCENT(
			"GREY_80_PERCENT", "#333333"), AUTOMATIC("AUTOMATIC", "#ffffff");

	private String simpleName;
	private String hex;
	private HSSFColor hssfColor;
	private Colour jxlColor;

	WrapperColor(final String simpleName, final String hex) {
		this.simpleName = simpleName;
		this.hex = hex;
		try {
			Class<?> jxlClazz = Class.forName("jxl.format.Colour");
			this.jxlColor = (Colour) jxlClazz.getDeclaredField(simpleName).get(
					null);
		} catch (ClassNotFoundException e) {
			this.jxlColor = null;
		} catch (IllegalArgumentException e) {
			this.jxlColor = null;
		} catch (IllegalAccessException e) {
			this.jxlColor = null;
		} catch (NoSuchFieldException e) {
			this.jxlColor = null;
		} catch (SecurityException e) {
			this.jxlColor = null;
		}
		try {
			Class<?> hssfClazz = Class
					.forName("org.apache.poi.hssf.util.HSSFColor$" + simpleName);
			this.hssfColor = (HSSFColor) hssfClazz.newInstance();
		} catch (ClassNotFoundException e) {
			this.hssfColor = null;
		} catch (IllegalArgumentException e) {
			this.hssfColor = null;
		} catch (IllegalAccessException e) {
			this.hssfColor = null;
		} catch (SecurityException e) {
			this.hssfColor = null;
		} catch (InstantiationException e) {
			this.hssfColor = null;
		}
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

	public String getSimpleName() {
		return this.simpleName;
	}
}
