package com.github.jferard.spreadsheetwrapper;

import jxl.format.Colour;

import org.apache.poi.hssf.util.HSSFColor;

/**
 * The class WrapperColor wraps a color for ods (hex value) or xls (index). It
 * does not need anymore to have jxl or poi loaded to work.
 */
public enum WrapperColor {
	AQUA("AQUA", "#33cccc"), AUTOMATIC("AUTOMATIC", "#ffffff"), BLACK("BLACK",
			"#000000"), BLUE("BLUE", "#0000ff"), BLUE_GREY("BLUE_GREY",
			"#6666cc"), BLUE2("BLUE2", "#0000ff"), BRIGHT_GREEN("BRIGHT_GREEN",
			"#00ff00"), BROWN("BROWN", "#993300"), CORAL("CORAL", "#ff8080"), DARK_BLUE(
			"DARK_BLUE", "#000080"), DARK_BLUE2("DARK_BLUE2", "#000080"), DARK_GREEN(
			"DARK_GREEN", "#003300"), DARK_PURPLE("DARK_PURPLE", "#660066"), DARK_RED(
			"DARK_RED", "#800000"), DARK_RED2("DARK_RED2", "#800000"), DARK_TEAL(
			"DARK_TEAL", "#003366"), DARK_YELLOW("DARK_YELLOW", "#808000"), DEFAULT_BACKGROUND(
			"DEFAULT_BACKGROUND", "#ffffff"), DEFAULT_BACKGROUND1(
			"DEFAULT_BACKGROUND1", "#ffffff"), GOLD("GOLD", "#ffcc00"), GREEN(
			"GREEN", "#008000"), GREY_25_PERCENT("GREY_25_PERCENT", "#c0c0c0"), GREY_40_PERCENT(
			"GREY_40_PERCENT", "#969696"), GREY_50_PERCENT("GREY_50_PERCENT",
			"#808080"), GREY_80_PERCENT("GREY_80_PERCENT", "#333333"), ICE_BLUE(
			"ICE_BLUE", "#ccccff"), INDIGO("INDIGO", "#333399"), IVORY("IVORY",
			"#ffffcc"), LAVENDER("LAVENDER", "#cc99ff"), LIGHT_BLUE(
			"LIGHT_BLUE", "#3366ff"), LIGHT_GREEN("LIGHT_GREEN", "#ccffcc"), LIGHT_ORANGE(
			"LIGHT_ORANGE", "#ff9900"), LIGHT_TURQUOISE("LIGHT_TURQUOISE",
			"#ccffff"), LIGHT_TURQUOISE2("LIGHT_TURQUOISE2", "#ccffff"), LIME(
			"LIME", "#99cc00"), OCEAN_BLUE("OCEAN_BLUE", "#0066cc"), OLIVE_GREEN(
			"OLIVE_GREEN", "#333300"), ORANGE("ORANGE", "#ff6600"), PALE_BLUE(
			"PALE_BLUE", "#99ccff"), PALETTE_BLACK("PALETTE_BLACK", "#010000"), PERIWINKLE(
			"PERIWINKLE", "#9999ff"), PINK("PINK", "#ff00ff"), PINK2("PINK2",
			"#ff00ff"), PLUM("PLUM", "#993366"), PLUM2("PLUM2", "#993366"), RED(
			"RED", "#ff0000"), ROSE("ROSE", "#ff99cc"), SEA_GREEN("SEA_GREEN",
			"#339966"), SKY_BLUE("SKY_BLUE", "#00ccff"), TAN("TAN", "#ffcc99"), TEAL(
			"TEAL", "#008080"), TEAL2("TEAL2", "#008080"), TURQOISE2(
			"TURQOISE2", "#00ffff"), TURQUOISE("TURQUOISE", "#00ffff"), VERY_LIGHT_YELLOW(
			"VERY_LIGHT_YELLOW", "#ffff99"), VIOLET("VIOLET", "#808000"), VIOLET2(
			"VIOLET2", "#800080"), WHITE("WHITE", "#ffffff"), YELLOW("YELLOW",
			"#ffff00"), YELLOW2("YELLOW2", "#ffff00");

	/** hex value of the color (R G B) */
	private String hex;
	/** the color as defined in poi */
	private HSSFColor hssfColor;
	/** the color as defined in jxl */
	private Colour jxlColor;
	/** the xls (poi & jxl) name of the color, for dynamic load */
	private String simpleName;

	/**
	 * @param simpleName
	 *            the xls (poi & jxl) name of the color
	 * @param hex
	 *            hex value of the color (R G B)
	 */
	WrapperColor(final String simpleName, final String hex) {
		this.simpleName = simpleName;
		this.hex = hex;
		try {
			final Class<?> jxlClazz = Class.forName("jxl.format.Colour");
			this.jxlColor = (Colour) jxlClazz.getDeclaredField(simpleName).get(
					null);
		} catch (final ClassNotFoundException e) {
			this.jxlColor = null;
		} catch (final IllegalArgumentException e) {
			this.jxlColor = null;
		} catch (final IllegalAccessException e) {
			this.jxlColor = null;
		} catch (final NoSuchFieldException e) {
			this.jxlColor = null;
		} catch (final SecurityException e) {
			this.jxlColor = null;
		}
		try {
			final Class<?> hssfClazz = Class
					.forName("org.apache.poi.hssf.util.HSSFColor$" + simpleName);
			this.hssfColor = (HSSFColor) hssfClazz.newInstance();
		} catch (final ClassNotFoundException e) {
			this.hssfColor = null;
		} catch (final IllegalArgumentException e) {
			this.hssfColor = null;
		} catch (final IllegalAccessException e) {
			this.hssfColor = null;
		} catch (final SecurityException e) {
			this.hssfColor = null;
		} catch (final InstantiationException e) {
			this.hssfColor = null;
		}
	}

	public HSSFColor getHssfColor() {
		return this.hssfColor;
	}

	public Colour getJxlColor() {
		return this.jxlColor;
	}

	public String getSimpleName() {
		return this.simpleName;
	}

	public String toHex() {
		return this.hex;
	}
}
