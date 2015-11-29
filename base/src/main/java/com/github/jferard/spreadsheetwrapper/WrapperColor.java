/*******************************************************************************
 *     SpreadsheetWrapper - An abstraction layer over some APIs for Excel or Calc
 *     Copyright (C) 2015  J. Férard
 *
 *     This program is free software: you can redistribute it and/or modify
 *     it under the terms of the GNU General Public License as published by
 *     the Free Software Foundation, either version 3 of the License, or
 *     (at your option) any later version.
 *
 *     This program is distributed in the hope that it will be useful,
 *     but WITHOUT ANY WARRANTY; without even the implied warranty of
 *     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *     GNU General Public License for more details.
 *
 *     You should have received a copy of the GNU General Public License
 *     along with this program.  If not, see <http://www.gnu.org/licenses/>.
 *******************************************************************************/
package com.github.jferard.spreadsheetwrapper;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

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

	/** the start char for hex strings */
	private static final char HASH_FOR_HEX = '#';

	/**
	 * @param colorString
	 *            #000000
	 * @return the color
	 */
	public static/*@Nullable*/WrapperColor stringToColor(
			/*@Nullable*/final String colorString) {
		if (colorString == null)
			return null;

		WrapperColor wrapperColor = null;
		if (colorString.charAt(0) == WrapperColor.HASH_FOR_HEX) {
			int minDistance = 1024;
			for (final WrapperColor otherWrapperColor : WrapperColor.values()) {
				final String otherColorAsHex = otherWrapperColor.toHex();
				final int distance = WrapperColor.distance(colorString,
						otherColorAsHex);
				if (distance < minDistance) {
					wrapperColor = otherWrapperColor;
					minDistance = distance;
				}
			}
		} else {
			try {
				wrapperColor = WrapperColor.valueOf(colorString);
			} catch (final IllegalArgumentException e) { // NOPMD by Julien on
				// 21/11/15 11:28
				// do nothing : already null
			}
		}
		return wrapperColor;
	}

	private static int distance(final String colorAsHex,
			final String otherColorAsHex) {
		int distance = 0;
		for (int i = 1; i < 7; i += 2) {
			distance += Math.abs(Integer.valueOf(
					colorAsHex.substring(i, i + 2), 16)
					- Integer.valueOf(otherColorAsHex.substring(i, i + 2), 16));
		}
		return distance;
	}

	/** hex value of the color (R G B) */
	private String hex;

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
	}

	/** @return the xls (poi & jxl) name of the color, for dynamic load */
	public String getSimpleName() {
		return this.simpleName;
	}

	/** @return hex value of the color (R G B) */
	public String toHex() {
		return this.hex;
	}
}
