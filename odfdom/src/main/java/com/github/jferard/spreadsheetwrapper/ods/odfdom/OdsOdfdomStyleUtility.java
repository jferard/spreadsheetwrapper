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
package com.github.jferard.spreadsheetwrapper.ods.odfdom;

import java.util.HashMap;
import java.util.Map;

import org.odftoolkit.odfdom.dom.style.props.OdfStyleProperty;
import org.odftoolkit.odfdom.dom.style.props.OdfTableCellProperties;
import org.odftoolkit.odfdom.dom.style.props.OdfTextProperties;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;

import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;
import com.github.jferard.spreadsheetwrapper.impl.StyleUtility;

/**
 * A little style utility for odftoolkit files
 *
 */
public class OdsOdfdomStyleUtility extends StyleUtility {
	private final Map<String, WrapperColor> colorByHex;

	/**
	 * @param helper
	 *            a helper for wrapper cell style
	 */
	public OdsOdfdomStyleUtility() {
		super();
		this.colorByHex = new HashMap<String, WrapperColor>();

		for (final WrapperColor wrapperColor : WrapperColor.values())
			this.colorByHex.put(wrapperColor.toHex(), wrapperColor);
	}

	/**
	 * @param style
	 *            *internal* cell style
	 * @return the new cell style format
	 */
	public WrapperCellStyle getCellStyle(final OdfStyle style) {
		final String fontWeight = style
				.getProperty(OdfTextProperties.FontWeight);
		final String backgroundColor = style
				.getProperty(OdfTableCellProperties.BackgroundColor);

		WrapperFont wrapperFont = null;
		if ("bold".equals(fontWeight))
			wrapperFont = new WrapperFont(true, false, -1, WrapperColor.BLACK);
		final WrapperColor wrapperColor;
		if (this.colorByHex.containsKey(backgroundColor))
			wrapperColor = this.colorByHex.get(backgroundColor);
		else
			wrapperColor = WrapperColor.WHITE;
		return new WrapperCellStyle(wrapperColor, wrapperFont);
	}

	/**
	 * @param styleString
	 *            the old style string
	 * @return the properties extracted from the style string
	 */
	public Map<OdfStyleProperty, String> getProperties(final String styleString) {
		final Map<OdfStyleProperty, String> properties = new HashMap<OdfStyleProperty, String>();
		final Map<String, String> props = this.getPropertiesMap(styleString);
		for (final Map.Entry<String, String> entry : props.entrySet()) {
			if (entry.getKey().equals(StyleUtility.FONT_WEIGHT)) {
				properties.put(OdfTextProperties.FontWeight, entry.getValue());
			} else if (entry.getKey().equals(StyleUtility.BACKGROUND_COLOR)) {
				properties.put(OdfTableCellProperties.BackgroundColor,
						entry.getValue());
			}
		}
		return properties;
	}

	/**
	 * @param wrapperCellStyle
	 *            the new style format
	 * @return the properties extracted from the style string
	 */
	public Map<OdfStyleProperty, String> getProperties(
			final WrapperCellStyle wrapperCellStyle) {
		final Map<OdfStyleProperty, String> properties = new HashMap<OdfStyleProperty, String>();
		final WrapperFont wrapperFont = wrapperCellStyle.getCellFont();
		if (wrapperFont != null && wrapperFont.isBold())
			properties.put(OdfTextProperties.FontWeight, "bold");
		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();
		if (backgroundColor != null) {
			properties.put(OdfTableCellProperties.BackgroundColor,
					backgroundColor.toHex());
		}
		return properties;
	}

	/**
	 * @param style
	 *            *internal* style
	 * @return the old style string that contains the properties fo the style
	 */
	public String getStyleString(final OdfStyle style) {
		final String fontWeight = style
				.getProperty(OdfTextProperties.FontWeight);
		final String backgroundColor = style
				.getProperty(OdfTableCellProperties.BackgroundColor);
		return String.format("font-weight:%s;background-color:%s", fontWeight,
				backgroundColor);
	}
}
