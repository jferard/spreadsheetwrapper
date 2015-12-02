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

import org.odftoolkit.odfdom.dom.element.table.TableTableCellElementBase;
import org.odftoolkit.odfdom.dom.style.props.OdfStyleProperty;
import org.odftoolkit.odfdom.dom.style.props.OdfTableCellProperties;
import org.odftoolkit.odfdom.dom.style.props.OdfTextProperties;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;

import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;

/**
 * A little style utility for odftoolkit files
 *
 */
public class OdsOdfdomStyleHelper {
	/** hex (#000000) -> color */
	private final Map<String, WrapperColor> colorByHex;

	/**
	 * @param helper
	 *            a helper for wrapper cell style
	 */
	public OdsOdfdomStyleHelper() {
		super();
		this.colorByHex = new HashMap<String, WrapperColor>();

		for (final WrapperColor wrapperColor : WrapperColor.values())
			this.colorByHex.put(wrapperColor.toHex(), wrapperColor);
	}

	/** @return the cell style from the element */
	public WrapperCellStyle getCellStyle(
			final TableTableCellElementBase odfElement) {
		final String backgroundColor = odfElement
				.getProperty(OdfTableCellProperties.BackgroundColor);
		final String fontWeight = odfElement
				.getProperty(OdfTextProperties.FontWeight);
		final String fontSize = odfElement
				.getProperty(OdfTextProperties.FontSize);
		final String fontStyle = odfElement
				.getProperty(OdfTextProperties.FontStyle);

		final WrapperFont wrapperFont = new WrapperFont();
		if (fontWeight == null)
			wrapperFont.setBold(WrapperCellStyle.DEFAULT);
		else if ("bold".equals(fontWeight))
			wrapperFont.setBold();
		else if ("normal".equals(fontWeight))
			wrapperFont.setBold(WrapperCellStyle.NO);

		if (fontStyle == null)
			wrapperFont.setItalic(WrapperCellStyle.DEFAULT);
		else if ("italic".equals(fontStyle))
			wrapperFont.setItalic();
		else if ("normal".equals(fontStyle))
			wrapperFont.setItalic(WrapperCellStyle.NO);

		if (fontSize == null)
			wrapperFont.setSize(WrapperCellStyle.DEFAULT);
		else
			wrapperFont.setSize(OdsOdfdomStyleHelper.sizeToPoints(fontSize)); // 10pt, 10cm, units
		
		return new WrapperCellStyle(
				WrapperColor.stringToColor(backgroundColor), wrapperFont);
	}

	private static int sizeToPoints(String fontSize) {
		final double ret;
		final int length = fontSize.length();
		if (length > 2) {
			String value = fontSize.substring(0, length-2);
			String unit = fontSize.substring(length-2);
			double tempValue = Double.valueOf(value);
			if ("in".equals(unit))
				ret = tempValue*72.0;
			else if ("cm".equals(unit))
				ret = tempValue/2.54 * 72.0;
			else if ("mm".equals(unit))
				ret = tempValue/2.54 * 72.0;
			else if ("px".equals(unit))
				ret = tempValue;
			else if ("pc".equals(unit))
				ret = tempValue/6.0 *72.0;
			else
				ret = tempValue;
		} else
			ret = Double.valueOf(fontSize);
		return (int) ret;
	}

	/**
	 * @param wrapperCellStyle
	 *            the new style format
	 * @return the properties extracted from the style string
	 */
	public Map<OdfStyleProperty, String> getProperties(
			final WrapperCellStyle wrapperCellStyle) {
		final Map<OdfStyleProperty, String> properties = new HashMap<OdfStyleProperty, String>(); // NOPMD
		// by
		// Julien
		// on
		// 24/11/15
		// 19:24
		final WrapperFont wrapperFont = wrapperCellStyle.getCellFont();
		if (wrapperFont != null
				&& wrapperFont.getBold() == WrapperCellStyle.YES)
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
	@Deprecated
	public String getStyleString(final OdfStyle style) {
		final String fontWeight = style
				.getProperty(OdfTextProperties.FontWeight);
		final String backgroundColor = style
				.getProperty(OdfTableCellProperties.BackgroundColor);
		return String.format("font-weight:%s;background-color:%s", fontWeight,
				backgroundColor);
	}

	/**
	 * @param style
	 *            *internal* cell style
	 * @return the new cell style format
	 */
	public WrapperCellStyle toCellStyle(final OdfStyle style) {
		final String fontWeight = style
				.getProperty(OdfTextProperties.FontWeight);
		final String fontStyle = style.getProperty(OdfTextProperties.FontStyle);
		final String fontSize = style.getProperty(OdfTextProperties.FontSize);
		final String fontColor = style.getProperty(OdfTextProperties.Color);
		final String backgroundColor = style
				.getProperty(OdfTableCellProperties.BackgroundColor);

		final WrapperFont wrapperFont = new WrapperFont();
		if ("bold".equals(fontWeight))
			wrapperFont.setBold();
		if (fontSize != null)
			wrapperFont.setSize(Integer.valueOf(fontSize));
		if (fontColor != null && this.colorByHex.containsKey(fontColor))
			wrapperFont.setColor(this.colorByHex.get(fontColor));
		final WrapperColor wrapperColor;
		if (this.colorByHex.containsKey(backgroundColor))
			wrapperColor = this.colorByHex.get(backgroundColor);
		else
			wrapperColor = null;
		return new WrapperCellStyle(wrapperColor, wrapperFont);
	}
}
