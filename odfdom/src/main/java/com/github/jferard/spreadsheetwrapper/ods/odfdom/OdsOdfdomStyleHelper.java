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

import com.github.jferard.spreadsheetwrapper.Util;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;

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
		final String fontColor = odfElement
				.getProperty(OdfTextProperties.Color);

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

		if (fontSize != null)
			wrapperFont.setSize(OdsConstants.sizeToPoints(fontSize)); // 10pt, 10cm, units
		
		if (fontColor != null)
			wrapperFont.setColor(WrapperColor.stringToColor(fontColor));
		
		return new WrapperCellStyle(
				WrapperColor.stringToColor(backgroundColor), WrapperCellStyle.DEFAULT, wrapperFont);
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
		if (wrapperFont != null) {
				final int bold = wrapperFont.getBold();
				final double size = wrapperFont.getSize();
				final int italic = wrapperFont.getItalic();
				final WrapperColor fontColor = wrapperFont.getColor();
				if (bold == WrapperCellStyle.YES) {
					properties.put(OdfTextProperties.FontWeight, "bold");
					properties.put(OdfTextProperties.FontWeightAsian, "bold");
					properties.put(OdfTextProperties.FontWeightComplex, "bold");
				} else if (bold == WrapperCellStyle.NO) {
					properties.put(OdfTextProperties.FontWeight, "normal");
					properties.put(OdfTextProperties.FontWeightAsian, "normal");
					properties.put(OdfTextProperties.FontWeightComplex, "normal");
				}
				
				if (italic == WrapperCellStyle.YES) {
					properties.put(OdfTextProperties.FontStyle, "italic");
					properties.put(OdfTextProperties.FontStyleAsian, "italic");
					properties.put(OdfTextProperties.FontStyleComplex, "italic");
				} else if (italic == WrapperCellStyle.NO) {
					properties.put(OdfTextProperties.FontStyle, "normal");
					properties.put(OdfTextProperties.FontStyleAsian, "normal");
					properties.put(OdfTextProperties.FontStyleComplex, "normal");
				}
				
				if (!Util.almostEqual(size, WrapperCellStyle.DEFAULT)) {
					properties.put(OdfTextProperties.FontSize, Double.toString(size)+"pt");
				}
				
				if (fontColor != null) {
					properties.put(OdfTextProperties.Color, fontColor.toHex());
				}
		}
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
		return new WrapperCellStyle(wrapperColor, WrapperCellStyle.DEFAULT, wrapperFont);
	}
}
