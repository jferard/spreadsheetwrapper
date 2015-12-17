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
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;

/**
 * A little style utility for odftoolkit files
 *
 */
public class OdsOdfdomStyleHelper {
	private static void setProperties(
			final Map<OdfStyleProperty, String> properties,
			final String attributeValue,
			final OdfStyleProperty... propertyArray) {
		for (final OdfStyleProperty property : propertyArray)
			properties.put(property, attributeValue);
	}

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
		final String fontWeight = odfElement
				.getProperty(OdfTextProperties.FontWeight);
		final String fontSize = odfElement
				.getProperty(OdfTextProperties.FontSize);
		final String fontStyle = odfElement
				.getProperty(OdfTextProperties.FontStyle);
		final String fontFamily = odfElement
				.getProperty(OdfTextProperties.FontFamily);
		final String fontColor = odfElement
				.getProperty(OdfTextProperties.Color);
		final String backgroundColor = odfElement
				.getProperty(OdfTableCellProperties.BackgroundColor);
		final String border = odfElement
				.getProperty(OdfTableCellProperties.Border);

		return this.getCellStyle(fontWeight, fontStyle, fontSize, fontColor,
				fontFamily, backgroundColor, border);
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
			final String fontFamily = wrapperFont.getFamily();
			if (bold == WrapperCellStyle.YES) {
				OdsOdfdomStyleHelper.setProperties(properties,
						OdsConstants.BOLD_ATTR_VALUE,
						OdfTextProperties.FontWeight,
						OdfTextProperties.FontWeightAsian,
						OdfTextProperties.FontWeightComplex);
			} else if (bold == WrapperCellStyle.NO) {
				OdsOdfdomStyleHelper.setProperties(properties,
						OdsConstants.NORMAL_ATTR_VALUE,
						OdfTextProperties.FontWeight,
						OdfTextProperties.FontWeightAsian,
						OdfTextProperties.FontWeightComplex);
			}

			if (italic == WrapperCellStyle.YES) {
				OdsOdfdomStyleHelper.setProperties(properties,
						OdsConstants.ITALIC_ATTR_VALUE,
						OdfTextProperties.FontStyle,
						OdfTextProperties.FontStyleAsian,
						OdfTextProperties.FontStyleComplex);
			} else if (italic == WrapperCellStyle.NO) {
				OdsOdfdomStyleHelper.setProperties(properties,
						OdsConstants.NORMAL_ATTR_VALUE,
						OdfTextProperties.FontStyle,
						OdfTextProperties.FontStyleAsian,
						OdfTextProperties.FontStyleComplex);
			}

			if (!Util.almostEqual(size, WrapperCellStyle.DEFAULT)) {
				properties.put(OdfTextProperties.FontSize,
						Double.toString(size) + "pt");
			}

			if (fontColor != null) {
				properties.put(OdfTextProperties.Color, fontColor.toHex());
			}

			if (fontFamily != null) {
				OdsOdfdomStyleHelper.setProperties(properties, fontFamily,
						OdfTextProperties.FontFamily,
						OdfTextProperties.FontName,
						OdfTextProperties.FontFamilyAsian,
						OdfTextProperties.FontNameAsian,
						OdfTextProperties.FontFamilyComplex,
						OdfTextProperties.FontNameComplex);
			}
		}
		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();
		final double borderLineWidth = wrapperCellStyle.getBorderLineWidth();
		if (backgroundColor != null) {
			properties.put(OdfTableCellProperties.BackgroundColor,
					backgroundColor.toHex());
		}
		if (borderLineWidth != WrapperCellStyle.DEFAULT) {
			properties.put(OdfTableCellProperties.Border,
					Double.toString(borderLineWidth) + "pt solid #000000");
		}
		return properties;
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
		final String fontFamily = style
				.getProperty(OdfTextProperties.FontFamily);
		final String backgroundColor = style
				.getProperty(OdfTableCellProperties.BackgroundColor);
		final String border = style.getProperty(OdfTableCellProperties.Border);

		return this.getCellStyle(fontWeight, fontStyle, fontSize, fontColor,
				fontFamily, backgroundColor, border);
	}

	private WrapperCellStyle getCellStyle(final String fontWeight,
			final String fontStyle, final String fontSize,
			final String fontColor, final String fontFamily,
			final String backgroundColor, final String border) {
		final WrapperFont wrapperFont = new WrapperFont();
		if (fontWeight != null) {
			if (fontWeight.equals("bold"))
				wrapperFont.setBold();
			else if (fontWeight.equals(OdsConstants.NORMAL_ATTR_VALUE))
				wrapperFont.setBold(WrapperCellStyle.NO);
		}
		if (fontSize != null)
			wrapperFont.setSize(OdsConstants.sizeToPoints(fontSize));
		if (fontColor != null && this.colorByHex.containsKey(fontColor))
			wrapperFont.setColor(this.colorByHex.get(fontColor));
		if (fontFamily != null)
			wrapperFont.setFamily(fontFamily);
		if (fontStyle != null) {
			if (fontStyle.equals(OdsConstants.ITALIC_ATTR_VALUE))
				wrapperFont.setItalic();
			else if (fontStyle.equals(OdsConstants.NORMAL_ATTR_VALUE))
				wrapperFont.setItalic(WrapperCellStyle.NO);
		}

		final WrapperCellStyle wrapperStyle = new WrapperCellStyle()
		.setCellFont(wrapperFont);
		if (this.colorByHex.containsKey(backgroundColor))
			wrapperStyle.setBackgroundColor(this.colorByHex
					.get(backgroundColor));
		if (border != null) {
			final String[] split = border.split("\\s+");
			wrapperStyle
			.setBorderLineWidth(OdsConstants.sizeToPoints(split[0]));
		}

		return wrapperStyle;
	}
}
