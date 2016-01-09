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

import org.odftoolkit.odfdom.dom.element.style.StyleStyleElement;
import org.odftoolkit.odfdom.dom.element.table.TableTableCellElementBase;
import org.odftoolkit.odfdom.dom.style.OdfStylePropertySet;
import org.odftoolkit.odfdom.dom.style.props.OdfStyleProperty;
import org.odftoolkit.odfdom.dom.style.props.OdfTableCellProperties;
import org.odftoolkit.odfdom.dom.style.props.OdfTextProperties;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;

import com.github.jferard.spreadsheetwrapper.StyleHelper;
import com.github.jferard.spreadsheetwrapper.Util;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;
import com.github.jferard.spreadsheetwrapper.style.Borders;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;

/**
 * A little style utility for odftoolkit files
 *
 */
public class OdsOdfdomStyleHelper implements
		StyleHelper<OdfStylePropertySet, TableTableCellElementBase> {
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
	@Override
	public WrapperCellStyle getWrapperCellStyle(
			final TableTableCellElementBase odfElement) {
		StyleStyleElement style = odfElement.getAutomaticStyle();

		if (style == null) {
			style = odfElement.reuseDocumentStyle(odfElement.getStyleName());
		}

		if (style != null) {
			return this.toWrapperCellStyle(style);
		} else
			return WrapperCellStyle.EMPTY;
	}

	/**
	 * @param wrapperCellStyle
	 *            the new style format
	 * @return the properties extracted from the style string
	 */
	private static Map<OdfStyleProperty, String> toProperties(
			final WrapperCellStyle wrapperCellStyle) {
		final Map<OdfStyleProperty, String> properties = new HashMap<OdfStyleProperty, String>();
		final WrapperFont wrapperFont = wrapperCellStyle.getCellFont();
		if (wrapperFont != null)
			properties.putAll(OdsOdfdomStyleHelper
					.getFontProperties(wrapperFont));

		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();
		if (backgroundColor != null) {
			properties.put(OdfTableCellProperties.BackgroundColor,
					backgroundColor.toHex());
		}
		final Borders borders = wrapperCellStyle.getBorders();
		if (borders != null)
			properties.putAll(OdsOdfdomStyleHelper
					.getBordersProperties(borders));

		return properties;
	}

	private static Map<OdfStyleProperty, String> getBordersProperties(
			final Borders borders) {
		final Map<OdfStyleProperty, String> bordersProperties = new HashMap<OdfStyleProperty, String>(); // NOPMD
		StringBuilder builder = new StringBuilder();
		final double borderLineWidth = borders.getLineWidth();
		if (borderLineWidth != WrapperCellStyle.DEFAULT) {
			builder.append(borderLineWidth).append("pt");
			builder.append(" solid ");
			final WrapperColor borderLineColor = borders.getLineColor();
			if (borderLineColor == null)
				builder.append("#000000");
			else
				builder.append(borderLineColor.toHex());
			bordersProperties.put(OdfTableCellProperties.Border,
					builder.toString());
		}
		return bordersProperties;
	}

	private final static Map<OdfStyleProperty, String> getFontProperties(
			final WrapperFont wrapperFont) {
		final Map<OdfStyleProperty, String> fontProperties = new HashMap<OdfStyleProperty, String>(); // NOPMD
		final int bold = wrapperFont.getBold();
		final double size = wrapperFont.getSize();
		final int italic = wrapperFont.getItalic();
		final WrapperColor fontColor = wrapperFont.getColor();
		final String fontFamily = wrapperFont.getFamily();
		if (bold == WrapperCellStyle.YES) {
			OdsOdfdomStyleHelper.setProperties(fontProperties,
					OdsConstants.BOLD_ATTR_VALUE, OdfTextProperties.FontWeight,
					OdfTextProperties.FontWeightAsian,
					OdfTextProperties.FontWeightComplex);
		} else if (bold == WrapperCellStyle.NO) {
			OdsOdfdomStyleHelper.setProperties(fontProperties,
					OdsConstants.NORMAL_ATTR_VALUE,
					OdfTextProperties.FontWeight,
					OdfTextProperties.FontWeightAsian,
					OdfTextProperties.FontWeightComplex);
		}

		if (italic == WrapperCellStyle.YES) {
			OdsOdfdomStyleHelper.setProperties(fontProperties,
					OdsConstants.ITALIC_ATTR_VALUE,
					OdfTextProperties.FontStyle,
					OdfTextProperties.FontStyleAsian,
					OdfTextProperties.FontStyleComplex);
		} else if (italic == WrapperCellStyle.NO) {
			OdsOdfdomStyleHelper.setProperties(fontProperties,
					OdsConstants.NORMAL_ATTR_VALUE,
					OdfTextProperties.FontStyle,
					OdfTextProperties.FontStyleAsian,
					OdfTextProperties.FontStyleComplex);
		}

		if (!Util.almostEqual(size, WrapperCellStyle.DEFAULT)) {
			fontProperties.put(OdfTextProperties.FontSize,
					Double.toString(size) + "pt");
		}

		if (fontColor != null) {
			fontProperties.put(OdfTextProperties.Color, fontColor.toHex());
		}

		if (fontFamily != null) {
			OdsOdfdomStyleHelper.setProperties(fontProperties, fontFamily,
					OdfTextProperties.FontFamily, OdfTextProperties.FontName,
					OdfTextProperties.FontFamilyAsian,
					OdfTextProperties.FontNameAsian,
					OdfTextProperties.FontFamilyComplex,
					OdfTextProperties.FontNameComplex);
		}
		return fontProperties;
	}

	/**
	 * @param style
	 *            *internal* cell style
	 * @return the new cell style format
	 */
	@Override
	public WrapperCellStyle toWrapperCellStyle(final OdfStylePropertySet style) {
		return this.toWrapperCellStyle(style.getProperties(style
				.getStrictProperties()));
	}

	private WrapperCellStyle toWrapperCellStyle(
			Map<OdfStyleProperty, String> propertyByAttrName) {
		WrapperFont wrapperFont = this.toWrapperCellFont(propertyByAttrName);
		final WrapperCellStyle wrapperStyle = new WrapperCellStyle()
				.setCellFont(wrapperFont);

		final String backgroundColor = propertyByAttrName
				.get(OdfTableCellProperties.BackgroundColor);
		if (this.colorByHex.containsKey(backgroundColor))
			wrapperStyle.setBackgroundColor(this.colorByHex
					.get(backgroundColor));

		Borders borders = OdsOdfdomStyleHelper.toCellBorders(propertyByAttrName);
		wrapperStyle.setBorders(borders);

		return wrapperStyle;
	}

	private static Borders toCellBorders(
			final Map<OdfStyleProperty, String> propertyByAttrName) {
		final Borders borders = new Borders();
		final String borderAttrValue = propertyByAttrName
				.get(OdfTableCellProperties.Border);
		if (borderAttrValue != null) {
			final String[] split = borderAttrValue.split("\\s+");
			borders.setLineWidth(OdsConstants.sizeToPoints(split[0]));
			// borders.setLineType(split[1]);
			if (!WrapperColor.BLACK.toHex().equals(split[2]))
				borders.setLineColor(WrapperColor.stringToColor(split[2]));
		}
		return borders;
	}

	private WrapperFont toWrapperCellFont(
			final Map<OdfStyleProperty, String> propertyByAttrName) {
		final WrapperFont wrapperFont = new WrapperFont();
		final String fontWeight = propertyByAttrName
				.get(OdfTextProperties.FontWeight);
		final String fontStyle = propertyByAttrName
				.get(OdfTextProperties.FontStyle);
		final String fontSize = propertyByAttrName
				.get(OdfTextProperties.FontSize);
		final String fontColor = propertyByAttrName
				.get(OdfTextProperties.Color);
		final String fontFamily = propertyByAttrName
				.get(OdfTextProperties.FontFamily);
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
		return wrapperFont;
	}

	@Override
	public void setWrapperCellStyle(TableTableCellElementBase element,
			WrapperCellStyle wrapperCellStyle) {
		Map<OdfStyleProperty, String> properties = OdsOdfdomStyleHelper.toProperties(wrapperCellStyle);
		element.setProperties(properties);
	}

	public void setWrapperCellStyle(OdfStyle style,
			WrapperCellStyle wrapperCellStyle) {
		Map<OdfStyleProperty, String> properties = OdsOdfdomStyleHelper.toProperties(wrapperCellStyle);
		style.setProperties(properties);
	}
}
