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
package com.github.jferard.spreadsheetwrapper.ods.jopendocument1_2;

import java.util.Map;

import org.jdom.Element;
import org.jdom.Namespace;

import com.github.jferard.spreadsheetwrapper.StyleUtility;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;

class OdsJOpenStyleHelper {
	/** the name space fo (fonts ?) */
	public static final Namespace FO_NS = Namespace.getNamespace(OdsConstants.FO_NS_NAME,
			"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
	/** the name space office (table, row, cell, etc.) */
	public static final Namespace OFFICE_NS = Namespace.getNamespace(OdsConstants.OFFICE_NS_NAME,
			"urn:oasis:names:tc:opendocument:xmlns:office:1.0");
	/** the name space style (styles) */
	public static final Namespace STYLE_NS = Namespace.getNamespace(OdsConstants.STYLE_NS_NAME,
			"urn:oasis:names:tc:opendocument:xmlns:style:1.0");

	/**
	 * @param styleName
	 *            name of the style
	 * @return the base Element, which has to be
	 */
	public Element createBaseStyle(final String styleName) {
		final Element style = new Element("style", OdsJOpenStyleHelper.STYLE_NS);
		style.setAttribute("name", styleName, OdsJOpenStyleHelper.STYLE_NS);
		return style;
	}

	/**
	 * @param styleName
	 *            the name of the style
	 * @param styleString
	 *            the new style format
	 * @return the DOM element
	 */
	public Element createStyleElement(final String styleName,
			final WrapperCellStyle wrapperCellStyle) {
		final Element style = this.getBaseStyle(styleName);
		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();
		if (backgroundColor != null) {
			final Element tableCellProps = new Element(
					OdsConstants.TABLE_CELL_PROPERTIES_NAME,
					OdsJOpenStyleHelper.STYLE_NS);
			tableCellProps.setAttribute(OdsConstants.BACKGROUND_COLOR,
					backgroundColor.toHex(), OdsJOpenStyleHelper.FO_NS);
			style.addContent(tableCellProps);
		}
		final WrapperFont cellFont = wrapperCellStyle.getCellFont();
		if (cellFont != null && cellFont.getBold() == WrapperCellStyle.YES) {
			final Element textProps = new Element(
					OdsConstants.TEXT_PROPERTIES_NAME,
					OdsJOpenStyleHelper.STYLE_NS);
			textProps.setAttribute(OdsConstants.FONT_WEIGHT, "bold",
					OdsJOpenStyleHelper.FO_NS);
			style.addContent(textProps);
		}
		return style;
	}

	/**
	 * @param style
	 *            the style Node
	 * @param wrapperCellStyle
	 *            the style to set
	 * @return the style Node itself
	 */
	public Element setStyle(final Element style,
			final WrapperCellStyle wrapperCellStyle) {
		style.setAttribute("family", "table-cell", OdsJOpenStyleHelper.STYLE_NS);
		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();
		if (backgroundColor != null) {
			final Element tableCellProps = new Element(
					OdsConstants.TABLE_CELL_PROPERTIES_NAME,
					OdsJOpenStyleHelper.STYLE_NS);
			tableCellProps.setAttribute(OdsConstants.BACKGROUND_COLOR,
					backgroundColor.toHex(), OdsJOpenStyleHelper.FO_NS);
			style.addContent(tableCellProps);
		}
		final WrapperFont cellFont = wrapperCellStyle.getCellFont();
		if (cellFont != null && cellFont.getBold() == WrapperCellStyle.YES) {
			final Element textProps = new Element(
					OdsConstants.TEXT_PROPERTIES_NAME,
					OdsJOpenStyleHelper.STYLE_NS);
			textProps.setAttribute(OdsConstants.FONT_WEIGHT, "bold",
					OdsJOpenStyleHelper.FO_NS);
			style.addContent(textProps);
		}
		return style;
	}

	/**
	 * @param styleName
	 *            the name of the style
	 * @param propertiesMap
	 *            a map key-value of the properties
	 * @return the DOM element
	 */
	private Element createStyle(final String styleName,
			final Map<String, String> propertiesMap) {
		final Element style = this.getBaseStyle(styleName);
		if (propertiesMap.containsKey(OdsConstants.BACKGROUND_COLOR)) {
			final String backgroundColor = propertiesMap
					.get(OdsConstants.BACKGROUND_COLOR);
			final Element tableCellProps = new Element(
					OdsConstants.TABLE_CELL_PROPERTIES_NAME,
					OdsJOpenStyleHelper.STYLE_NS);
			tableCellProps.setAttribute(OdsConstants.BACKGROUND_COLOR,
					backgroundColor, OdsJOpenStyleHelper.FO_NS);
			style.addContent(tableCellProps);
		}
		if (propertiesMap.containsKey(OdsConstants.FONT_WEIGHT)) {
			final String fontWeight = propertiesMap
					.get(OdsConstants.FONT_WEIGHT);
			final Element textProps = new Element(
					OdsConstants.TEXT_PROPERTIES_NAME,
					OdsJOpenStyleHelper.STYLE_NS);
			textProps.setAttribute(OdsConstants.FONT_WEIGHT, fontWeight,
					OdsJOpenStyleHelper.FO_NS);
			style.addContent(textProps);
		}
		return style;
	}

	/**
	 * @param styleName
	 *            the name of the style
	 * @return the base of the DOM element for a style
	 */
	private Element getBaseStyle(final String styleName) {
		final Element style = this.createBaseStyle(styleName);
		style.setAttribute("family", "table-cell", OdsJOpenStyleHelper.STYLE_NS);
		return style;
	}

	private Element setStyle(final Element style,
			final Map<String, String> propertiesMap) {
		style.setAttribute("family", "table-cell", OdsJOpenStyleHelper.STYLE_NS);
		if (propertiesMap.containsKey(OdsConstants.BACKGROUND_COLOR)) {
			final String backgroundColorAsHex = propertiesMap
					.get(OdsConstants.BACKGROUND_COLOR);
			final Element tableCellProps = new Element(
					OdsConstants.TABLE_CELL_PROPERTIES_NAME,
					OdsJOpenStyleHelper.STYLE_NS);
			tableCellProps.setAttribute(OdsConstants.BACKGROUND_COLOR,
					backgroundColorAsHex, OdsJOpenStyleHelper.FO_NS);
			style.addContent(tableCellProps);
		}
		if (propertiesMap.containsKey(OdsConstants.FONT_WEIGHT)) {
			final String fontWeight = propertiesMap
					.get(OdsConstants.FONT_WEIGHT);
			final Element textProps = new Element(
					OdsConstants.TEXT_PROPERTIES_NAME,
					OdsJOpenStyleHelper.STYLE_NS);
			textProps.setAttribute(OdsConstants.FONT_WEIGHT, fontWeight,
					OdsJOpenStyleHelper.FO_NS);
			style.addContent(textProps);
		}
		return style;
	}

}