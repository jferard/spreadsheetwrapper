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
package com.github.jferard.spreadsheetwrapper.ods.jopendocument13;

import java.util.Map;

import org.jdom.Element;
import org.jdom.Namespace;

import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;
import com.github.jferard.spreadsheetwrapper.impl.StyleUtility;

public class OdsJOpenStyleHelper extends StyleUtility {
	/** the name space fo (fonts ?) */
	public static final Namespace foNS = Namespace.getNamespace("fo",
			"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
	/** the name space office (table, row, cell, etc.) */
	public static final Namespace officeNS = Namespace.getNamespace("office",
			"urn:oasis:names:tc:opendocument:xmlns:office:1.0");
	/** the name space style (styles) */
	public static final Namespace styleNS = Namespace.getNamespace("style",
			"urn:oasis:names:tc:opendocument:xmlns:style:1.0");

	/**
	 * @param styleName
	 *            the name of the style
	 * @return the DOM element, with name attribute only
	 */
	public Element createBaseStyle(final String styleName) {
		final Element style = new Element("style", OdsJOpenStyleHelper.styleNS);
		style.setAttribute("name", styleName, OdsJOpenStyleHelper.styleNS);
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
			final Element tableCellProps = new Element("table-cell-properties",
					OdsJOpenStyleHelper.styleNS);
			tableCellProps.setAttribute(StyleUtility.BACKGROUND_COLOR,
					backgroundColor.toHex(), OdsJOpenStyleHelper.foNS);
			style.addContent(tableCellProps);
		}
		final WrapperFont cellFont = wrapperCellStyle.getCellFont();
		if (cellFont != null && cellFont.getBold() == WrapperCellStyle.YES) {
			final Element textProps = new Element("text-properties",
					OdsJOpenStyleHelper.styleNS);
			textProps.setAttribute(StyleUtility.FONT_WEIGHT, "bold",
					OdsJOpenStyleHelper.foNS);
			style.addContent(textProps);
		}
		return style;
	}

	public Element setStyle(final Element style,
			final WrapperCellStyle wrapperCellStyle) {
		style.setAttribute("family", "table-cell", OdsJOpenStyleHelper.styleNS);
		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();
		if (backgroundColor != null) {
			final Element tableCellProps = new Element("table-cell-properties",
					OdsJOpenStyleHelper.styleNS);
			tableCellProps.setAttribute(StyleUtility.BACKGROUND_COLOR,
					backgroundColor.toHex(), OdsJOpenStyleHelper.foNS);
			style.addContent(tableCellProps);
		}
		final WrapperFont cellFont = wrapperCellStyle.getCellFont();
		if (cellFont != null && cellFont.getBold() == WrapperCellStyle.YES) {
			final Element textProps = new Element("text-properties",
					OdsJOpenStyleHelper.styleNS);
			textProps.setAttribute(StyleUtility.FONT_WEIGHT, "bold",
					OdsJOpenStyleHelper.foNS);
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
		if (propertiesMap.containsKey(StyleUtility.BACKGROUND_COLOR)) {
			final Element tableCellProps = new Element("table-cell-properties",
					OdsJOpenStyleHelper.styleNS);
			tableCellProps.setAttribute(StyleUtility.BACKGROUND_COLOR,
					propertiesMap.get(StyleUtility.BACKGROUND_COLOR),
					OdsJOpenStyleHelper.foNS);
			style.addContent(tableCellProps);
		}
		if (propertiesMap.containsKey(StyleUtility.FONT_WEIGHT)) {
			final Element textProps = new Element("text-properties",
					OdsJOpenStyleHelper.styleNS);
			textProps.setAttribute(StyleUtility.FONT_WEIGHT,
					propertiesMap.get(StyleUtility.FONT_WEIGHT),
					OdsJOpenStyleHelper.foNS);
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
		style.setAttribute("family", "table-cell", OdsJOpenStyleHelper.styleNS);
		return style;
	}

	private Element setStyle(final Element style,
			final Map<String, String> propertiesMap) {
		style.setAttribute("family", "table-cell", OdsJOpenStyleHelper.styleNS);
		if (propertiesMap.containsKey(StyleUtility.BACKGROUND_COLOR)) {
			final Element tableCellProps = new Element("table-cell-properties",
					OdsJOpenStyleHelper.styleNS);
			tableCellProps.setAttribute(StyleUtility.BACKGROUND_COLOR,
					propertiesMap.get(StyleUtility.BACKGROUND_COLOR),
					OdsJOpenStyleHelper.foNS);
			style.addContent(tableCellProps);
		}
		if (propertiesMap.containsKey(StyleUtility.FONT_WEIGHT)) {
			final Element textProps = new Element("text-properties",
					OdsJOpenStyleHelper.styleNS);
			textProps.setAttribute(StyleUtility.FONT_WEIGHT,
					propertiesMap.get(StyleUtility.FONT_WEIGHT),
					OdsJOpenStyleHelper.foNS);
			style.addContent(textProps);
		}
		return style;
	}

}