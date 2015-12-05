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
package com.github.jferard.spreadsheetwrapper.ods.jopendocument${jopendocument.version};

import java.util.Map;

import org.jdom.Element;
import org.jdom.Namespace;
import org.jopendocument.dom.spreadsheet.CellStyle;
import org.jopendocument.dom.spreadsheet.CellStyle.${jopendocument.style}TableCellProperties;
import org.jopendocument.dom.text.TextStyle.${jopendocument.style}TextProperties;

import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;

class OdsJOpenStyleHelper {
	/** the name space fo (fonts ?) */
	public static final Namespace FO_NS = Namespace.getNamespace(
			OdsConstants.FO_NS_NAME,
			"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
	/** the name space office (table, row, cell, etc.) */
	public static final Namespace OFFICE_NS = Namespace.getNamespace(
			OdsConstants.OFFICE_NS_NAME,
			"urn:oasis:names:tc:opendocument:xmlns:office:1.0");
	/** the name space style (styles) */
	public static final Namespace STYLE_NS = Namespace.getNamespace(
			OdsConstants.STYLE_NS_NAME,
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
			textProps.setAttribute(OdsConstants.FONT_WEIGHT_ATTR_NAME, "bold",
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
		if (cellFont != null) {
			final Element textProps = new Element(
					OdsConstants.TEXT_PROPERTIES_NAME,
					OdsJOpenStyleHelper.STYLE_NS);
			this.setBold(textProps, cellFont.getBold());
			this.setItalic(textProps, cellFont.getItalic());
			this.setSize(textProps, cellFont.getSize());
			final WrapperColor color = cellFont.getColor();
			if (color != null)
				this.setColor(textProps, color);
			style.addContent(textProps);
		}
		return style;
	}

	/**
	 * @param cellStyle the internal cell style
	 * @return the wrapper cell style
	 */
	public WrapperCellStyle toWrapperCellStyle(final /*@Nullable*/ CellStyle cellStyle) {
		if (cellStyle == null)
			return WrapperCellStyle.EMPTY;

		final ${jopendocument.style}TableCellProperties tableCellProperties = cellStyle
				.getTableCellProperties();
		final String bColorAsHex = tableCellProperties.getRawBackgroundColor();
		final WrapperColor backgroundColor = WrapperColor
				.stringToColor(bColorAsHex);

		final WrapperFont wrapperFont = new WrapperFont();
		final ${jopendocument.style}TextProperties textProperties = cellStyle.getTextProperties();
		final Element odfElement = textProperties.getElement();

		final String fontWeight = odfElement.getAttributeValue(
				OdsConstants.FONT_WEIGHT_ATTR_NAME, OdsJOpenStyleHelper.FO_NS);
		if ("bold".equals(fontWeight))
			wrapperFont.setBold();
		else if ("normal".equals(fontWeight))
			wrapperFont.setBold(WrapperCellStyle.NO);

		final String fontStyle = odfElement.getAttributeValue(
				OdsConstants.FONT_STYLE_ATTR_NAME, OdsJOpenStyleHelper.FO_NS);
		if ("italic".equals(fontStyle))
			wrapperFont.setItalic();
		else if ("normal".equals(fontStyle))
			wrapperFont.setItalic(WrapperCellStyle.NO);

		final String fontSize = odfElement.getAttributeValue(
				OdsConstants.FONT_SIZE, OdsJOpenStyleHelper.FO_NS);
		if (fontSize != null)
			wrapperFont.setSize(OdsConstants.sizeToPoints(fontSize));

		final String fColorAsHex = odfElement.getAttributeValue(
				OdsConstants.COLOR_ATTR_NAME, OdsJOpenStyleHelper.FO_NS);
		final WrapperColor fontColor = WrapperColor.stringToColor(fColorAsHex);
		if (fontColor != null)
			wrapperFont.setColor(fontColor);

		return new WrapperCellStyle(backgroundColor, wrapperFont);
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
		if (propertiesMap.containsKey(OdsConstants.FONT_WEIGHT_ATTR_NAME)) {
			final String fontWeight = propertiesMap
					.get(OdsConstants.FONT_WEIGHT_ATTR_NAME);
			final Element textProps = new Element(
					OdsConstants.TEXT_PROPERTIES_NAME,
					OdsJOpenStyleHelper.STYLE_NS);
			textProps.setAttribute(OdsConstants.FONT_WEIGHT_ATTR_NAME, fontWeight,
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

	private void setBold(final Element textProps, final int key) {
		final String value;
		if (key == WrapperCellStyle.YES)
			value = "bold";
		else if (key == WrapperCellStyle.YES)
			value = "normal";
		else
			value = null;

		if (value != null) {
			textProps.setAttribute(OdsConstants.FONT_WEIGHT_ATTR_NAME, value,
					OdsJOpenStyleHelper.FO_NS);
			textProps.setAttribute(OdsConstants.FONT_WEIGHT_ASIAN_ATTR_NAME, value,
					OdsJOpenStyleHelper.STYLE_NS);
			textProps.setAttribute(OdsConstants.FONT_WEIGHT_COMPLEX_ATTR_NAME, value,
					OdsJOpenStyleHelper.STYLE_NS);
		}
	}

	private void setColor(final Element textProps, final WrapperColor color) {
		if (color != null) {
			textProps.setAttribute(OdsConstants.COLOR_ATTR_NAME, color.toHex(),
					OdsJOpenStyleHelper.FO_NS);
		}
	}

	private void setItalic(final Element textProps, final int key) {
		final String value;
		if (key == WrapperCellStyle.YES)
			value = "italic";
		else if (key == WrapperCellStyle.YES)
			value = "normal";
		else
			value = null;

		if (value != null) {
			textProps.setAttribute(OdsConstants.FONT_STYLE_ATTR_NAME, value,
					OdsJOpenStyleHelper.FO_NS);
			textProps.setAttribute(OdsConstants.FONT_STYLE_ASIAN_ATTR_NAME, value,
					OdsJOpenStyleHelper.STYLE_NS);
			textProps.setAttribute(OdsConstants.FONT_STYLE_COMPLEX_ATTR_NAME, value,
					OdsJOpenStyleHelper.STYLE_NS);
		}
	}

	private void setSize(final Element textProps, final int size) {
		if (size != WrapperCellStyle.DEFAULT) {
			textProps.setAttribute(OdsConstants.FONT_SIZE, Integer.toString(size)+"pt",
					OdsJOpenStyleHelper.FO_NS);
		}
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
		if (propertiesMap.containsKey(OdsConstants.FONT_WEIGHT_ATTR_NAME)) {
			final String fontWeight = propertiesMap
					.get(OdsConstants.FONT_WEIGHT_ATTR_NAME);
			final Element textProps = new Element(
					OdsConstants.TEXT_PROPERTIES_NAME,
					OdsJOpenStyleHelper.STYLE_NS);
			textProps.setAttribute(OdsConstants.FONT_WEIGHT_ATTR_NAME, fontWeight,
					OdsJOpenStyleHelper.FO_NS);
			style.addContent(textProps);
		}
		return style;
	}


	/**
	 * @param cellStyle the internal cell style, to update
	 * @param wrapperStyle the wrapper cell
	 */
	public boolean setCellStyle(final CellStyle cellStyle, final WrapperCellStyle wrapperStyle) {
		final WrapperColor backgroundColor = wrapperStyle.getBackgroundColor();
		if (backgroundColor != null) {
			final ${jopendocument.style}TableCellProperties tableCellProperties = cellStyle
					.getTableCellProperties();
			tableCellProperties.setBackgroundColor(backgroundColor.toHex());
		}
		final WrapperFont font = wrapperStyle.getCellFont();
		if (font != null) {
			final ${jopendocument.style}TextProperties textProperties = cellStyle
					.getTextProperties();
			final Element textPropertiesElement = textProperties.getElement();
			this.setColor(textPropertiesElement, font.getColor());
			this.setBold(textPropertiesElement, font.getBold());
			this.setItalic(textPropertiesElement, font.getItalic());
			this.setSize(textPropertiesElement, font.getSize());
		}
		return true;
	}
	//
	//	/**
	//	 * <style:text-properties fo:font-size="14pt" fo:font-style="italic" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color" fo:font-weight="bold" style:font-size-asian="14pt" style:font-style-asian="italic" style:font-weight-asian="bold" style:font-size-complex="14pt" style:font-style-complex="italic" style:font-weight-complex="bold"/>
	//	 * <style:text-properties fo:font-size="10pt" fo:font-style="normal" style:text-underline-style="none" fo:font-weight="normal" style:font-size-asian="10pt" style:font-style-asian="normal" style:font-weight-asian="normal" style:font-size-complex="10pt" style:font-style-complex="normal" style:font-weight-complex="normal"/>
	//	 */
	//	private void setBold(final ${jopendocument.style}TextProperties textProperties, int boldKey) {
	//		final String value;
	//		switch (boldKey) {
	//			case WrapperCellStyle.YES: value = "bold"; break;
	//			case WrapperCellStyle.NO: value = "normal"; break;
	//			default: value = null; break;
	//		}
	//
	//		if (value != null) {
	//			textProperties.getElement().setAttribute(OdsConstants.FONT_WEIGHT_ATTR_NAME, value,
	//					OdsJOpenStyleHelper.FO_NS);
	//			textProperties.getElement().setAttribute(OdsConstants.FONT_WEIGHT_ASIAN_ATTR_NAME, value,
	//					OdsJOpenStyleHelper.STYLE_NS);
	//			textProperties.getElement().setAttribute(OdsConstants.FONT_WEIGHT_COMPLEX_ATTR_NAME, value,
	//					OdsJOpenStyleHelper.STYLE_NS);
	//		}
	//	}
	//
	//	private void setItalic(final ${jopendocument.style}TextProperties textProperties, int italicKey) {
	//		final String value;
	//		switch (italicKey) {
	//			case WrapperCellStyle.YES: value = "italic"; break;
	//			case WrapperCellStyle.NO: value = "normal"; break;
	//			default: value = null; break;
	//		}
	//
	//		if (value != null) {
	//			textProperties.getElement().setAttribute(OdsConstants.FONT_STYLE_ATTR_NAME, value,
	//					OdsJOpenStyleHelper.FO_NS);
	//			textProperties.getElement().setAttribute(OdsConstants.FONT_STYLE_ASIAN_ATTR_NAME, value,
	//					OdsJOpenStyleHelper.STYLE_NS);
	//			textProperties.getElement().setAttribute(OdsConstants.FONT_STYLE_COMPLEX_ATTR_NAME, value,
	//					OdsJOpenStyleHelper.STYLE_NS);
	//		}
	//	}
	//
	//	private void setSize(final ${jopendocument.style}TextProperties textProperties, int size) {
	//		if (size != WrapperCellStyle.DEFAULT) {
	//			textProperties.getElement().setAttribute(
	//					OdsConstants.FONT_SIZE, Integer.toString(size)+"pt", OdsJOpenStyleHelper.FO_NS);
	//		}
	//	}

}