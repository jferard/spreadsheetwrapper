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
package com.github.jferard.spreadsheetwrapper.ods.${jopendocument.pkg};

import java.util.List;
import java.util.Map;

import org.jdom.Element;
import org.jdom.Namespace;
import org.jopendocument.dom.ODXMLDocument;
import org.jopendocument.dom.StyledNode;
import org.jopendocument.dom.spreadsheet.CellStyle;
import org.jopendocument.dom.spreadsheet.CellStyle.${jopendocument.styletablecellproperties.cls};
import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.SpreadSheet;
import org.jopendocument.dom.text.TextStyle.${jopendocument.styletextproperties.cls};

import com.github.jferard.spreadsheetwrapper.StyleHelper;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;
import com.github.jferard.spreadsheetwrapper.style.Borders;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;

class OdsJOpenStyleHelper implements StyleHelper<CellStyle, StyledNode<CellStyle, SpreadSheet>> {
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
	 * XPath could do better...
	 *
	 * @param name
	 *            the name to find in attributes
	 * @param elements
	 *            the elements to look at
	 * @return the matching element, or null
	 */
	private static/*@Nullable*/Element findElementWithName(final String name,
			final List<Element> elements) {
		for (final Element element : elements) {
			if (name.equals(element.getAttributeValue("name")))
				return element;
		}
		return null;
	}

	/**
	 * @param styleName
	 *            name of the style
	 * @return the base Element, which has to be
	 */
	private Element createEmptyStyleElement(final String styleName) {
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
	private Element createStyleElement(final String styleName,
			final WrapperCellStyle wrapperCellStyle) {
		final Element style = this.createBaseStyleElement(styleName);
		this.fillStyle(style, wrapperCellStyle);
		return style;
	}

	public boolean addStyle(ODXMLDocument stylesDoc, String styleName, 
			WrapperCellStyle wrapperCellStyle) {
		final Element namedStyles = stylesDoc.getChild("styles", true);
		@SuppressWarnings("unchecked")
		final List<Element> styles = namedStyles.getChildren("style");
		Element newStyle = OdsJOpenStyleHelper.findElementWithName(
				styleName, styles);
		if (newStyle == null) {
			newStyle = this.createEmptyStyleElement(styleName);
			namedStyles.addContent(newStyle);
		} else
			newStyle.removeContent();

		this.fillStyle(newStyle, wrapperCellStyle);
		return true;
	}
	
	
	/**
	 * @param cellStyle
	 * 	          the internal cell style, to update
	 * @param wrapperStyle
	 *            the wrapper cell
	 */
	private void setCellStyle(final CellStyle cellStyle,
			final WrapperCellStyle wrapperCellStyle) {
		final ${jopendocument.styletablecellproperties.cls} tableCellProperties = cellStyle
				.getTableCellProperties();
		final Element tableCellPropertiesElement = tableCellProperties
				.getElement();
		final ${jopendocument.styletextproperties.cls} textProperties = cellStyle.getTextProperties();
		final Element textPropertiesElement = textProperties.getElement();
		this.setStyle(tableCellPropertiesElement, textPropertiesElement,
				wrapperCellStyle);
	}

	/**
	 * @param style
	 *            the style Node
	 * @param wrapperCellStyle
	 *            the style to set
	 * @return the style Node itself
	 */
	private Element fillStyle(final Element style,
			final WrapperCellStyle wrapperCellStyle) {
		style.setAttribute("family", "table-cell", OdsJOpenStyleHelper.STYLE_NS);
		final Element tableCellProps = new Element(
				OdsConstants.TABLE_CELL_PROPERTIES_NAME,
				OdsJOpenStyleHelper.STYLE_NS);
		final Element textProps = new Element(
				OdsConstants.TEXT_PROPERTIES_NAME, OdsJOpenStyleHelper.STYLE_NS);
		this.setStyle(tableCellProps, textProps, wrapperCellStyle);
		if (!tableCellProps.getAttributes().isEmpty())
			style.addContent(tableCellProps);
		if (!textProps.getAttributes().isEmpty())
			style.addContent(textProps);
		return style;
	}

	/**
	 * @param cellStyle
	 *            the internal cell style
	 * @return the wrapper cell style
	 */
	@Override
	public WrapperCellStyle toWrapperCellStyle(
			final /*@Nullable*/ CellStyle cellStyle) {
		if (cellStyle == null)
			return WrapperCellStyle.EMPTY;

		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle();
		final ${jopendocument.styletablecellproperties.cls} tableCellProperties = cellStyle
				.getTableCellProperties();
		final String bColorAsHex = tableCellProperties.getRawBackgroundColor();
		if (bColorAsHex != null) {
			final WrapperColor backgroundColor = WrapperColor
					.stringToColor(bColorAsHex);
			wrapperCellStyle.setBackgroundColor(backgroundColor);
		}
		
		final Element cellPropertiesElment = tableCellProperties.getElement();  
		final Borders borders = OdsJOpenStyleBorderHelper.toBorders(cellPropertiesElment);
		if (borders != null)
			wrapperCellStyle.setBorders(borders);

		final ${jopendocument.styletextproperties.cls} textProperties = cellStyle.getTextProperties();
		final Element odfElement = textProperties.getElement();

		final WrapperFont wrapperFont = OdsJOpenStyleFontHelper.toWrapperFont(odfElement);
		if (wrapperFont != null)
			wrapperCellStyle.setCellFont(wrapperFont);
		return wrapperCellStyle;
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
		final Element style = this.createBaseStyleElement(styleName);
		if (propertiesMap.containsKey(OdsConstants.BACKGROUND_COLOR_ATTR_NAME)) {
			final String backgroundColor = propertiesMap
					.get(OdsConstants.BACKGROUND_COLOR_ATTR_NAME);
			final Element tableCellProps = new Element(
					OdsConstants.TABLE_CELL_PROPERTIES_NAME,
					OdsJOpenStyleHelper.STYLE_NS);
			tableCellProps.setAttribute(OdsConstants.BACKGROUND_COLOR_ATTR_NAME,
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
	private Element createBaseStyleElement(final String styleName) {
		final Element style = this.createEmptyStyleElement(styleName);
		style.setAttribute("family", "table-cell", OdsJOpenStyleHelper.STYLE_NS);
		return style;
	}

	private void setBackgroundColor(final Element tableCellProps,
			final WrapperColor backgroundColor) {
		if (backgroundColor != null) {
			tableCellProps.setAttribute(OdsConstants.BACKGROUND_COLOR_ATTR_NAME,
					backgroundColor.toHex(), OdsJOpenStyleHelper.FO_NS);
		}
	}

	private void setStyle(final Element tableCellProps,
			final Element textProps, final WrapperCellStyle wrapperCellStyle) {
		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();
		this.setBackgroundColor(tableCellProps, backgroundColor);

		final /*@Nullable*/ Borders borders = wrapperCellStyle.getBorders();
		if (borders != null)
			OdsJOpenStyleBorderHelper.setBorders(tableCellProps, borders);

		final /*@Nullable*/ WrapperFont cellFont = wrapperCellStyle.getCellFont();
		if (cellFont != null) {
			OdsJOpenStyleFontHelper.setFont(textProps, cellFont);
		}
	}

	/** {@inheritDoc} */
	@Override
	public WrapperCellStyle getWrapperCellStyle(StyledNode<CellStyle, SpreadSheet> node) {
		if (node == null)
			return null;
		
		CellStyle style = node.getStyle();
		return style == null ? WrapperCellStyle.EMPTY : this.toWrapperCellStyle(style);
	}

	public String getStyleName(MutableCell<SpreadSheet> cell) {
		if (cell == null)
			return null;
		
		return cell.getStyleName();
	}

	@Override
	public void setWrapperCellStyle(StyledNode<CellStyle, SpreadSheet> element,
			WrapperCellStyle wrapperCellStyle) {
		CellStyle cellStyle = element.getStyle();
		if (cellStyle == null)
			cellStyle = element.getPrivateStyle();

		if (cellStyle != null)
		this.setCellStyle(cellStyle, wrapperCellStyle);
	}
	
	public boolean setStyleName(StyledNode<CellStyle, SpreadSheet> element,
			String styleName) {
		element.setStyleName(styleName);
		return true;
	}
	
	public /*@Nullable*/WrapperCellStyle getStyle(StyledNode<CellStyle, SpreadSheet> node) {
		if (node == null)
			return null;
		
		return this.getWrapperCellStyle(node);
	}
	
	public boolean setStyle(StyledNode<CellStyle, SpreadSheet> element,
			WrapperCellStyle wrapperCellStyle) {
		this.setWrapperCellStyle(element, wrapperCellStyle);
		return true;
	}
	
	
}