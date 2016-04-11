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
package com.github.jferard.spreadsheetwrapper.ods.apache;

import java.util.HashMap;
import java.util.Map;

import org.odftoolkit.odfdom.dom.element.style.StyleStyleElement;
import org.odftoolkit.odfdom.dom.element.table.TableTableCellElementBase;
import org.odftoolkit.odfdom.dom.style.OdfStyleFamily;
import org.odftoolkit.odfdom.dom.style.OdfStylePropertySet;
import org.odftoolkit.odfdom.dom.style.props.OdfStyleProperty;
import org.odftoolkit.odfdom.dom.style.props.OdfTableCellProperties;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;

import com.github.jferard.spreadsheetwrapper.StyleHelper;
import com.github.jferard.spreadsheetwrapper.style.Borders;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;

/**
 * A little style utility for odftoolkit files
 *
 */
public class OdsApacheStyleHelper implements
		StyleHelper<OdfStylePropertySet, TableTableCellElementBase> {
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
			properties.putAll(OdsApacheStyleFontHelper
					.getFontProperties(wrapperFont));

		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();
		if (backgroundColor != null) {
			properties.put(OdfTableCellProperties.BackgroundColor,
					backgroundColor.toHex());
		}
		final Borders borders = wrapperCellStyle.getBorders();
		if (borders != null)
			properties.putAll(OdsApacheStyleBorderHelper
					.getBordersProperties(borders));

		return properties;
	}

	private static WrapperCellStyle toWrapperCellStyle(
			Map<OdfStyleProperty, String> propertyByAttrName) {
		WrapperFont wrapperFont = OdsApacheStyleFontHelper.toWrapperCellFont(propertyByAttrName);
		final WrapperCellStyle wrapperStyle = new WrapperCellStyle()
				.setCellFont(wrapperFont);

		final String backgroundColor = propertyByAttrName
				.get(OdfTableCellProperties.BackgroundColor);
		if (backgroundColor != null)
			wrapperStyle.setBackgroundColor(WrapperColor.stringToColor(backgroundColor));

		Borders borders = OdsApacheStyleBorderHelper
				.toCellBorders(propertyByAttrName);
		wrapperStyle.setBorders(borders);

		return wrapperStyle;
	}

	public WrapperCellStyle getCellStyle(final OdfOfficeStyles documentStyles, final String styleName) {
		final OdfStyle existingStyle = documentStyles.getStyle(styleName,
				OdfStyleFamily.TableCell);
		return this.toWrapperCellStyle(existingStyle);
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

	public boolean setStyle(OdfOfficeStyles documentStyles, String styleName,
			WrapperCellStyle wrapperCellStyle) {
		final OdfStyle newStyle = documentStyles.newStyle(styleName,
				OdfStyleFamily.TableCell);
		this.setWrapperCellStyle(newStyle, wrapperCellStyle);
		return true;
	}

	public void setWrapperCellStyle(OdfStyle style,
			WrapperCellStyle wrapperCellStyle) {
		Map<OdfStyleProperty, String> properties = OdsApacheStyleHelper
				.toProperties(wrapperCellStyle);
		style.setProperties(properties);
	}
	
	@Override
	public void setWrapperCellStyle(TableTableCellElementBase element,
			WrapperCellStyle wrapperCellStyle) {
		Map<OdfStyleProperty, String> properties = OdsApacheStyleHelper
				.toProperties(wrapperCellStyle);
		element.setProperties(properties);
	}

	/**
	 * @param style
	 *            *internal* cell style
	 * @return the new cell style format
	 */
	@Override
	public WrapperCellStyle toWrapperCellStyle(final OdfStylePropertySet style) {
		return OdsApacheStyleHelper.toWrapperCellStyle(style.getProperties(style
				.getStrictProperties()));
	}
}	