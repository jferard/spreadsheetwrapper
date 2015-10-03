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
package com.github.jferard.spreadsheetwrapper.ods.jopendocument;

import java.util.Map;

import org.jdom.Element;
import org.jdom.Namespace;

import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle.Color;
import com.github.jferard.spreadsheetwrapper.WrapperFont;

public class OdsJOpenUtility {
	private static Namespace foNS = Namespace.getNamespace("fo",
			"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
	private static Namespace officeNS = Namespace.getNamespace("office",
			"urn:oasis:names:tc:opendocument:xmlns:office:1.0");
	private static Namespace styleNS = Namespace.getNamespace("style",
			"urn:oasis:names:tc:opendocument:xmlns:style:1.0");

	public static Element createStyle(final String styleName,
			final WrapperCellStyle wrapperCellStyle) {
		final Element style = OdsJOpenUtility.getBaseStyle(styleName);
		final Color backgroundColor = wrapperCellStyle.getBackgroundColor();
		if (backgroundColor != null) {
			final Element tableCellProps = new Element("table-cell-properties",
					OdsJOpenUtility.styleNS);
			tableCellProps.setAttribute("background-color",
					backgroundColor.toHex(), OdsJOpenUtility.foNS);
			style.addContent(tableCellProps);
		}
		final WrapperFont cellFont = wrapperCellStyle.getCellFont();
		if (cellFont != null && cellFont.isBold()) {
			final Element textProps = new Element("text-properties",
					OdsJOpenUtility.styleNS);
			textProps.setAttribute("font-weight", "bold", OdsJOpenUtility.foNS);
			style.addContent(textProps);
		}
		return style;
	}

	public static Element createStyle(final String styleName,
			final Map<String, String> propertiesMap) {
		final Element style = OdsJOpenUtility.getBaseStyle(styleName);
		if (propertiesMap.containsKey("background-color")) {
			final Element tableCellProps = new Element("table-cell-properties",
					OdsJOpenUtility.styleNS);
			tableCellProps
			.setAttribute("background-color",
					propertiesMap.get("background-color"),
					OdsJOpenUtility.foNS);
			style.addContent(tableCellProps);
		}
		if (propertiesMap.containsKey("font-weight")) {
			final Element textProps = new Element("text-properties",
					OdsJOpenUtility.styleNS);
			textProps.setAttribute("font-weight",
					propertiesMap.get("font-weight"), OdsJOpenUtility.foNS);
			style.addContent(textProps);
		}
		return style;
	}

	private static Element getBaseStyle(final String styleName) {
		final Element style = new Element("style", OdsJOpenUtility.styleNS);
		style.setAttribute("name", styleName, OdsJOpenUtility.styleNS);
		style.setAttribute("family", "table-cell", OdsJOpenUtility.styleNS);
		return style;
	}

}
