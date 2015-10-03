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

import com.github.jferard.spreadsheetwrapper.CellStyle;
import com.github.jferard.spreadsheetwrapper.CellStyle.Color;
import com.github.jferard.spreadsheetwrapper.Font;

public class OdsJOpenUtility {
	private static Namespace officeNS = Namespace.getNamespace("office",
			"urn:oasis:names:tc:opendocument:xmlns:office:1.0");
	private static Namespace foNS = Namespace.getNamespace("fo",
			"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
	private static Namespace styleNS = Namespace.getNamespace("style",
			"urn:oasis:names:tc:opendocument:xmlns:style:1.0");

	public static Element createStyle(String styleName,
			Map<String, String> propertiesMap) {
		Element style = getBaseStyle(styleName);
		if (propertiesMap.containsKey("background-color")) {
			Element tableCellProps = new Element("table-cell-properties",
					styleNS);
			tableCellProps.setAttribute("background-color",
					propertiesMap.get("background-color"), foNS);
			style.addContent(tableCellProps);
		}
		if (propertiesMap.containsKey("font-weight")) {
			Element textProps = new Element("text-properties", styleNS);
			textProps.setAttribute("font-weight",
					propertiesMap.get("font-weight"), foNS);
			style.addContent(textProps);
		}
		return style;
	}

	private static Element getBaseStyle(String styleName) {
		Element style = new Element("style", styleNS);
		style.setAttribute("name", styleName, styleNS);
		style.setAttribute("family", "table-cell", styleNS);
		return style;
	}

	public static Element createStyle(String styleName, CellStyle cellStyle) {
		Element style = getBaseStyle(styleName);
		final Color backgroundColor = cellStyle.getBackgroundColor();
		if (backgroundColor != null) {
			Element tableCellProps = new Element("table-cell-properties",
					styleNS);
			tableCellProps.setAttribute("background-color",
					backgroundColor.toHex(), foNS);
			style.addContent(tableCellProps);
		}
		final Font cellFont = cellStyle.getCellFont();
		if (cellFont != null && cellFont.isBold()) {
			Element textProps = new Element("text-properties", styleNS);
			textProps.setAttribute("font-weight", "bold", foNS);
			style.addContent(textProps);
		}
		return style;
	}

}
