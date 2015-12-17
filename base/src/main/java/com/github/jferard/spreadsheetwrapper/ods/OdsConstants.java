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
package com.github.jferard.spreadsheetwrapper.ods;

/**
 * Constants for ODS.
 */
public final class OdsConstants {
	/** for xml */

	/** name of the bold attribute in fo */
	public static final String BOLD_ATTR_VALUE = "bold";
	/** name of the italic attribute in fo */
	public static final String ITALIC_ATTR_VALUE = "italic";
	/** name of the normal attribute in fo */
	public static final String NORMAL_ATTR_VALUE = "normal";
	
	/** boolean type name */
	public static final String BOOLEAN_TYPE = "boolean";
	/** currency type name */
	public static final Object CURRENCY_TYPE = "currency";
	/** date type name */
	public static final String DATE_TYPE = "date";
	/** float type name */
	public static final String FLOAT_TYPE = "float";
	
	/** standard extension for files */
	public static final String EXTENSION = "ods";
	
	/** color of the cell in hex */
	public static final String BACKGROUND_COLOR_ATTR_NAME = "background-color";
	
	/** the border */
	public static final String BORDER_ATTR_NAME = "border";
	/** name of the color attribute in fo */
	public static final String COLOR_ATTR_NAME = "color";
	/** the font family, asian */
	public static final String FAMILY_ASIAN_ATTR_NAME = "font-family-asian";

	/** the font family */
	public static final String FAMILY_ATTR_NAME = "font-family";
	/** the font family, complex */
	public static final String FAMILY_COMPLEX_ATTR_NAME = "font-family-complex";

	/** namespaces XML */
	/** fo = ? */
	public static final String FO_NS_NAME = "fo";
	/** color of the font in hex */
	public static final String FONT_COLOR_ATTR_NAME = "font-color";

	/** size of the font */
	public static final String FONT_SIZE_ATTR_NAME = "font-size";
	/** style of the font, asian version */
	public static final String FONT_STYLE_ASIAN_ATTR_NAME = "font-style-asian";
	/** style of the font */
	public static final String FONT_STYLE_ATTR_NAME = "font-style";

	/** style of the font, complex version */
	public static final String FONT_STYLE_COMPLEX_ATTR_NAME = "font-style-complex";

	/** weight of the font asian version */
	public static final String FONT_WEIGHT_ASIAN_ATTR_NAME = "font-weight-asian";
	/** weight of the font */
	public static final String FONT_WEIGHT_ATTR_NAME = "font-weight";
	/** weight of the font complex version */
	public static final String FONT_WEIGHT_COMPLEX_ATTR_NAME = "font-weight-complex";

	/** name of the bold attribute in fo */
	public static final String FORMULA_ATTR_NAME = "formula";
	/** office */
	public static final String OFFICE_NS_NAME = "office";

	/** percentage type name */
	public static final Object PERCENTAGE_TYPE = "percentage";
	/** string type name */
	public static final String STRING_TYPE = "string";

	/** styles */
	public static final String STYLE_NS_NAME = "style";
	/** properties for a cell */
	public static final String TABLE_CELL_PROPERTIES_NAME = "table-cell-properties";
	/** properties for a text */
	public static final String TEXT_PROPERTIES_NAME = "text-properties";
	/** time type name */
	public static final String TIME_TYPE = "time";

	/**
	 * @param fontSize a string e.g. "10pt" or "5px"
	 * @return the font size in points, as a double
	 */
	public static double sizeToPoints(final String fontSize) {
		final double ret;
		final int length = fontSize.length();
		if (length > 2) {
			final String value = fontSize.substring(0, length - 2);
			final String unit = fontSize.substring(length - 2);
			final double tempValue = Float.valueOf(value);
			if ("in".equals(unit))
				ret = tempValue * 72.0;
			else if ("cm".equals(unit))
				ret = tempValue / 2.54 * 72.0;
			else if ("mm".equals(unit))
				ret = tempValue / 2.54 * 72.0;
			else if ("px".equals(unit))
				ret = tempValue;
			else if ("pc".equals(unit))
				ret = tempValue / 6.0 * 72.0;
			else
				ret = tempValue;
		} else
			ret = Double.valueOf(fontSize);
		return ret;
	}

	private OdsConstants() {
	}
}
