package com.github.jferard.spreadsheetwrapper.ods;

public final class OdsConstants {
	private OdsConstants() {}
	
	/** boolean type name */
	public static final String BOOLEAN_TYPE = "boolean";
	/** float type name */
	public static final String FLOAT_TYPE = "float";
	/** string type name */
	public static final String STRING_TYPE = "string";
	/** date type name */
	public static final String DATE_TYPE = "date";
	/** time type name */
	public static final String TIME_TYPE = "time";
	/** currency type name */
	public static final Object CURRENCY_TYPE = "currency";	
	/** percentage type name */
	public static final Object PERCENTAGE_TYPE = "percentage";
	
	/** for xml */
	/** color of the cell in hex */
	public static final String BACKGROUND_COLOR = "background-color";
	/** color of the font in hex */
	public static final String FONT_COLOR = "font-color";
	/** size of the font */
	public static final String FONT_SIZE = "font-size";
	/** style of the font */
	public static final String FONT_STYLE = "font-style";
	/** weight of the font */
	public static final String FONT_WEIGHT = "font-weight";
	
	/** namespaces XML */
	/** fo = ? */
	public static final String FO_NS_NAME = "fo";
	/** office */
	public static final String OFFICE_NS_NAME = "office";
	/** styles */
	public static final String STYLE_NS_NAME = "style";
	
	/** properties for a cell */
	public static final String TABLE_CELL_PROPERTIES_NAME = "table-cell-properties";
	/** properties for a text */
	public static final String TEXT_PROPERTIES_NAME = "text-properties";
	
	/** name of the color attribute in fo */
	public static final String COLOR_ATTR_NAME = "color";
	/** name of the bold attribute in fo */
	public static final String BOLD_ATTR_NAME = "bold";
	/** name of the bold attribute in fo */
	public static final String FORMULA_ATTR_NAME = "formula";
}
