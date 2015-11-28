package com.github.jferard.spreadsheetwrapper.xls;

/**
 * Only a few constants for XLS.
 *
 */
public final class XlsConstants {
	/** standard extension for files */
	public static final String EXTENSION1 = "xls";

	/** extended extension for files */
	public static final String EXTENSION2 = "xlsx";

	/**
	 * The maximum number of columns per row
	 */
	public static final int MAX_COLUMNS = 256;

	/**
	 * The maximum number of rows excel allows in a worksheet
	 */
	public static final int MAX_ROWS_PER_SHEET = 65536;

	private XlsConstants() {
	}
}
