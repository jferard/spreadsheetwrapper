package com.github.jferard.spreadsheetwrapper.xls;

/**
 * Only two constants for XLS.
 *
 */
public final class XlsConstants {
	private XlsConstants() {}
	
	/**
	 * The maximum number of columns per row
	 */
	public static final int MAX_COLUMNS = 256;

	/**
	 * The maximum number of rows excel allows in a worksheet
	 */
	public static final int MAX_ROWS_PER_SHEET = 65536;
}
