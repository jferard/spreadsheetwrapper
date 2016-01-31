package com.github.jferard.spreadsheetwrapper.ods.apache;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import java.util.logging.Level;
import java.util.logging.Logger;

final public class OdsApacheUtil {
	private static final String DATE_FORMAT = "yyyy-MM-dd'T'HH:mm:ss";

	/**
	 * @param dateString
	 *            the date as a String (e.g "2015-01-01")
	 * @return the date, null if can't parse
	 */
	public static/*@Nullable*/Date parseString(final String dateString) {
		return OdsApacheUtil.parseString(dateString, DATE_FORMAT);
	}
	
	/**
	 * @param dateString
	 *            the date as a String (e.g "2015-01-01")
	 * @param format
	 *            the format
	 * @return the date, null if can't parse
	 */
	public static/*@Nullable*/Date parseString(final String dateString,
			final String format) {
		final SimpleDateFormat simpleFormat = new SimpleDateFormat(format,
				Locale.US);
		Date simpleDate;
		try {
			simpleDate = simpleFormat.parse(dateString);
		} catch (final ParseException e) {
			String message = e.getMessage();
			if (message == null)
				message = "???";
			Logger.getLogger(OdsApacheUtil.class.getName())
					.log(Level.SEVERE, message, e);
			simpleDate = null;
		}
		return simpleDate;
	}

	private OdsApacheUtil() {
	}
}
