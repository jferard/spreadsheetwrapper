package com.github.jferard.spreadsheetwrapper.xls.poi;

import java.util.HashMap;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;

import com.github.jferard.spreadsheetwrapper.style.WrapperColor;

class XlsPoiStyleColorHelper {

	/** internal -> wrapper */
	private final Map<HSSFColor, WrapperColor> colorByHssfColor;
	/** wrapper -> internal */
	private final Map<WrapperColor, HSSFColor> hssfColorByColor;

	// 		final Logger logger = Logger.getLogger(this.getClass().getName());
	XlsPoiStyleColorHelper(Logger logger) {
		final WrapperColor[] colors = WrapperColor.values();
		this.hssfColorByColor = new HashMap<WrapperColor, HSSFColor>(
				colors.length);
		this.colorByHssfColor = new HashMap<HSSFColor, WrapperColor>(
				colors.length);
		final Map<Integer, HSSFColor> hssfColorByIndex = HSSFColor
				.getIndexHash();
		for (final HSSFColor hssfColor : hssfColorByIndex.values()) {
			final String hssfColorName = hssfColor.getClass().getName();
			final int index = hssfColorName.indexOf('$');
			if (index == -1)
				continue;

			final String colorName = hssfColorName.substring(index + 1);
			try {
				final WrapperColor wrapperColor = WrapperColor
						.valueOf(colorName);
				this.hssfColorByColor.put(wrapperColor, hssfColor);
				this.colorByHssfColor.put(hssfColor, wrapperColor);
			} catch (final IllegalArgumentException e) {
				logger.warning(String
						.format("Missing color '%s' in WrapperColor class. Those colors won't be available for POI wrapper.",
								colorName));
			}
		}

	}

	/**
	 * @param wrapperColor
	 *            the color to convert
	 * @return the HSSF color
	 */
	public HSSFColor toHSSFColor(final WrapperColor wrapperColor) {
		HSSFColor hssfColor;
		if (this.hssfColorByColor.containsKey(wrapperColor))
			hssfColor = this.hssfColorByColor.get(wrapperColor);
		else
			hssfColor = new HSSFColor.WHITE();
		return hssfColor;
	}

	public HSSFColor getHSSFColor(final short fontColorIndex) {
		final Map<Integer, HSSFColor> indexHash = HSSFColor.getIndexHash();
		final HSSFColor poiFontColor = indexHash.get(Integer
				.valueOf(fontColorIndex));
		return poiFontColor;
	}
	
	public WrapperColor toWrapperColor(final HSSFColor hssfColor) {
		return this.colorByHssfColor.get(hssfColor);
	}
	

	public WrapperColor toWrapperColor(CellStyle cellStyle) {
		WrapperColor wrapperColor = null;
		final short backgroundColorIndex = cellStyle.getFillForegroundColor();
		final HSSFColor poiBackgroundColor = this
				.getHSSFColor(backgroundColorIndex);
		if (cellStyle.getFillPattern() == CellStyle.SOLID_FOREGROUND)
			wrapperColor = this.toWrapperColor(poiBackgroundColor);

		if (WrapperColor.DEFAULT_BACKGROUND.equals(wrapperColor))
			wrapperColor = null;

		return wrapperColor;
	}

}
