package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.util.Map;

import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WriteException;

import com.github.jferard.spreadsheetwrapper.impl.StyleUtility;

public class XlsJxlStyleUtility extends StyleUtility {
	public WritableCellFormat getCellFormat(final String styleString)
			throws WriteException {
		final Map<String, String> props = this.getPropertiesMap(styleString);
		final WritableFont cellFont = new WritableFont(WritableFont.ARIAL);
		final WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
		for (final Map.Entry<String, String> entry : props.entrySet()) {
			if (entry.getKey().equals("font-weight")) {
				if (entry.getValue().equals("bold"))
					cellFont.setBoldStyle(WritableFont.BOLD);
			} else if (entry.getKey().equals("background-color")) {
				// do nothing
			}
		}
		return cellFormat;
	}
}
