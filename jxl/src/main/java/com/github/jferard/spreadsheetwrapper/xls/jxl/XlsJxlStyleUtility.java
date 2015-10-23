package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.util.HashMap;
import java.util.Map;

import jxl.format.Colour;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WriteException;

import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.impl.StyleUtility;

public class XlsJxlStyleUtility extends StyleUtility {
	private final Map<WrapperColor, Colour> jxlColorByColor;

	XlsJxlStyleUtility() {
		final WrapperColor[] colors = WrapperColor.values();
		this.jxlColorByColor = new HashMap<WrapperColor, Colour>(colors.length);
		for (final WrapperColor color : colors) {
			Colour jxlColor;
			try {
				final Class<?> jxlClazz = Class.forName("jxl.format.Colour");
				jxlColor = (Colour) jxlClazz.getDeclaredField(
						color.getSimpleName()).get(null);
			} catch (final ClassNotFoundException e) {
				jxlColor = null;
			} catch (final IllegalArgumentException e) {
				jxlColor = null;
			} catch (final IllegalAccessException e) {
				jxlColor = null;
			} catch (final NoSuchFieldException e) {
				jxlColor = null;
			} catch (final SecurityException e) {
				jxlColor = null;
			}
			this.jxlColorByColor.put(color, jxlColor);
		}
	}

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

	public Colour getJxlColor(final WrapperColor backgroundColor) {
		return this.jxlColorByColor.get(backgroundColor);
	}
}
