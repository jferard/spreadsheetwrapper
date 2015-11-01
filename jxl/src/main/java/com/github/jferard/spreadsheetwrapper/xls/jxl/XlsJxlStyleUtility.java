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
package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.util.HashMap;
import java.util.Map;

import jxl.format.Colour;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WriteException;

import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.impl.StyleUtility;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

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
				if (jxlColor != null)
					this.jxlColorByColor.put(color, jxlColor);
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
		}
	}

	@Deprecated
	public WritableCellFormat getCellFormat(final String styleString)
			throws WriteException {
		final Map<String, String> props = this.getPropertiesMap(styleString);
		final WritableFont cellFont = new WritableFont(WritableFont.ARIAL);
		final WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
		for (final Map.Entry<String, String> entry : props.entrySet()) {
			if (entry.getKey().equals(StyleUtility.FONT_WEIGHT)) {
				if (entry.getValue().equals("bold"))
					cellFont.setBoldStyle(WritableFont.BOLD);
			} else if (entry.getKey().equals(StyleUtility.BACKGROUND_COLOR)) {
				// do nothing
			}
		}
		return cellFormat;
	}

	public/*@Nullable*/Colour getJxlColor(final WrapperColor backgroundColor) {
		return this.jxlColorByColor.get(backgroundColor);
	}
}
