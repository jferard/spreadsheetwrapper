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
package com.github.jferard.spreadsheetwrapper;

import java.util.Date;

import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public interface SpreadsheetWriterCursor extends SpreadsheetReaderCursor {
	/**
	 * {@link SpreadsheetWriter.setCellContent(int, int, Object)}
	 */
	/*@Nullable*/ Object setCellContent(Object content);

	/**
	 * {@link SpreadsheetWriter.setCellContent(int, int, Object, String)}
	 */
	/*@Nullable*/ Object setCellContent(Object content, String styleName);

	/**
	 * {@link SpreadsheetWriter.setDate(int, int, Date)}
	 */
	/*@Nullable*/ Date setDate(Date date);

	/**
	 * {@link SpreadsheetWriter.setDate(int, int, Date, String)}
	 */
	/*@Nullable*/Date setDate(Date date, String styleName);

	/**
	 * {@link SpreadsheetWriter.setDouble(int, int, Double)}
	 */
	/*@Nullable*/Double setDouble(Number value);

	/**
	 * {@link SpreadsheetWriter.setDouble(int, int, Number, String)}
	 */
	/*@Nullable*/Double setDouble(Number value, String styleName);

	/**
	 * {@link SpreadsheetWriter.setFormula(int, int, String)}
	 */
	/*@Nullable*/String setFormula(String formula);

	/**
	 * {@link SpreadsheetWriter.setFormula(int, int, String, String)}
	 */
	/*@Nullable*/String setFormula(String formula, String styleName);

	/**
	 * {@link SpreadsheetWriter.setInteger(int, int, Number)}
	 */
	/*@Nullable*/Integer setInteger(Number value);

	/**
	 * {@link SpreadsheetWriter.setInteger(int, int, Number, String)}
	 */
	/*@Nullable*/Integer setInteger(Number value, String styleName);

	/**
	 * {@link SpreadsheetWriter.setStyle(int, int, WrapperCellStyle)}
	 */
	boolean setStyle(WrapperCellStyle wrapperStyle);

	/**
	 * {@link SpreadsheetWriter.setStyleName(int, int, String)}
	 */
	boolean setStyleName(String styleName);

	/**
	 * {@link SpreadsheetWriter.setText(int, int, String)}
	 */
	/*@Nullable*/String setText(String text);

	/**
	 * {@link SpreadsheetWriter.setText(int, int, String, String)}
	 */
	/*@Nullable*/String setText(String text, String styleName);

}