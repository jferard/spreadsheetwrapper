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

public interface SpreadsheetReaderCursor extends Cursor {

	/**
	 * {@link SpreadsheetReader.getCellContent}
	 */
	/*@Nullable*/Object getCellContent();

	/**
	 * {@link SpreadsheetReader.getDate}
	 */
	/*@Nullable*/Date getDate();

	/**
	 * {@link SpreadsheetReader.getDouble}
	 */
	/*@Nullable*/Double getDouble();

	/**
	 * {@link SpreadsheetReader.getFormula}
	 */
	/*@Nullable*/String getFormula();

	/**
	 * {@link SpreadsheetReader.getInteger}
	 */
	/*@Nullable*/Integer getInteger();

	/**
	 * {@link SpreadsheetReader.getStyle}
	 */
	/*@Nullable*/WrapperCellStyle getStyle();

	/**
	 * {@link SpreadsheetReader.getStyleName}
	 */
	/*@Nullable*/String getStyleName();

	/**
	 * {@link SpreadsheetReader.getText}
	 */
	/*@Nullable*/String getText();
}