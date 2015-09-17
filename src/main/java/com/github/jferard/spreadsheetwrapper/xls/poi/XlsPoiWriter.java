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
package com.github.jferard.spreadsheetwrapper.xls.poi;

import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 */
class XlsPoiWriter extends AbstractSpreadsheetWriter implements
SpreadsheetWriter {
	/**
	 * the style for cell dates, since Excel does not know the difference
	 * between a date and a number
	 */
	private final/*@Nullable*/CellStyle dateCellStyle;

	/** the internal reader */
	private final XlsPoiReader preader;

	/**
	 */
	XlsPoiWriter(final Sheet sheet, final CellStyle dateCellStyle) {
		super(new XlsPoiReader(sheet));
		this.preader = (XlsPoiReader) this.reader;
		this.dateCellStyle = dateCellStyle;

	}

	/** {@inheritDoc} */
	@Override
	public void setBoolean(final int r, final int c, final Boolean value) {
		final Cell cell = this.preader.getPOICell(r, c);
		cell.setCellValue(value);
	}

	/** */
	@Override
	public void setDate(final int r, final int c, final Date date) {
		final Cell cell = this.preader.getPOICell(r, c);
		cell.setCellValue(date);
		if (this.dateCellStyle != null)
			cell.setCellStyle(this.dateCellStyle);
	}

	/**
	 */
	@Override
	public void setDouble(final int r, final int c, final Double value) {
		final Cell cell = this.preader.getPOICell(r, c);
		cell.setCellValue(value.doubleValue());
	}

	/** */
	@Override
	public void setFormula(final int r, final int c, final String formula) {
		final Cell cell = this.preader.getPOICell(r, c);
		cell.setCellFormula(formula);
	}

	/** */
	@Override
	public void setInteger(final int r, final int c, final Integer value) {
		final Cell cell = this.preader.getPOICell(r, c);
		cell.setCellValue(value.doubleValue());
	}

	/** */
	@Override
	public void setText(final int r, final int c, final String text) {
		final Cell cell = this.preader.getPOICell(r, c);
		cell.setCellValue(text);
	}
}