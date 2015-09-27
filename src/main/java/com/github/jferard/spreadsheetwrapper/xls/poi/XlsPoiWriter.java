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
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
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

	private Sheet sheet;

	/**
	 */
	XlsPoiWriter(final Sheet sheet, final /*@Nullable*/ CellStyle dateCellStyle) {
		super(new XlsPoiReader(sheet));
		this.sheet = sheet;
		this.preader = (XlsPoiReader) this.reader;
		this.dateCellStyle = dateCellStyle;

	}

	/** {@inheritDoc} 
	 * @return */
	@Override
	public Boolean setBoolean(final int r, final int c, final Boolean value) {
		final Cell cell = this.preader.getPOICell(r, c);
		cell.setCellValue(value);
		return value;
	}

	/**
	 * @return  */
	@Override
	public Date setDate(final int r, final int c, final Date date) {
		final Cell cell = this.preader.getPOICell(r, c);
		cell.setCellValue(date);
		if (this.dateCellStyle != null)
			cell.setCellStyle(this.dateCellStyle);
		return date;
	}

	/**
	 * @return 
	 */
	@Override
	public Double setDouble(final int r, final int c, final Number value) {
		final Cell cell = this.preader.getPOICell(r, c);
		final double retValue = value.doubleValue();
		cell.setCellValue(retValue);
		return retValue;
	}

	/**
	 * @return  */
	@Override
	public String setFormula(final int r, final int c, final String formula) {
		final Cell cell = this.preader.getPOICell(r, c);
		cell.setCellFormula(formula);
		return formula;
	}

	/**
	 * @return  */
	@Override
	public Integer setInteger(final int r, final int c, final Number value) {
		final Cell cell = this.preader.getPOICell(r, c);
		int retValue = value.intValue();
		cell.setCellValue(retValue);
		return retValue;
	}

	/** */
	@Override
	public String setText(final int r, final int c, final String text) {
		final Cell cell = this.preader.getPOICell(r, c);
		cell.setCellValue(text);
		return text;
	}
	
	/** {@inheritDoc} */
	@Override
	public void insertCol(int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public void insertRow(int r) {
		this.sheet.shiftRows(r,  this.getRowCount(), 1);
		this.sheet.createRow(r);
	}

	/** {@inheritDoc} */
	@Override
	public List<Object> removeCol(int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public List<Object> removeRow(int r) {
		List<Object> ret = this.getRowContents(r);
		Row row = this.sheet.getRow(r);
		this.sheet.removeRow(row);
		this.sheet.shiftRows(r,  this.getRowCount(), 1);
		return ret;
	}
	
	/** {@inheritDoc} */
	@Override
	public String getStyleName(int r, int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleString(int r, int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(int r, int c, String styleName) {
		throw new UnsupportedOperationException();
	}
	
}