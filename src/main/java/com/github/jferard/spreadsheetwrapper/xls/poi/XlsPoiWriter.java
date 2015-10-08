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
import java.util.List;
import java.util.Map;

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
	private final Map<String, CellStyle> cellStyleByName;

	/**
	 * the style for cell dates, since Excel does not know the difference
	 * between a date and a number
	 */
	private final/*@Nullable*/CellStyle dateCellStyle;

	/** the internal reader */
	private final XlsPoiReader preader;

	private final Sheet sheet;

	/**
	 * @param cellStyleByName
	 */
	XlsPoiWriter(final Sheet sheet, final/*@Nullable*/CellStyle dateCellStyle,
			final Map<String, CellStyle> cellStyleByName) {
		super(new XlsPoiReader(sheet));
		this.sheet = sheet;
		this.cellStyleByName = cellStyleByName;
		this.preader = (XlsPoiReader) this.reader;
		this.dateCellStyle = dateCellStyle;

	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getStyleName(final int r, final int c) {
		final Cell cell = this.preader.getPOICell(r, c);
		final CellStyle cellStyle = cell.getCellStyle();
		for (final Map.Entry<String, CellStyle> entry : this.cellStyleByName
				.entrySet()) {
			if (entry.getValue().equals(cellStyle))
				return entry.getKey();
		}
		return null;
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleString(final int r, final int c) {
		final Cell cell = this.preader.getPOICell(r, c);
		final CellStyle cellStyle = cell.getCellStyle();
		return ""; // this.xlsPoiUtil.getStyleString(this.workbook, cellStyle);
	}

	/** {@inheritDoc} */
	@Override
	public void insertCol(final int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public void insertRow(final int r) {
		this.sheet.shiftRows(r, this.getRowCount(), 1);
		this.sheet.createRow(r);
	}

	/** {@inheritDoc} */
	@Override
	public List<Object> removeCol(final int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public List<Object> removeRow(final int r) {
		final List<Object> ret = this.getRowContents(r);
		final Row row = this.sheet.getRow(r);
		this.sheet.removeRow(row);
		this.sheet.shiftRows(r, this.getRowCount(), 1);
		return ret;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Boolean setBoolean(final int r, final int c, final Boolean value) {
		final Cell cell = this.preader.getPOICell(r, c);
		cell.setCellValue(value);
		return value;
	}

	/**
	 * @return
	 */
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
	 * @return
	 */
	@Override
	public String setFormula(final int r, final int c, final String formula) {
		final Cell cell = this.preader.getPOICell(r, c);
		cell.setCellFormula(formula);
		return formula;
	}

	/**
	 * @return
	 */
	@Override
	public Integer setInteger(final int r, final int c, final Number value) {
		final Cell cell = this.preader.getPOICell(r, c);
		final int retValue = value.intValue();
		cell.setCellValue(retValue);
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c, final String styleName) {
		if (this.cellStyleByName.containsKey(styleName)) {
			final Cell cell = this.preader.getPOICell(r, c);
			cell.setCellStyle(this.cellStyleByName.get(styleName));
			return true;
		} else
			return false;
	}

	/** */
	@Override
	public String setText(final int r, final int c, final String text) {
		final Cell cell = this.preader.getPOICell(r, c);
		cell.setCellValue(text);
		return text;
	}

}