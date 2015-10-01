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
package com.github.jferard.spreadsheetwrapper.ods.jopendocument;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;

import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;

/**
 */
class OdsJOpenWriter extends AbstractSpreadsheetWriter implements
		SpreadsheetWriter {
	/** reader for delegation */
	private final OdsJOpenReader preader;
	private Sheet sheet;

	/**
	 * @param sheet
	 *            the *internal* sheet
	 */
	OdsJOpenWriter(final Sheet sheet) {
		super(new OdsJOpenReader(sheet));
		this.sheet = sheet;
		this.preader = (OdsJOpenReader) this.reader;

	}

	/** {@inheritDoc} */
	@Override
	public Boolean setBoolean(final int r, final int c, final Boolean value) {
		final MutableCell<SpreadSheet> cell = this.preader.getCell(r, c);
		cell.setValue(value);
		return value;
	}

	/**
	 */
	@Override
	public Date setDate(final int r, final int c, final Date date) {
		final MutableCell<SpreadSheet> cell = this.preader.getCell(r, c);
		cell.setValue(date);
		return date;
	}

	/**
	 * @return 
	 */
	@Override
	public Double setDouble(final int r, final int c, final Number value) {
		final MutableCell<SpreadSheet> cell = this.preader.getCell(r, c);
		final double retValue = value.doubleValue();
		cell.setValue(retValue);
		return retValue;
	}

	/**  */
	@Override
	public String setFormula(final int r, final int c, final String formula) {
		final MutableCell<SpreadSheet> cell = this.preader.getCell(r, c);
        cell.getElement().setAttribute("formula", "=" + formula, cell.getElement().getNamespace());
		return formula;
	}

	/**
	 * @return   */
	@Override
	public Integer setInteger(final int r, final int c, final Number value) {
		final MutableCell<SpreadSheet> cell = this.preader.getCell(r, c);
		final int retValue = value.intValue();
		cell.setValue(Double.valueOf(retValue));
		return retValue;
	}

	/**  */
	@Override
	public String setText(final int r, final int c, final String text) {
		final MutableCell<SpreadSheet> cell = this.preader.getCell(r, c);
		cell.setValue(text);
		return text;
	}

	@Override
	public void insertCol(int c) {
		throw new UnsupportedOperationException();
	}

	@Override
	public void insertRow(int r) {
		this.sheet.insertDuplicatedRows(r, 1);
	}

	@Override
	public List<Object> removeCol(int c) {
		this.sheet.removeColumn(c, true);
		return null;
	}

	@Override
	public List<Object> removeRow(int r) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleName(int r, int c) {
		final MutableCell<SpreadSheet> cell = this.preader.getCell(r, c);
		return cell.getStyleName();
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(int r, int c, String styleName) {
		final MutableCell<SpreadSheet> cell = this.preader.getCell(r, c);
		cell.setStyleName(styleName);
		return true;
	}
}