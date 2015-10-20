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
package com.github.jferard.spreadsheetwrapper.ods.jopendocument12;

import java.math.BigDecimal;
import java.util.Date;

import org.jdom.Element;
import org.jopendocument.dom.ODValueType;
import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetReader;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable; */

/**
 */
class OdsJOpenReader extends AbstractSpreadsheetReader implements
SpreadsheetReader {
	/** the *internal* table */
	private final Sheet sheet;

	/**
	 * @param sheet
	 *            the *internal* sheet
	 */
	OdsJOpenReader(final Sheet sheet) {
		super();
		this.sheet = sheet;
	}

	/** {@inheritDoc} */
	@Override
	public Boolean getBoolean(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		final String type = this.getTypeName(cell);
		// 1.3b1
		// if (!"boolean".equals(type))
		// throw new IllegalArgumentException();
		// return (Boolean) cell.getValue();
		// 1.2
		if (!"float".equals(type))
			throw new IllegalArgumentException();
		return ((BigDecimal) cell.getValue()).doubleValue() != 0.0;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Object getCellContent(final int rowIndex,
			final int colIndex) {
		final MutableCell<SpreadSheet> cell = this.getCell(rowIndex, colIndex);
		if (cell == null)
			return null;
		final String formula = this.getFormula(cell);
		if (formula != null && formula.charAt(0) == '=')
			return formula.substring(1);

		Object value = cell.getValue();
		if (value instanceof BigDecimal) {
			final BigDecimal decimal = (BigDecimal) value;
			try {
				value = Integer.valueOf(decimal.intValueExact());
			} catch (final ArithmeticException e) {
				value = Double.valueOf(decimal.doubleValue());
			}
		}
		return value;
	}

	/** {@inheritDoc} */
	@Override
	public int getCellCount(final int r) {
		return this.sheet.getColumnCount();
	}

	/** {@inheritDoc} */
	@Override
	public Date getDate(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		final String type = this.getTypeName(cell);
		if (!"date".equals(type) && !"time".equals(type))
			throw new IllegalArgumentException();
		return (Date) cell.getValue();
	}

	/** {@inheritDoc} */
	@Override
	public Double getDouble(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		final String type = this.getTypeName(cell);
		if (!"float".equals(type))
			throw new IllegalArgumentException();
		final Object value = cell.getValue();
		if (!(value instanceof BigDecimal))
			throw new IllegalStateException();
		final BigDecimal decimal = (BigDecimal) value;
		return decimal.doubleValue();
	}

	/** {@inheritDoc} */
	@Override
	public String getFormula(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		final String formula = this.getFormula(cell);
		if (formula == null || formula.charAt(0) != '=')
			throw new IllegalArgumentException();

		return formula.substring(1);
	}

	/** {@inheritDoc} */
	@Override
	public String getName() {
		return this.sheet.getName();
	}

	/** {@inheritDoc} */
	@Override
	public int getRowCount() {
		return this.sheet.getRowCount();
	}

	/** {@inheritDoc} */
	@Override
	public String getText(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		final String type = this.getTypeName(cell);
		if (!"string".equals(type))
			throw new IllegalArgumentException();
		return (String) cell.getValue();
	}

	private String getFormula(final MutableCell<SpreadSheet> cell) {
		// 1.3b1
		// return cell.getFormula();
		// 1.2
		final Element element = cell.getElement();
		return element.getAttributeValue("formula", element.getNamespace());
	}

	private String getTypeName(final MutableCell<SpreadSheet> cell) {
		final ODValueType valueType = cell.getValueType();
		return valueType == null ? "string" : valueType.getName();
	}

	/**
	 * Simple optimization hidden inside a method.
	 *
	 * @param r
	 *            the row index
	 * @param c
	 *            the column index
	 * @return
	 * @return the cell
	 */
	protected MutableCell<SpreadSheet> getCell(final int r, final int c) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();

		if (r >= this.sheet.getRowCount()) {
			// do not set more because this will affect the row count value
			this.sheet.ensureRowCount(r + 1);
		}
		if (c >= this.sheet.getColumnCount()) {
			// do not set more because this will affect the row count value
			this.sheet.ensureColumnCount(c + 1);
		}
		return this.sheet.getCellAt(c, r);
	}
}