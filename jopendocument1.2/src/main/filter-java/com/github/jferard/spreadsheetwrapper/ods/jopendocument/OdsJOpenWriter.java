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
package com.github.jferard.spreadsheetwrapper.ods.${jopendocument.pkg};

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;

import org.jdom.Namespace;
import org.jopendocument.dom.ODValueType;
import org.jopendocument.dom.OOUtils;
import org.jopendocument.dom.spreadsheet.CellStyle;
import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 */
class OdsJOpenWriter extends AbstractSpreadsheetWriter implements
SpreadsheetWriter {
	/** the *internal* sheet wrapped */
	private final Sheet sheet;

	/** the style helper */
	private final OdsJOpenStyleHelper styleHelper;


	/**
	 * @param sheet
	 *            the *internal* sheet
	 * @param styleHelper the style helper
	 */
	OdsJOpenWriter(final OdsJOpenStyleHelper styleHelper, final Sheet sheet) {
		this.styleHelper = styleHelper;
		this.sheet = sheet;
	}

	/** {@inheritDoc} */
	@Override
	public void insertCol(final int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public void insertRow(final int r) {
		this.sheet.insertDuplicatedRows(r, 1);
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> removeCol(final int c) {
		final List</*@Nullable*/Object> ret = this.getColContents(c);
		this.sheet.removeColumn(c, true);
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> removeRow(final int r) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public Boolean setBoolean(final int r, final int c, final Boolean value) {
		final MutableCell<SpreadSheet> cell = this.getOrCreateCell(r, c);
		${jopendocument.util.cls}.setValue(cell, ODValueType.BOOLEAN, value);
		return value;
	}

	/** {@inheritDoc} */
	@Override
	public Date setDate(final int r, final int c, final Date date) {
		final MutableCell<SpreadSheet> cell = this.getOrCreateCell(r, c);
		cell.setValue(date);
		return date;
	}

	/** {@inheritDoc} */
	@Override
	public Double setDouble(final int r, final int c, final Number value) {
		final MutableCell<SpreadSheet> cell = this.getOrCreateCell(r, c);
		final double retValue = value.doubleValue();
		cell.setValue(retValue);
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public String setFormula(final int r, final int c, final String formula) {
		final MutableCell<SpreadSheet> cell = this.getOrCreateCell(r, c);
		cell.getElement().setAttribute("formula", "=" + formula,
				cell.getElement().getNamespace());
		cell.setValue(0);
		return formula;
	}

	/** {@inheritDoc} */
	@Override
	public Integer setInteger(final int r, final int c, final Number value) {
		final MutableCell<SpreadSheet> cell = this.getOrCreateCell(r, c);
		final int retValue = value.intValue();
		cell.setValue(Double.valueOf(retValue));
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyle(final int r, final int c,
			final WrapperCellStyle wrapperStyle) {
		final MutableCell<SpreadSheet> cell = this.getOrCreateCell(r, c);
		return this.styleHelper.setStyle(cell, wrapperStyle);
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c, final String styleName) {
		final MutableCell<SpreadSheet> cell = this.getOrCreateCell(r, c);
		return this.styleHelper.setStyleName(cell, styleName);
	}

	/** {@inheritDoc} */
	@Override
	public String setText(final int r, final int c, final String text) {
		final MutableCell<SpreadSheet> cell = this.getOrCreateCell(r, c);
		cell.setValue(text);
		return text;
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
	protected MutableCell<SpreadSheet> getOrCreateCell(final int r, final int c) {
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
	
	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Boolean getBoolean(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		if (cell == null)
			return null;

		final String type = this.getTypeName(cell);
		if (!OdsConstants.BOOLEAN_TYPE.equals(type))
			throw new IllegalArgumentException();
		return (Boolean) cell.getValue();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Object getCellContent(final int rowIndex,
			final int colIndex) {
		final MutableCell<SpreadSheet> cell = this.getCell(rowIndex, colIndex);
		if (cell == null)
			return null;
		final String formula = ${jopendocument.util.cls}.getFormula(cell);
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
		if (r >= this.getRowCount())
			throw new IllegalArgumentException();
		
		return this.getCellCountUnsafe(r);
	}
		
	private int getCellCountUnsafe(final int r) {
		final int columnCount = this.sheet.getColumnCount();
		for (int c = columnCount - 1; c >= 0; c--) {
			final MutableCell<SpreadSheet> cell = this.sheet.getCellAt(c, r);
			if (cell != null && cell.getValue() != null
					&& !cell.getTextValue().isEmpty())
				return c + 1;
		}
		return 0;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Date getDate(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		if (cell == null)
			return null;

		final String type = this.getTypeName(cell);
		if (!OdsConstants.DATE_TYPE.equals(type)
				&& !OdsConstants.TIME_TYPE.equals(type))
			throw new IllegalArgumentException();
		return (Date) cell.getValue();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Double getDouble(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		if (cell == null)
			return null;

		final String type = this.getTypeName(cell);
		if (!OdsConstants.FLOAT_TYPE.equals(type))
			throw new IllegalArgumentException();
		final Object value = cell.getValue();
		if (!(value instanceof BigDecimal))
			throw new IllegalStateException();
		final BigDecimal decimal = (BigDecimal) value;
		return decimal.doubleValue();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getFormula(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		if (cell == null)
			return null;

		final String formula = ${jopendocument.util.cls}.getFormula(cell);
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
		int rowCount = this.sheet.getRowCount();
		if (rowCount == 1 && this.getCellCountUnsafe(0) == 0)
			rowCount = 0;
		return rowCount;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/WrapperCellStyle getStyle(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		return this.styleHelper.getStyle(cell);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getStyleName(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		return this.styleHelper.getStyleName(cell);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getText(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		if (cell == null)
			return null;

		final String type = this.getTypeName(cell);
		if (!OdsConstants.STRING_TYPE.equals(type))
			throw new IllegalArgumentException();
		return (String) cell.getValue();
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
	private/*@Nullable*/MutableCell<SpreadSheet> getCell(final int r,
			final int c) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();
		if (r >= this.getRowCount() || c >= this.getCellCount(r))
			return null;

		return this.sheet.getCellAt(c, r);
	}

	private String getTypeName(final MutableCell<SpreadSheet> cell) {
		final ODValueType valueType = cell.getValueType();
		return valueType == null ? OdsConstants.STRING_TYPE : valueType
				.getName();
	}}