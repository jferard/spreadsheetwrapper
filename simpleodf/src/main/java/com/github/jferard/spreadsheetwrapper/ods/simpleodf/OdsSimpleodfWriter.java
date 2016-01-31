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
package com.github.jferard.spreadsheetwrapper.ods.simpleodf;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;

import org.odftoolkit.odfdom.dom.OdfDocumentNamespace;
import org.odftoolkit.odfdom.dom.element.table.TableTableCellElementBase;
import org.odftoolkit.simple.table.Cell;
import org.odftoolkit.simple.table.Row;
import org.odftoolkit.simple.table.Table;
import org.w3c.dom.Node;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;
import com.github.jferard.spreadsheetwrapper.ods.apache.OdsApacheUtil;
import com.github.jferard.spreadsheetwrapper.ods.apache.OdsOdfdomStyleHelper;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * A sheet writer for simple odf sheet.
 */
class OdsSimpleodfWriter extends AbstractSpreadsheetWriter implements
		SpreadsheetWriter {
	/** format string for integers (internal : double) */
	private static final String INT_FORMAT_STR = "#";

	/** index of the current row, -1 if none */
	private int curR;

	/** current row, null if none */
	private/*@Nullable*/Row curRow;

	/** the style helper */
	private final OdsOdfdomStyleHelper styleHelper;

	/** internal table */
	private final Table table;

	/**
	 * @param table
	 *            the *internal* table
	 * @param delegateStyleHelper
	 */
	OdsSimpleodfWriter(final Table table, final OdsOdfdomStyleHelper styleHelper) {
		this.styleHelper = styleHelper;
		this.table = table;
	}

	/** {@inheritDoc} */
	@Override
	public void insertCol(final int c) {
		this.table.insertColumnsBefore(c, 1);
	}

	/** {@inheritDoc} */
	@Override
	public void insertRow(final int r) {
		this.table.insertRowsBefore(r, 1);
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> removeCol(final int c) {
		final List</*@Nullable*/Object> retValue = this.getColContents(c);
		this.table.removeColumnsByIndex(c, 1);
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> removeRow(final int r) {
		final List</*@Nullable*/Object> retValue = this.getRowContents(r);
		this.table.removeRowsByIndex(r, 1);
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public Boolean setBoolean(final int r, final int c, final Boolean value) {
		final Cell cell = this.getOrCreateSimpleCell(r, c);
		cell.setBooleanValue(value);
		return value;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Date setDate(final int r, final int c, final Date date) {
		final Cell cell = this.getOrCreateSimpleCell(r, c);
		final Calendar cal = Calendar.getInstance();
		cal.setTime(date);
		cell.setDateValue(cal); // for the hidden manipulations.
		final TableTableCellElementBase odfElement = cell.getOdfElement();
		final SimpleDateFormat simpleFormat = new SimpleDateFormat(
				"yyyy-MM-dd'T'HH:mm:ss", Locale.US);
		final String svalue = simpleFormat.format(date);
		odfElement.setOfficeDateValueAttribute(svalue);
		final Date retDate = new Date();
		retDate.setTime(date.getTime() / 1000 * 1000);
		return retDate;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Double setDouble(final int r, final int c, final Number value) {
		final Cell cell = this.getOrCreateSimpleCell(r, c);
		final double retValue = value.doubleValue();
		cell.setDoubleValue(retValue);
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public String setFormula(final int r, final int c, final String formula) {
		final Cell cell = this.getOrCreateSimpleCell(r, c);
		cell.setFormula("=" + formula);
		cell.setStringValue("");
		return formula;
	}

	/** {@inheritDoc} */
	@Override
	public Integer setInteger(final int r, final int c, final Number value) {
		final Cell cell = this.getOrCreateSimpleCell(r, c);
		final int retValue = value.intValue();
		cell.setDoubleValue(Double.valueOf(retValue));
		cell.setFormatString(OdsSimpleodfWriter.INT_FORMAT_STR);
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyle(final int r, final int c,
			final WrapperCellStyle wrapperCellStyle) {
		final Cell cell = this.getOrCreateSimpleCell(r, c);
		if (cell == null)
			return false;

		final TableTableCellElementBase odfElement = cell.getOdfElement();
		this.styleHelper.setWrapperCellStyle(odfElement, wrapperCellStyle);
		return true;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c, final String styleName) {
		final Cell cell = this.getOrCreateSimpleCell(r, c);
		cell.setCellStyleName(styleName);
		return true;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public String setText(final int r, final int c, final String text) {
		final Cell cell = this.getOrCreateSimpleCell(r, c);
		cell.setStringValue(text);
		return text;
	}

	/**
	 * Simple optimization hidden inside a method.
	 *
	 * @param r
	 *            the row index
	 * @param c
	 *            the column index
	 * @return the cell
	 */
	private Cell getOrCreateSimpleCell(final int r, final int c) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();
		if (r != this.curR || this.curRow == null) {
			// Hack for clean style @see Table.appendRows(count, false)
			// 1. append "manually" the missing rows
			// 2. clean style
			final int lastIndex = this.table.getRowCount() - 1;
			if (r > lastIndex) {
				final List<Row> resultList = this.table.appendRows(r
						- lastIndex);
				final String tableNameSpace = OdfDocumentNamespace.TABLE
						.getUri();
				for (final Row row : resultList) {
					Node cellE = row.getOdfElement().getFirstChild();
					while (cellE != null) {
						((TableTableCellElementBase) cellE).removeAttributeNS(
								tableNameSpace, "style-name");
						cellE = cellE.getNextSibling();
					}
				}
			}
			this.curRow = this.table.getRowByIndex(r);
			this.curR = r;
		}
		final Cell cell = this.curRow.getCellByIndex(c);
		cell.getStyleHandler().getStyleElementForWrite();
		return cell;
	}
	
	private static/*@Nullable*/Date getDate(final Cell cell) {
		cell.getDateValue(); // HACK : throws IllegalArgumentException
		final TableTableCellElementBase odfElement = cell.getOdfElement();
		final String dateStr = odfElement.getOfficeDateValueAttribute();
		if (dateStr == null) {
			return null;
		}
		return OdsApacheUtil.parseString(dateStr);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Boolean getBoolean(final int r, final int c) {
		final Cell cell = this.getSimpleCell(r, c);
		if (cell == null)
			return null;
		if (!OdsConstants.BOOLEAN_TYPE.equals(cell.getValueType()))
			throw new IllegalArgumentException();
		return cell.getBooleanValue();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Object getCellContent(final int rowIndex,
			final int colIndex) {
		final Cell cell = this.getSimpleCell(rowIndex, colIndex);
		if (cell == null)
			return null;
		final String formula = cell.getFormula();
		if (formula != null && formula.charAt(0) == '=')
			return formula.substring(1);

		final String type = cell.getValueType();
		Object result;

		// from the doc
		// The type can be "boolean", "currency", "date", "float", "percentage",
		// "string" or "time".
		if (type == null)
			result = null;
		else if (type.equals("boolean"))
			result = cell.getBooleanValue();
		else if (type.equals("date") || type.equals("time"))
			result = OdsSimpleodfWriter.getDate(cell);
		else if (type.equals("float") || type.equals("currency")
				|| type.equals("percentage")) {
			final double value = cell.getDoubleValue();
			if (value == Math.rint(value))
				result = Integer.valueOf((int) value);
			else
				result = Double.valueOf(value);
		} else if (type.equals("string"))
			result = cell.getStringValue();
		else
			throw new IllegalArgumentException(String.format(
					"Unknown type of cell %s", type));
		return result;
	}

	/** {@inheritDoc} */
	@Override
	public int getCellCount(final int r) {
		if (r < 0 || r >= this.table.getRowCount())
			throw new IllegalArgumentException();

		final Row row = this.table.getRowByIndex(r);
		// return row.getCellCount();

		for (int i = this.table.getColumnCount() - 1; i >= 0; i--) {
			final Cell cell = row.getCellByIndex(i);
			if (cell.getOdfElement().getChildNodes().getLength() != 0)
				return i + 1;
		}
		return 0;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Date getDate(final int r, final int c) {
		final Cell cell = this.getSimpleCell(r, c);
		if (cell == null)
			return null;
		final Date date = OdsSimpleodfWriter.getDate(cell);
		if (date == null)
			throw new IllegalArgumentException();
		return date;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Double getDouble(final int r, final int c) {
		final Cell cell = this.getSimpleCell(r, c);
		if (cell == null)
			return null;
		return cell.getDoubleValue();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getFormula(final int r, final int c) {
		final Cell cell = this.getSimpleCell(r, c);
		if (cell == null)
			return null;
		final String formula = cell.getFormula();
		if (formula == null || formula.charAt(0) != '=')
			throw new IllegalArgumentException();

		return formula.substring(1);
	}

	/** {@inheritDoc} */
	@Override
	public String getName() {
		return this.table.getTableName();
	}

	/** {@inheritDoc} */
	@Override
	public int getRowCount() {
		int rowCount = this.table.getRowCount();
		if (rowCount == 1 && this.getCellCount(0) == 0)
			rowCount = 0;

		return rowCount;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/WrapperCellStyle getStyle(final int r, final int c) {
		final Cell cell = this.getSimpleCell(r, c);
		if (cell == null)
			return null;

		final TableTableCellElementBase odfElement = cell.getOdfElement();
		return this.styleHelper.getWrapperCellStyle(odfElement);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getStyleName(final int r, final int c) {
		final Cell cell = this.getSimpleCell(r, c);
		if (cell == null)
			return null;

		return cell.getStyleName();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getText(final int r, final int c) {
		final Cell cell = this.getSimpleCell(r, c);
		if (cell == null)
			return null;
		if (!OdsConstants.STRING_TYPE.equals(cell.getValueType()))
			throw new IllegalArgumentException();
		return cell.getStringValue();
	}

	/**
	 * Simple optimization hidden inside a method.
	 *
	 * @param r
	 *            the row index
	 * @param c
	 *            the column index
	 * @return the cell
	 */
	protected/*@Nullable*/Cell getSimpleCell(final int r, final int c) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();
		if (r >= this.getRowCount() || c >= this.getCellCount(r))
			return null;

		if (r != this.curR || this.curRow == null) {
			this.curRow = this.table.getRowByIndex(r);
			this.curR = r;
		}
		final Cell cell = this.curRow.getCellByIndex(c);
		cell.getStyleHandler().getStyleElementForWrite();
		return cell;
	}	
}