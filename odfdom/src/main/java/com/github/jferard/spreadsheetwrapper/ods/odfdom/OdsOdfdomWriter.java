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
package com.github.jferard.spreadsheetwrapper.ods.odfdom;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;

import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.doc.table.OdfTableCell;
import org.odftoolkit.odfdom.doc.table.OdfTableRow;
import org.odftoolkit.odfdom.dom.OdfDocumentNamespace;
import org.odftoolkit.odfdom.dom.element.table.TableTableCellElementBase;
import org.odftoolkit.odfdom.dom.element.table.TableTableElement;
import org.odftoolkit.odfdom.dom.element.table.TableTableRowElement;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;
import com.github.jferard.spreadsheetwrapper.ods.apache.OdsApacheUtil;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 */
class OdsOdfdomWriter extends AbstractSpreadsheetWriter implements SpreadsheetWriter {
	/** format for dates */
	private static int getRowCellCount(final OdfTableRow row) {
		final TableTableRowElement rowElement = row.getOdfElement();
		return getRowElementCellCount(rowElement);
	}

	private static int getRowElementCellCount(
			final TableTableRowElement rowElement) {
		int cellCount = 0;
		int pendingCellCount = 0;

		final NodeList cellList = rowElement
				.getElementsByTagName("table:table-cell");
		final int length = cellList.getLength();
		for (int c = 0; c < length; c++) {
			final TableTableCellElementBase cellElement = (TableTableCellElementBase) cellList
					.item(c);
			final Integer repeat = cellElement
					.getTableNumberColumnsRepeatedAttribute();
			pendingCellCount += repeat;
			if (cellElement.getFirstChild() != null) {
				cellCount += pendingCellCount;
				pendingCellCount = 0;
			}
		}
		return cellCount;
	}

	/** index of current row, -1 if none */
	private int curR;

	/** current row, null if none */
	private/*@Nullable*/OdfTableRow curRow;

	/** helper object for style */
	private final OdsOdfdomStyleHelper styleHelper;

	/** *internal* table */
	private final OdfTable table;

	/**
	 * @param table
	 *            the *internal* sheet
	 */
	OdsOdfdomWriter(final OdfTable table, final OdsOdfdomStyleHelper styleHelper) {
		super();
		this.table = table;
		this.styleHelper = styleHelper;
		this.curR = -1;
		this.curRow = null;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Boolean getBoolean(final int r, final int c) {
		final OdfTableCell cell = this.getOdfCell(r, c);
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
		final OdfTableCell cell = this.getOdfCell(rowIndex, colIndex);
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
		else if (type.equals(OdsConstants.BOOLEAN_TYPE))
			result = cell.getBooleanValue();
		else if (type.equals(OdsConstants.DATE_TYPE)
				|| type.equals(OdsConstants.TIME_TYPE))
			result = getDate(cell);
		else if (type.equals(OdsConstants.FLOAT_TYPE)
				|| type.equals(OdsConstants.CURRENCY_TYPE)
				|| type.equals(OdsConstants.PERCENTAGE_TYPE)) {
			final double value = cell.getDoubleValue();
			if (value == Math.rint(value))
				result = Integer.valueOf((int) value);
			else
				result = Double.valueOf(value);
		} else if (type.equals(OdsConstants.STRING_TYPE))
			result = cell.getStringValue();
		else
			throw new IllegalArgumentException(String.format(
					"Unknown type of cell %s", type));
		return result;
	}

	/** {@inheritDoc} */
	@Override
	public int getCellCount(final int r) {
		if (r < 0 || r >= this.getRowCount())
			throw new IllegalArgumentException();

		final OdfTableRow row = this.table.getRowByIndex(r);
		return getRowCellCount(row);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Date getDate(final int r, final int c) {
		final OdfTableCell cell = this.getOdfCell(r, c);
		if (cell == null)
			return null;

		if (!"date".equals(cell.getValueType())
				&& !"time".equals(cell.getValueType()))
			throw new IllegalArgumentException();
		final Date date = getDate(cell);
		if (date == null)
			throw new IllegalArgumentException();
		return date;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Double getDouble(final int r, final int c) {
		final OdfTableCell cell = this.getOdfCell(r, c);
		if (cell == null)
			return null;

		if (!OdsConstants.FLOAT_TYPE.equals(cell.getValueType()))
			throw new IllegalArgumentException();
		return cell.getDoubleValue();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getFormula(final int r, final int c) {
		final OdfTableCell cell = this.getOdfCell(r, c);
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
		int calculatedRowCount = 0;
		int pendingRowCount = 0;

		final TableTableElement tte = this.table.getOdfElement();
		final NodeList rowList = tte.getElementsByTagName("table:table-row");
		final int length = rowList.getLength();
		for (int r = 0; r < length; r++) {
			final TableTableRowElement rowElement = (TableTableRowElement) rowList
					.item(r);
			final Integer repeat = rowElement
					.getTableNumberRowsRepeatedAttribute();
			pendingRowCount += repeat;
			// not the end ?
			if (getRowElementCellCount(rowElement) != 0) {
				calculatedRowCount += pendingRowCount;
				pendingRowCount = 0;
			}
		}
		return calculatedRowCount;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/WrapperCellStyle getStyle(final int r, final int c) {
		final OdfTableCell odfCell = this.getOdfCell(r, c);
		return this.styleHelper.getStyle(odfCell);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getStyleName(final int r, final int c) {
		final OdfTableCell odfCell = this.getOdfCell(r, c);
		if (odfCell == null)
			return null;

		return odfCell.getStyleName();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getText(final int r, final int c) {
		final OdfTableCell cell = this.getOdfCell(r, c);
		if (cell == null)
			return null;

		if (!OdsConstants.STRING_TYPE.equals(cell.getValueType()))
			throw new IllegalArgumentException();
		return cell.getStringValue();
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
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
		cell.setBooleanValue(value);
		return value;
	}

	/** {@inheritDoc} */
	@Override
	public Date setDate(final int r, final int c, final Date date) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
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

	/** {@inheritDoc} */
	@Override
	public Double setDouble(final int r, final int c, final Number value) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
		final double retValue = value.doubleValue();
		cell.setDoubleValue(retValue);
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public String setFormula(final int r, final int c, final String formula) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
		cell.setFormula("=" + formula);
		cell.setStringValue("");
		return formula;
	}

	/** {@inheritDoc} */
	@Override
	public Integer setInteger(final int r, final int c, final Number value) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
		final int retValue = value.intValue();
		cell.setDoubleValue(Double.valueOf(retValue));
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyle(final int r, final int c,
			final WrapperCellStyle wrapperCellStyle) {
		final OdfTableCell odfCell = this.getOrCreateOdfCell(r, c);
		return this.styleHelper.setWrapperCellStyle(odfCell, wrapperCellStyle);
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c, final String styleName) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
		cell.getOdfElement().setStyleName(styleName);
		return true;
	}

	/** {@inheritDoc} */
	@Override
	public String setText(final int r, final int c, final String text) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
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
	private/*@Nullable*/OdfTableCell getOdfCell(final int r, final int c) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();
		if (r >= this.getRowCount() || c >= this.getCellCount(r))
			return null;
		if (r != this.curR || this.curRow == null) {
			this.curRow = this.table.getRowByIndex(r);
			this.curR = r;
		}
		return this.curRow.getCellByIndex(c);
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
	private OdfTableCell getOrCreateOdfCell(final int r, final int c) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();
		if (r != this.curR || this.curRow == null) {
			// Hack for clean style @see Table.appendRows(count, false)
			// 1. append "manually" the missing rows
			// 2. clean style
			final int lastIndex = this.table.getRowCount() - 1;
			if (r > lastIndex) {
				final List<OdfTableRow> resultList = this.table.appendRows(r
						- lastIndex);
				final String tableNameSpace = OdfDocumentNamespace.TABLE
						.getUri();
				for (final OdfTableRow row : resultList) {
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
		return this.curRow.getCellByIndex(c);
	}

	public static/*@Nullable*/Date getDate(final OdfTableCell cell) {
		cell.getDateValue(); // HACK : throws IllegalArgumentException
		final TableTableCellElementBase odfElement = cell.getOdfElement();
		final String dateStr = odfElement.getOfficeDateValueAttribute();
		if (dateStr == null) {
			return null;
		}
		return OdsApacheUtil.parseString(dateStr);
	}
}