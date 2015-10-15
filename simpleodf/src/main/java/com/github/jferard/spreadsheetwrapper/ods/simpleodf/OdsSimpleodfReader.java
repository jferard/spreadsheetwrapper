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

import java.util.Date;
import java.util.List;

import org.odftoolkit.odfdom.dom.OdfDocumentNamespace;
import org.odftoolkit.odfdom.dom.element.table.TableTableCellElementBase;
import org.odftoolkit.simple.table.Cell;
import org.odftoolkit.simple.table.Row;
import org.odftoolkit.simple.table.Table;
import org.w3c.dom.Node;

import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetReader;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/
/**
 * the sheet reader for SimpleODS.
 */
class OdsSimpleodfReader extends AbstractSpreadsheetReader implements
SpreadsheetReader {
	private static/*@Nullable*/Date getDate(final Cell cell) {
		cell.getDateValue(); // HACK : throws IllegalArgumentException
		final TableTableCellElementBase odfElement = cell.getOdfElement();
		final String dateStr = odfElement.getOfficeDateValueAttribute();
		if (dateStr == null) {
			return null;
		}
		final Date date = AbstractSpreadsheetReader.parseString(dateStr,
				"yyyy-MM-dd'T'HH:mm:ss");
		return date;
	}

	/** index of the current row, -1 if none */
	private int curR;

	/** current row, null if none */
	private/*@Nullable*/Row curRow;

	/** the *internal* table */
	private final Table table;

	/**
	 * Creates a reader from an *internal* table
	 *
	 * @param table
	 */
	OdsSimpleodfReader(final Table table) {
		super();
		this.table = table;
		this.curR = -1;
		this.curRow = null;
	}

	/** {@inheritDoc} */
	@Override
	public Boolean getBoolean(final int r, final int c) {
		final Cell cell = this.getSimpleCell(r, c);
		if (!"boolean".equals(cell.getValueType()))
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
		//  The type can be "boolean", "currency", "date", "float", "percentage", "string" or "time". 
		if (type == null)
			result = null;
		else if (type.equals("boolean"))
			result = cell.getBooleanValue();
		else if (type.equals("date") || type.equals("time"))
			result = OdsSimpleodfReader.getDate(cell);
		else if (type.equals("float") || type.equals("currency") || type.equals("percentage")) {
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
		final Row row = this.table.getRowByIndex(r);
		return row.getCellCount();
	}

	/** {@inheritDoc} */
	@Override
	public Date getDate(final int r, final int c) {
		final Cell cell = this.getSimpleCell(r, c);
		final Date date = OdsSimpleodfReader.getDate(cell);
		if (date == null)
			throw new IllegalArgumentException();
		return date;
	}

	/** {@inheritDoc} */
	@Override
	public Double getDouble(final int r, final int c) {
		final Cell cell = this.getSimpleCell(r, c);
		return cell.getDoubleValue();
	}

	/** {@inheritDoc} */
	@Override
	public String getFormula(final int r, final int c) {
		final Cell cell = this.getSimpleCell(r, c);
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
		return this.table.getRowCount();
	}

	/** {@inheritDoc} */
	@Override
	public String getText(final int r, final int c) {
		final Cell cell = this.getSimpleCell(r, c);
		if (!"string".equals(cell.getValueType()))
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
	protected Cell getSimpleCell(final int r, final int c) {
		if (r != this.curR || this.curRow == null) {
			// Hack for clean style @see Table.appendRows(count, false)
			// 1. append "manually" the missing rows
			// 2. clean style
			final int lastIndex = this.getRowCount() - 1;
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
		// cell.getStyleHandler().getStyleElementForWrite();
		return cell;
	}
}