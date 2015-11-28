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
package com.github.jferard.spreadsheetwrapper.ods.jopendocument1_2;

import java.math.BigDecimal;
import java.util.Date;

import org.jdom.Element;
import org.jopendocument.dom.ODValueType;
import org.jopendocument.dom.spreadsheet.CellStyle;
import org.jopendocument.dom.spreadsheet.CellStyle.SyleTableCellProperties;
import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;
import org.jopendocument.dom.text.TextStyle.SyleTextProperties;

import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;

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
		int rowCount = this.sheet.getRowCount();
		if (rowCount == 1 && this.getCellCount(0) == 0)
			rowCount = 0;
		return rowCount;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/WrapperCellStyle getStyle(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		if (cell == null)
			return null;

		final CellStyle cellStyle = cell.getStyle();
		if (cellStyle == null)
			return WrapperCellStyle.EMPTY;

		final SyleTableCellProperties tableCellProperties = cellStyle
				.getTableCellProperties();
		final String bColorAsHex = tableCellProperties.getRawBackgroundColor();
		final WrapperColor backgroundColor = WrapperColor
				.stringToColor(bColorAsHex);

		final WrapperFont wrapperFont = new WrapperFont();
		final SyleTextProperties textProperties = cellStyle.getTextProperties();
		final Element odfElement = textProperties.getElement();
		final String fColorAsHex = odfElement.getAttributeValue(
				OdsConstants.COLOR_ATTR_NAME, OdsJOpenStyleHelper.FO_NS);
		final WrapperColor fontColor = WrapperColor.stringToColor(fColorAsHex);
		if (fontColor != null)
			wrapperFont.setColor(fontColor);
		final String fontWeight = odfElement.getAttributeValue(
				OdsConstants.FONT_WEIGHT, OdsJOpenStyleHelper.FO_NS);
		if (OdsConstants.BOLD_ATTR_NAME.equals(fontWeight))
			wrapperFont.setBold();

		return new WrapperCellStyle(backgroundColor, wrapperFont);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getStyleName(final int r, final int c) {
		final MutableCell<SpreadSheet> cell = this.getCell(r, c);
		if (cell == null)
			return null;

		return cell.getStyleName();
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

	private String getFormula(final MutableCell<SpreadSheet> cell) {
		final Element element = cell.getElement();
		return element.getAttributeValue(OdsConstants.FORMULA_ATTR_NAME,
				element.getNamespace());
	}

	private String getTypeName(final MutableCell<SpreadSheet> cell) {
		final ODValueType valueType = cell.getValueType();
		return valueType == null ? OdsConstants.STRING_TYPE : valueType
				.getName();
	}
}