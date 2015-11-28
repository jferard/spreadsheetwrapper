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
package com.github.jferard.spreadsheetwrapper.ods.jopendocument1_3;

import java.util.Date;
import java.util.List;

import org.jopendocument.dom.OOUtils;
import org.jopendocument.dom.spreadsheet.CellStyle;
import org.jopendocument.dom.spreadsheet.CellStyle.StyleTableCellProperties;
import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;
import org.jopendocument.dom.text.TextStyle.StyleTextProperties;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.StyleUtility;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 */
class OdsJOpenWriter extends AbstractSpreadsheetWriter implements
		SpreadsheetWriter {
	/** the *internal* sheet wrapped */
	private final Sheet sheet;

	/**
	 * @param sheet
	 *            the *internal* sheet
	 */
	OdsJOpenWriter(final Sheet sheet) {
		super(new OdsJOpenReader(sheet));
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
		cell.setValue(value);
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

	@Override
	public boolean setStyle(final int r, final int c,
			final WrapperCellStyle wrapperStyle) {
		final MutableCell<SpreadSheet> cell = this.getOrCreateCell(r, c);
		final CellStyle cellStyle = cell.getStyle();
		final WrapperColor backgroundColor = wrapperStyle.getBackgroundColor();
		if (backgroundColor != null) {
			final StyleTableCellProperties tableCellProperties = cellStyle
					.getTableCellProperties();
			tableCellProperties.setBackgroundColor(backgroundColor.toHex());
		}
		final WrapperFont font = wrapperStyle.getCellFont();
		if (font != null) {
			final StyleTextProperties textProperties = cellStyle
					.getTextProperties();
			final WrapperColor color = font.getColor();
			if (color != null)
				textProperties.setColor(OOUtils.decodeRGB(color.toHex()));
			if (font.getBold() == WrapperCellStyle.YES)
				textProperties.getElement().setAttribute(
						OdsConstants.FONT_WEIGHT, "bold");
		}
		return true;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c, final String styleName) {
		final MutableCell<SpreadSheet> cell = this.getOrCreateCell(r, c);
		cell.setStyleName(styleName);
		return true;
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
}