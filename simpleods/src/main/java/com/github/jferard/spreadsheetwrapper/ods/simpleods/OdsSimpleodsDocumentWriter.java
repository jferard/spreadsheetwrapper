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
package com.github.jferard.spreadsheetwrapper.ods.simpleods;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.simpleods.BorderStyle;
import org.simpleods.OdsFile;
import org.simpleods.SimpleOdsException;
import org.simpleods.Table;
import org.simpleods.TableStyle;
import org.simpleods.TextStyle;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.Output;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.Util;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetWriterCursorImpl;
import com.github.jferard.spreadsheetwrapper.style.Borders;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;

/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

/**
 */
public class OdsSimpleodsDocumentWriter extends
		AbstractSpreadsheetDocumentWriter implements SpreadsheetDocumentWriter {
	
	private SpreadsheetWriter createNew(final Table table) {
		return new OdsSimpleodsWriter(table);
	}

	/** *internal* workbook */
	private final OdsFile file;

	/** logger */
	private final Logger logger;

	/**
	 * @param logger
	 *            simple logger
	 * @param file
	 *            *internal* value
	 * @param outputURL
	 *            where to write
	 */
	public OdsSimpleodsDocumentWriter(final Logger logger, final OdsFile file,
			final Output output) {
		super(logger, output);
		this.logger = logger;
		this.file = file;
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		try {
			this.output.close();
		} catch (final IOException e) {
			final String message = e.getMessage();
			this.logger.log(Level.SEVERE, message == null ? "" : message, e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursorByIndex(final int index) {
		return new SpreadsheetWriterCursorImpl(this.getSpreadsheet(index));
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursorByName(final String sheetName)
			throws SpreadsheetException {
		return new SpreadsheetWriterCursorImpl(this.getSpreadsheet(sheetName));
	}

	/** {@inheritDoc} */
	@Override
	public void save() throws SpreadsheetException {
		final OutputStream outputStream = this.output.getStream();
		if (outputStream == null)
			throw new IllegalStateException(
					String.format("Use saveAs when output file/stream is not specified"));
		if (!this.file.save(outputStream))
			throw new SpreadsheetException(String.format(
					"this.spreadsheetDocument.save(%s) not ok", outputStream));
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyle(final String styleName,
			final WrapperCellStyle wrapperCellStyle) {
		final TableStyle newStyle = new TableStyle(TableStyle.STYLE_TABLECELL,
				styleName, this.file);
		final WrapperFont wrapperFont = wrapperCellStyle.getCellFont();
		if (wrapperFont != null) {
			final TextStyle textStyle = newStyle.getTextStyle();
			final WrapperColor fontColor = wrapperFont.getColor();
			final int bold = wrapperFont.getBold();
			final int italic = wrapperFont.getItalic();
			final double size = wrapperFont.getSize();
			final String family = wrapperFont.getFamily();

			if (fontColor != null)
				textStyle.setFontColor(fontColor.toHex());
			if (bold == WrapperCellStyle.YES) {
				textStyle.setFontWeightBold();
				if (italic == WrapperCellStyle.YES)
					textStyle.setFontWeightItalic();
			} else {
				if (italic == WrapperCellStyle.YES)
					textStyle.setFontWeightItalic();
				else
					textStyle.setFontWeightNormal();
			}
			if (!Util.almostEqual(size, WrapperCellStyle.DEFAULT))
				textStyle.setFontSize((int) size);
			if (family != null)
				textStyle.setFontName(family);

		}
		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();
		final Borders borders = wrapperCellStyle.getBorders();
		if (backgroundColor != null)
			newStyle.setBackgroundColor(backgroundColor.toHex());
		if (borders != null) {
			final double borderLineWidth = borders.getLineWidth();
			if (borderLineWidth != WrapperCellStyle.DEFAULT)
				newStyle.addBorderStyle(Double.valueOf(borderLineWidth) + "pt",
						"#000000", BorderStyle.BORDER_SOLID,
						BorderStyle.POSITION_ALL);
		}
		return true;
	}

	/** {@inheritDoc} */
	@Override
	public WrapperCellStyle getCellStyle(final String styleName) {
		throw new UnsupportedOperationException();
	}

	/** */
	@Override
	public int getSheetCount() {
		return this.file.lastTableNumber();
	}

	/** */
	@Override
	public List<String> getSheetNames() {
		final int count = this.getSheetCount();
		final List<String> sheetNames = new ArrayList<String>(count);
		try {
			for (int i = 0; i < count; i++)
				sheetNames.add(this.file.getTableName(i));
		} catch (final SimpleOdsException e) {
			throw new AssertionError(e);
		}
		return sheetNames;
	}

	/**
	 * @param index
	 *            index of the sheet
	 * @return the reader/writer
	 */
	public SpreadsheetWriter getSpreadsheet(final int index) {
		final SpreadsheetWriter spreadsheet;
		if (this.accessor.hasByIndex(index))
			spreadsheet = this.accessor.getByIndex(index);
		else {
			final int count = this.getSheetCount();
			if (index < 0 || index >= count)
				throw new NoSuchElementException(String.format(
						"No sheet at position %d", index));

			try {
				final String name = this.file.getTableName(index);
				final Table table = this.getTable(index);
				spreadsheet = this.createNew(table);
				this.accessor.put(name, index, spreadsheet);
			} catch (final SimpleOdsException e) {
				throw new AssertionError(String.format(
						"Can't reach sheet at position %d", index), e);
			}
		}
		return spreadsheet;
	}

	private Table getTable(
			/*>>> @UnknownInitialization AbstractOdsSimpleodsDocumentDelegate<T> this, */final int index) {
		if (this.file == null)
			throw new IllegalStateException();

		return (Table) this.file.getContent().getTableQueue().get(index);
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetWriter addSheetWithCheckedIndex(final int index, final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		if (index != this.getSheetCount())
			throw new UnsupportedOperationException();

		final int indexForSheetName = this.file.getTableNumber(sheetName);
		if (indexForSheetName != -1)
			throw new IllegalArgumentException(String.format("Sheet %s exists",
					sheetName));

		// at the end
		final SpreadsheetWriter spreadsheet;
		try {
			this.file.addTable(sheetName);
			final Table table = this.getTable(index);
			spreadsheet = this.createNew(table);
			this.accessor.put(sheetName, index, spreadsheet);
		} catch (final SimpleOdsException e) {
			throw new CantInsertElementInSpreadsheetException(e);
		}
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetWriter findSpreadsheetAndCreateReaderOrWriter(final String sheetName) {
		final int index = this.file.getTableNumber(sheetName);
		if (index == -1)
			throw new NoSuchElementException(String.format(
					"No %s sheet in workbook", sheetName));

		final Table table = this.getTable(index);
		final SpreadsheetWriter spreadsheet = this.createNew(table);
		this.accessor.put(sheetName, index, spreadsheet);
		return spreadsheet;
	}
}
