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
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.simpleods.OdsFile;
import org.simpleods.Table;
import org.simpleods.TableStyle;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.Output;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetWriterCursorImpl;
import com.github.jferard.spreadsheetwrapper.impl.StyleUtility;

/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

/**
 */
public class OdsSimpleodsDocumentWriter extends
		AbstractSpreadsheetDocumentWriter implements SpreadsheetDocumentWriter {
	/** class for delegation */
	private final class OdsSimpleodsDocumentWriterTrait extends
			AbstractOdsSimpleodsDocumentTrait<SpreadsheetWriter> {
		OdsSimpleodsDocumentWriterTrait(final OdsFile file) {
			super(file);
		}

		/** {@inheritDoc} */
		@Override
		protected SpreadsheetWriter createNew(
				/*>>> @UnknownInitialization OdsSimpleodsDocumentWriterTrait this, */final Table table) {
			return new OdsSimpleodsWriter(table);
		}
	}

	/** for delegation */
	private final AbstractOdsSimpleodsDocumentTrait<SpreadsheetWriter> documentTrait;

	/** *internal* workbook */
	private final OdsFile file;

	/** logger */
	private final Logger logger;

	/** for delegation */
	private final OdsSimpleodsDocumentReader reader;

	private final StyleUtility styleUtility;

	/**
	 * @param logger
	 *            simple logger
	 * @param file
	 *            *internal* value
	 * @param outputURL
	 *            where to write
	 */
	public OdsSimpleodsDocumentWriter(final Logger logger,
			final StyleUtility styleUtility, final OdsFile file,
			final Output output) {
		super(logger, output);
		this.styleUtility = styleUtility;
		this.reader = new OdsSimpleodsDocumentReader(file);
		this.logger = logger;
		this.file = file;
		this.documentTrait = new OdsSimpleodsDocumentWriterTrait(file);
	}

	/**
	 * {@inheritDoc}
	 *
	 * @throws CantInsertElementInSpreadsheetException
	 * @throws IndexOutOfBoundsException
	 */
	@Override
	public SpreadsheetWriter addSheet(final int index, final String sheetName)
			throws IndexOutOfBoundsException,
			CantInsertElementInSpreadsheetException {
		return this.documentTrait.addSheet(index, sheetName);
	}

	/**
	 * {@inheritDoc}
	 *
	 * @throws CantInsertElementInSpreadsheetException
	 */
	@Override
	public SpreadsheetWriter addSheet(final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		return this.documentTrait.addSheet(sheetName);
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
		this.reader.close();
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public boolean createStyle(final String styleName, final String styleString) {
		// add to content because of background-color.
		final TableStyle newStyle = new TableStyle(TableStyle.STYLE_TABLECELL,
				styleName, this.file);
		final Map<String, String> props = this.styleUtility
				.getPropertiesMap(styleString);
		for (final Map.Entry<String, String> entry : props.entrySet()) {
			if (entry.getKey().equals("font-weight")) {
				if (entry.getValue().equals("bold"))
					newStyle.setFontWeightBold();
				if (entry.getValue().equals("italic"))
					newStyle.setFontWeightItalic();
			} else if (entry.getKey().equals("background-color")) {
				newStyle.setBackgroundColor(entry.getValue());
			}
		}
		return true;
	}

	/** {@inheritDoc} */
	@Override
	public WrapperCellStyle getCellStyle(final String styleName) {
		return this.reader.getCellStyle(styleName);
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
	public int getSheetCount() {
		return this.reader.getSheetCount();
	}

	/** {@inheritDoc} */
	@Override
	public List<String> getSheetNames() {
		return this.reader.getSheetNames();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter getSpreadsheet(final int index) {
		return this.documentTrait.getSpreadsheet(index);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter getSpreadsheet(final String sheetName) {
		return this.documentTrait.getSpreadsheet(sheetName);
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public String getStyleString(final String styleName) {
		return this.reader.getStyleString(styleName);
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
			if (wrapperFont.getBold() == WrapperCellStyle.YES)
				newStyle.setFontWeightBold();
			if (wrapperFont.getItalic() == WrapperCellStyle.YES)
				newStyle.setFontWeightItalic();
		}
		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();
		if (backgroundColor != null)
			newStyle.setBackgroundColor(backgroundColor.toHex());
		return true;
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public boolean updateStyle(final String styleName, final String styleString) {
		/** @see javadoc */
		return this.createStyle(styleName, styleString);
	}

}
