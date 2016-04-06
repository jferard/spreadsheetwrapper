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

import java.io.IOException;
import java.io.OutputStream;
import java.util.NoSuchElementException;

import org.jopendocument.dom.ODXMLDocument;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.UnknownInitialization;

/**
 * Initializable means that there is a marker if the document is unitialized.
 *
 */
class InitializableDocument {
	/** true if initialized */
	private boolean initialized;

	/** the object */
	private final SpreadSheet document;

	static final InitializableDocument createInitialized(
			SpreadSheet document) {
		return new InitializableDocument(document, true);
	}

	static final InitializableDocument createUninitialized(
			SpreadSheet document) {
		return new InitializableDocument(document, false);
	}
	
	
	/**
	 * @param uiDocument
	 *            the wrapped document
	 */
	InitializableDocument(final SpreadSheet document,
			boolean initialized) {
		this.document = document;
		this.initialized = initialized;
	}

	/**
	 * Delegator for addSheet
	 *
	 * @param index
	 *            index of the sheet after the sheet to be inserted
	 * @param sheetName
	 *            the name of the new sheet.
	 * @return the *internal* sheet
	 */
	public Sheet addSheet(final int index, final String sheetName) {
		Sheet sheet;
		if (this.initialized) {
			sheet = this.getSheet(sheetName);
			if (sheet != null)
				throw new IllegalArgumentException(
						String.format("Sheet %s exists", sheetName));

			sheet = this.document.addSheet(index, sheetName);
		} else {
			assert this.document.getSheetCount() == 1 && index == 0;
			sheet = this.document.getSheet(0);
			sheet.setName(sheetName);
			this.initialized = true;
		}
		return sheet;
	}

	/**
	 * @param index
	 *            index of the sheet to get
	 * @return the *internal* sheet
	 */
	public Sheet getSheet(final int index) {
		return this.document.getSheet(index);
	}

	/**
	 * @param sheetName
	 *            name of the sheet to get
	 * @return the *internal* sheet
	 */
	public Sheet getSheet(final String sheetName) {
		return this.document.getSheet(sheetName);
	}

	/**
	 * @return the number of sheets in the document
	 */
	public int getSheetCount() {
		int count;
		if (this.initialized)
			count = this.document.getSheetCount();
		else // 1 dummy sheet exists
			count = 0;
		return count;
	}

	/**
	 * @return the *internal* styles.xml document
	 */
	public ODXMLDocument getStyles() {
		return ${jopendocument.util.cls}.getStyles(this.document);
	}

	/**
	 * @param outputStream
	 *            where to write
	 * @throws IOException
	 */
	public void save(final OutputStream outputStream) throws IOException {
		this.document.getPackage().save(outputStream);
	}
}
