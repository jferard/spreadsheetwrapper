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
package com.github.jferard.spreadsheetwrapper.impl;

import java.util.List;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReaderCursor;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>> import org.checkerframework.checker.nullness.qual.MonotonicNonNull;*/

public class SpreadsheetDocumentReaderImpl implements
SpreadsheetDocumentReader {
	private SpreadsheetDocumentWriter spreadsheetDocumentWriter;

	public SpreadsheetDocumentReaderImpl(SpreadsheetDocumentWriter spreadsheetDocumentWriter) {
		this.spreadsheetDocumentWriter = spreadsheetDocumentWriter;
		
	}
	
	/** {inheritDoc} */
	@Override
	public void close() throws SpreadsheetException {
		this.spreadsheetDocumentWriter.close();
	}

	/** {inheritDoc} */
	@Override
	public WrapperCellStyle getCellStyle(String styleName) {
		return this.spreadsheetDocumentWriter.getCellStyle(styleName);
	}

	/** {inheritDoc} */
	@Override
	public int getSheetCount() {
		return this.spreadsheetDocumentWriter.getSheetCount();
	}

	/** {inheritDoc} */
	@Override
	public List<String> getSheetNames() {
		return this.spreadsheetDocumentWriter.getSheetNames();
	}

	/** {inheritDoc} */
	@Override
	public SpreadsheetReader getSpreadsheet(int index) {
		return new SpreadsheetReaderImpl(this.spreadsheetDocumentWriter.getSpreadsheet(index));
	}

	/** {inheritDoc} */
	@Override
	public SpreadsheetReader getSpreadsheet(String name) {
		return new SpreadsheetReaderImpl(this.spreadsheetDocumentWriter.getSpreadsheet(name));
	}
	
	/** {inheritDoc} */
	@Override
	public SpreadsheetReaderCursor getNewCursorByIndex(int index) {
		return this.spreadsheetDocumentWriter.getNewCursorByIndex(index);
	}

	/** {inheritDoc} */
	@Override
	public SpreadsheetReaderCursor getNewCursorByName(String sheetName)
			throws SpreadsheetException {
		return this.spreadsheetDocumentWriter.getNewCursorByName(sheetName);
	}
	
}