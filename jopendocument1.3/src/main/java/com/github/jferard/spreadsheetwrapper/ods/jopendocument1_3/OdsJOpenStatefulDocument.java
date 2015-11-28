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

import java.io.IOException;
import java.io.OutputStream;

import org.jopendocument.dom.ODPackage;
import org.jopendocument.dom.ODXMLDocument;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.Stateful;

/**
 * Stateful means that there is a marker if the document is unitialized.
 *
 */
public class OdsJOpenStatefulDocument extends Stateful<SpreadSheet> {
	/**
	 * @param sfDocument
	 *            the wrapped document
	 */
	public OdsJOpenStatefulDocument(final Stateful<SpreadSheet> sfDocument) {
		super(sfDocument.getObject(), sfDocument.isNew());
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
	public Sheet addRawSheet(final int index, final String sheetName) {
		return this.object.addSheet(index, sheetName);
	}

	/**
	 * @param index
	 *            index of the sheet to get
	 * @return the *internal* sheet
	 */
	public Sheet getRawSheet(final int index) {
		return this.object.getSheet(index);
	}

	/**
	 * @param sheetName
	 *            name of the sheet to get
	 * @return the *internal* sheet
	 */
	public Sheet getRawSheet(final String sheetName) {
		return this.object.getSheet(sheetName);
	}

	/**
	 * @return the number of sheets in the document
	 */
	public int getRawSheetCount() {
		return this.object.getSheetCount();
	}

	/**
	 * @return the *internal* styles.xml document
	 */
	public ODXMLDocument getStyles() {
		final ODPackage odPackage = this.object.getPackage();
		return odPackage.getStyles();
	}

	/**
	 * @param outputStream
	 *            where to write
	 * @throws IOException
	 */
	public void save(final OutputStream outputStream) throws IOException {
		this.object.getPackage().save(outputStream);
	}
}
