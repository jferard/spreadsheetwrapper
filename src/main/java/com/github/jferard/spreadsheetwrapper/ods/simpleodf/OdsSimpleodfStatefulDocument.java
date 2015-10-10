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

import java.io.OutputStream;
import java.util.List;
import java.util.Locale;

import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.table.Table;

import com.github.jferard.spreadsheetwrapper.impl.Stateful;

/**
 * A OdsSimpleodfStatefulDocument is a specific Stateful SpreadsheetDocument.
 */
public class OdsSimpleodfStatefulDocument extends Stateful<SpreadsheetDocument> {
	/**
	 * @param sfDocument
	 *            a stateful SpreadsheetDocument, ie a SpreadsheetDocument that
	 *            is, or not, initialized
	 */
	public OdsSimpleodfStatefulDocument(
			final Stateful<SpreadsheetDocument> sfDocument) {
		super(sfDocument.getObject(), sfDocument.isNew());
	}

	/**
	 * close the document
	 */
	public void close() {
		this.object.close();
	}

	/**
	 * Raw means that we use the object directly
	 * 
	 * @param sheetIndex
	 *            the index of the sheet
	 * @return the sheet in the document
	 */
	public Table getRawSheet(final int sheetIndex) {
		return this.object.getTableList().get(sheetIndex);
	}

	/**
	 * @param sheetName
	 *            the name of the sheet
	 * @return the sheet in the document
	 */
	public Table getRawSheet(final String sheetName) {
		return this.object.getTableByName(sheetName);
	}

	/**
	 * @return the sheetCount of the document
	 */
	public int getRawSheetCount() {
		return this.object.getSheetCount();
	}

	/**
	 * @return the list of the sheets in the document
	 */
	public List<Table> getRawTableList() {
		return this.object.getTableList();
	}

	/**
	 * @return the document with the odf styles
	 */
	public OdfOfficeStyles getStyles() {
		return this.object.getDocumentStyles();
	}

	/**
	 * @param index
	 *            the index of the sheet after the new sheet
	 * @return the sheet
	 */
	public Table rawInsertSheet(final int index) {
		return this.object.insertSheet(index);
	}

	/**
	 * @return a new sheet on the object
	 */
	public Table rawNewTable() {
		return Table.newTable(this.object);
	}

	/**
	 * @param outputStream
	 *            the stream to write to
	 * @throws Exception
	 *             if odftoolkit throws an exception !
	 */
	public void rawSave(final OutputStream outputStream) throws Exception {
		this.object.save(outputStream);
	}

	/**
	 * @param locale
	 *            the locale to set
	 */
	public void rawSetLocale(final Locale locale) {
		this.object.setLocale(locale);
	}

}
