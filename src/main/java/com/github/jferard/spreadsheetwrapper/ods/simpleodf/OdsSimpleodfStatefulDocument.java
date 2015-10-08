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

public class OdsSimpleodfStatefulDocument extends Stateful<SpreadsheetDocument> {
	/**
	 * @param sfDocument a stateful SpreadsheetDocument, ie a SpreadsheetDocument that is, or not, initialized 
	 */
	public OdsSimpleodfStatefulDocument(
			final Stateful<SpreadsheetDocument> sfDocument) {
		super(sfDocument.getObject(), sfDocument.isNew());
	}

	public void close() {
		this.object.close();
	}

	public Table getRawSheet(final int sheetIndex) {
		return this.object.getTableList().get(sheetIndex);
	}

	public Table getRawSheet(final String sheetName) {
		return this.object.getTableByName(sheetName);
	}

	public int getRawSheetCount() {
		return this.object.getSheetCount();
	}

	public List<Table> getRawTableList() {
		return this.object.getTableList();
	}

	public OdfOfficeStyles getStyles() {
		return this.object.getDocumentStyles();
	}

	public Table insertSheet(final int index) {
		return this.object.insertSheet(index);
	}

	public Table newTable() {
		return Table.newTable(this.object);
	}

	public void save(final OutputStream outputStream) throws Exception {
		this.object.save(outputStream);
	}

	public void setLocale(final Locale locale) {
		this.object.setLocale(locale);
	}

}
