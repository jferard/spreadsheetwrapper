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

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Locale;

import org.jopendocument.dom.ODXMLDocument;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.table.Table;

import com.github.jferard.spreadsheetwrapper.impl.Stateful;

public class OdsSimpleodfStatefulDocument extends Stateful<SpreadsheetDocument> {
	public OdsSimpleodfStatefulDocument(Stateful<SpreadsheetDocument> sfDocument) {
		super(sfDocument.getValue(), sfDocument.isNew());
	}

	public Table getRawSheet(int s) {
		return this.value.getTableList().get(s);
	}

	public int getRawSheetCount() {
		return this.value.getSheetCount();
	}

	public Table getRawSheet(String sheetName) {
		return this.value.getTableByName(sheetName);
	}

	public void save(OutputStream outputStream) throws Exception {
		this.value.save(outputStream);
	}

	public OdfOfficeStyles getStyles() {
		return this.value.getDocumentStyles();
	}

	public void setLocale(Locale locale) {
		this.value.setLocale(locale);
	}

	public List<Table> getRawTableList() {
		return this.value.getTableList();
	}

	public Table newTable() {
		return 	Table.newTable(this.value);
	}

	public Table insertSheet(int index) {
		return this.value.insertSheet(index);
	}

	public void close() {
		this.value.close();
	}
	

}
