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
package com.github.jferard.spreadsheetwrapper.ods.jopendocument;

import java.io.IOException;
import java.io.OutputStream;

import org.jopendocument.dom.ODXMLDocument;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.impl.Stateful;

public class OdsJOpenStatefulDocument extends Stateful<SpreadSheet> {
	public OdsJOpenStatefulDocument(Stateful<SpreadSheet> sfDocument) {
		super(sfDocument.getValue(), sfDocument.isNew());
	}

	public Sheet getRawSheet(int s) {
		return this.value.getSheet(s);
	}

	public int getRawSheetCount() {
		return this.value.getSheetCount();
	}

	public Sheet getRawSheet(String sheetName) {
		return this.value.getSheet(sheetName);
	}

	public Sheet addRawSheet(int index, String sheetName) {
		return this.value.addSheet(index, sheetName);
	}

	public void save(OutputStream outputStream) throws IOException {
		this.value.getPackage().save(outputStream);
	}

	public ODXMLDocument getStyles() {
		return this.value.getPackage().getStyles();
	}
	

}
