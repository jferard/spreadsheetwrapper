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

import javax.swing.table.DefaultTableModel;

import org.jopendocument.dom.ODPackage;
import org.jopendocument.dom.ODValueType;
import org.jopendocument.dom.ODXMLDocument;
import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetException;

/**
 * Utility class for 1.3 version.
 * @author Julien
 *
 */
final class OdsJopenDocument1_3Util {
	/**
	 * @param cell
	 *           the cell
	 * @return the formula
	 */
	public static String getFormula(final MutableCell<SpreadSheet> cell) {
		return cell.getFormula();
	}
	
	/**
	 * @param odPackage
	 *           the zip file representation
	 * @return
	 *           a new spreadsheet
	 */
	public static SpreadSheet getSpreadSheet(final ODPackage odPackage) {
		return odPackage.getSpreadSheet();
	}

	/**
	 * @param spreadSheet the *internal* document.xml
	 * @return the styles as *internal* document.xml
	 */
	public static ODXMLDocument getStyles(final SpreadSheet spreadSheet) {
		final ODPackage odPackage = spreadSheet.getPackage();
		return odPackage.getStyles();
	}

	/**
	 * @return a new *internal* document.xml
	 * @throws SpreadsheetException
	 */
	public static SpreadSheet newSpreadsheetDocument()
			throws SpreadsheetException {
		try {
			return SpreadSheet.createEmpty(new DefaultTableModel());
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	/**
	 * @param cell the internal cell
	 * @param type the internal type
	 * @param value the value to set
	 */
	public static void setValue(final MutableCell<SpreadSheet> cell,
			final ODValueType type, final Object value) {
		cell.setValue(value);
	}

	private OdsJopenDocument1_3Util() {}
}
