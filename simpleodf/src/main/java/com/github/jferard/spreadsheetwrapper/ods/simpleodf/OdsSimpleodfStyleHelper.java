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

import org.odftoolkit.odfdom.dom.element.table.TableTableCellElementBase;
import org.odftoolkit.odfdom.dom.style.OdfStyleFamily;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;
import org.odftoolkit.simple.table.Cell;

import com.github.jferard.spreadsheetwrapper.ods.apache.OdsApacheStyleHelper;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/**
 * A little style utility
 *
 */
public class OdsSimpleodfStyleHelper {
	private OdsApacheStyleHelper styleHelperDelegate;

	OdsSimpleodfStyleHelper() {
		this.styleHelperDelegate = new OdsApacheStyleHelper();
	}
	
	public WrapperCellStyle getCellStyle(final OdfOfficeStyles documentStyles, final String styleName) {
		final OdfStyle existingStyle = documentStyles.getStyle(styleName,
				OdfStyleFamily.TableCell);
		return this.styleHelperDelegate.toWrapperCellStyle(existingStyle);
	}

	public /*@Nullable*/ WrapperCellStyle getStyle(Cell cell) {
		if (cell == null)
			return null;

		final TableTableCellElementBase odfElement = cell.getOdfElement();
		return this.styleHelperDelegate.getWrapperCellStyle(odfElement);
	}

	public boolean setStyle(Cell cell, WrapperCellStyle wrapperCellStyle) {
		if (cell == null)
			return false;

		final TableTableCellElementBase odfElement = cell.getOdfElement();
		this.styleHelperDelegate.setWrapperCellStyle(odfElement, wrapperCellStyle);
		return true;
	}

	public boolean setStyle(OdfOfficeStyles documentStyles, String styleName,
			WrapperCellStyle wrapperCellStyle) {
		final OdfStyle newStyle = documentStyles.newStyle(styleName,
				OdfStyleFamily.TableCell);
		this.styleHelperDelegate.setWrapperCellStyle(newStyle, wrapperCellStyle);
		return true;
	}
}	