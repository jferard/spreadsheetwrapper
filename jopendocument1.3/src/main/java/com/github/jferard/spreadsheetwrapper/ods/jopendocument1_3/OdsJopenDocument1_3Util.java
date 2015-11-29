package com.github.jferard.spreadsheetwrapper.ods.jopendocument1_3;

import java.io.IOException;

import javax.swing.table.DefaultTableModel;

import org.jopendocument.dom.ODPackage;
import org.jopendocument.dom.ODValueType;
import org.jopendocument.dom.ODXMLDocument;
import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetException;

public class OdsJopenDocument1_3Util {
	public static SpreadSheet getSpreadSheet(final ODPackage odPackage) {
		return odPackage.getSpreadSheet();
	}
	
	public static SpreadSheet newSpreadsheetDocument()
					throws SpreadsheetException {
		try {
			return SpreadSheet.createEmpty(new DefaultTableModel());
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}
	
	public static String getFormula(final MutableCell<SpreadSheet> cell) {
		return cell.getFormula();
	}
	
	public static ODXMLDocument getStyles(SpreadSheet spreadSheet) {
		final ODPackage odPackage = spreadSheet.getPackage();
		return odPackage.getStyles();
	}

	public static void setValue(MutableCell<SpreadSheet> cell, ODValueType type,
			Object value) {
		cell.setValue(value);
	}
}
