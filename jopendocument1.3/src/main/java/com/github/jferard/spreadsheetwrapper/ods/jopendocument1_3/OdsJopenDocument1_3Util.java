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
	public static String getFormula(final MutableCell<SpreadSheet> cell) {
		return cell.getFormula();
	}

	public static SpreadSheet getSpreadSheet(final ODPackage odPackage) {
		return odPackage.getSpreadSheet();
	}

	public static ODXMLDocument getStyles(final SpreadSheet spreadSheet) {
		final ODPackage odPackage = spreadSheet.getPackage();
		return odPackage.getStyles();
	}

	public static SpreadSheet newSpreadsheetDocument()
			throws SpreadsheetException {
		try {
			return SpreadSheet.createEmpty(new DefaultTableModel());
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	public static void setValue(final MutableCell<SpreadSheet> cell,
			final ODValueType type, final Object value) {
		cell.setValue(value);
	}
}
