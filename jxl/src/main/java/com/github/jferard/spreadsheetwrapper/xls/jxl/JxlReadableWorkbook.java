package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Locale;

import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetException;

public class JxlReadableWorkbook extends JxlWorkbook {
	private final Workbook workbook;

	public JxlReadableWorkbook(InputStream inputStream) 
			throws SpreadsheetException {
		try {
			this.workbook = Workbook.getWorkbook(inputStream,
					JxlReadableWorkbook.getReadSettings());
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		} catch (final BiffException e) {
			throw new SpreadsheetException(e);
		}
	}

	@Override
	public int getNumberOfSheets() {
		return this.workbook.getNumberOfSheets();
	}

	@Override
	public Sheet[] getSheets() {
		return this.workbook.getSheets();
	}

	@Override
	public void close() {
		this.workbook.close();
	}

	@Override
	public String[] getSheetNames() {
		return this.workbook.getSheetNames();
	}

	@Override
	public void write() {
		throw new UnsupportedOperationException();
	}

	@Override
	public Sheet getSheet(String sheetName) {
		return this.workbook.getSheet(sheetName);
	}

	@Override
	public WritableSheet createSheet(String sheetName, int index) {
		throw new UnsupportedOperationException();
	}

	private static WorkbookSettings getReadSettings() {
		final WorkbookSettings settings = new WorkbookSettings();
		settings.setLocale(Locale.US);
		settings.setEncoding("windows-1252");
		return settings;
	}
}
