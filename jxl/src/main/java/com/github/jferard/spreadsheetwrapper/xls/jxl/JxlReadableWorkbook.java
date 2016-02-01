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

public class JxlReadableWorkbook implements JxlWorkbook {
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

	
	/** {@inheritDoc} */
	@Override
	public int getNumberOfSheets() {
		return this.workbook.getNumberOfSheets();
	}

	/** {@inheritDoc} */
	@Override
	public Sheet[] getSheets() {
		return this.workbook.getSheets();
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		this.workbook.close();
	}

	/** {@inheritDoc} */
	@Override
	public String[] getSheetNames() {
		return this.workbook.getSheetNames();
	}

	/** {@inheritDoc} */
	@Override
	public void write() {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public Sheet getSheet(String sheetName) {
		return this.workbook.getSheet(sheetName);
	}

	/** {@inheritDoc} */
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
