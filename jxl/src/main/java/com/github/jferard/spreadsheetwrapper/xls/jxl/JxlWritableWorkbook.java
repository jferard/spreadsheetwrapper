package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Locale;

import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import com.github.jferard.spreadsheetwrapper.SpreadsheetException;

public class JxlWritableWorkbook extends JxlWorkbook {

	private WritableWorkbook writableWorkbook;

	public JxlWritableWorkbook(OutputStream outputStream) throws SpreadsheetException {
		try {
			this.writableWorkbook = Workbook.createWorkbook(outputStream,
					JxlWritableWorkbook.getWriteSettings());
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	public JxlWritableWorkbook(InputStream inputStream,
			OutputStream outputStream) throws SpreadsheetException {
		if (outputStream == null)
			throw new UnsupportedOperationException();

		try {
			final Workbook workbook = Workbook.getWorkbook(inputStream,
					getReadSettings());
			this.writableWorkbook = Workbook.createWorkbook(
					outputStream, workbook, getWriteSettings());
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		} catch (final BiffException e) {
			throw new SpreadsheetException(e);
		}
	}

	public JxlWritableWorkbook(File inputFile,
			File outputFile) throws SpreadsheetException {
		if (outputFile == null)
			throw new IllegalStateException();

		try {
			final Workbook workbook = Workbook.getWorkbook(inputFile,
					getReadSettings());
			this.writableWorkbook = Workbook.createWorkbook(
					outputFile, workbook, getWriteSettings());
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
		return this.writableWorkbook.getNumberOfSheets();
	}

	@Override
	public Sheet[] getSheets() {
		return this.writableWorkbook.getSheets();
	}

	@Override
	public void close() throws WriteException, IOException {
		this.writableWorkbook.close(); // ?
	}

	@Override
	public String[] getSheetNames() {
		return this.writableWorkbook.getSheetNames();
	}

	@Override
	public void write() throws IOException {
		this.writableWorkbook.write();
	}

	@Override
	public Sheet getSheet(String sheetName) {
		return this.writableWorkbook.getSheet(sheetName);
	}

	@Override
	public WritableSheet createSheet(String sheetName, int index) {
		return this.writableWorkbook.createSheet(sheetName, index);
	}

	private static WorkbookSettings getWriteSettings() {
		final WorkbookSettings settings = JxlWritableWorkbook.getReadSettings();
		settings.setWriteAccess("spreadsheetwrapper");
		return settings;
	}
	
	private static WorkbookSettings getReadSettings() {
		final WorkbookSettings settings = new WorkbookSettings();
		settings.setLocale(Locale.US);
		settings.setEncoding("windows-1252");
		return settings;
	}
}
