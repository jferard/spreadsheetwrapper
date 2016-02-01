package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.io.IOException;
import java.util.List;

import jxl.Sheet;
import jxl.write.WritableSheet;
import jxl.write.WriteException;

public interface JxlWorkbook {

	int getNumberOfSheets();

	Sheet[] getSheets();

	void close() throws WriteException, IOException;

	String[] getSheetNames();

	void write() throws IOException;

	Sheet getSheet(String sheetName);

	WritableSheet createSheet(String sheetName, int index);
}
