package com.github.jferard.spreadsheetwrapper;

public class TestProperties {
	final String extension;
	final SpreadsheetDocumentFactory factory;
	
	public TestProperties(String extension, SpreadsheetDocumentFactory factory) {
		this.extension = extension;
		this.factory = factory;
	}

	public String getExtension() {
		return this.extension;
	}

	public SpreadsheetDocumentFactory getFactory() {
		return this.factory;
	}
}
