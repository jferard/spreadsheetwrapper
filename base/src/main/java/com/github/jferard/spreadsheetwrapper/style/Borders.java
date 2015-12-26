package com.github.jferard.spreadsheetwrapper.style;

public class Borders {
	private double borderLineWidth;

	public Borders setLineWidth(double borderLineWidth) {
		this.borderLineWidth = borderLineWidth;
		return this;
	}

	public double getLineWidth() {
		return this.borderLineWidth;		
	}

}
