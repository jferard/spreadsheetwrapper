package com.github.jferard.spreadsheetwrapper.style;

public class Border {
	private double borderLineWidth;

	public Border setLineWidth(double borderLineWidth) {
		this.borderLineWidth = borderLineWidth;
		return this;
	}

	public double getLineWidth() {
		return this.borderLineWidth;		
	}

}
