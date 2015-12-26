package com.github.jferard.spreadsheetwrapper.style;

import java.util.Arrays;

import com.github.jferard.spreadsheetwrapper.Util;

public class Borders {
	private Border borderTop; 
	private Border borderBottom; 
	private Border borderLeft; 
	private Border borderRight; 
	
	public Borders setLineWidth(double borderLineWidth) {
		this.borderTop = (this.borderTop == null ? new Border() : this.borderTop).setLineWidth(borderLineWidth);
		this.borderBottom = (this.borderBottom == null ? new Border() : this.borderBottom).setLineWidth(borderLineWidth);
		(this.borderLeft == null ? new Border() : this.borderRight).setLineWidth(borderLineWidth);
		(this.borderRight == null ? new Border() : this.borderRight).setLineWidth(borderLineWidth);
		return this;
	}

	public double getLineWidth() {
		double lineWidth = this.borderTop == null ? WrapperCellStyle.NO_LINE : this.borderTop.getLineWidth();
		for (Border border : Arrays.asList(this.borderBottom, this.borderLeft, this.borderRight)) {
				if (border == null || !Util.almostEqual(lineWidth, border.getLineWidth()))
					return WrapperCellStyle.NO_LINE;
		}
		return lineWidth;
	}
	
	public Border getBorderTop() {
		return this.borderTop;
	}

	public Borders setBorderTop(Border borderTop) {
		this.borderTop = borderTop;
		return this;
	}

	public Border getBorderBottom() {
		return this.borderBottom;
	}

	public Borders setBorderBottom(Border borderBottom) {
		this.borderBottom = borderBottom;
		return this;
	}

	public Border getBorderLeft() {
		return this.borderLeft;
	}

	public Borders setBorderLeft(Border borderLeft) {
		this.borderLeft = borderLeft;
		return this;
	}

	public Border getBorderRight() {
		return this.borderRight;
	}

	public Borders setBorderRight(Border borderRight) {
		this.borderRight = borderRight;
		return this;
	}

	public void setLineType(final String type) {
		// TODO Auto-generated method stub
		
	}

	public void setLineColor(final String colorAsHex) {
		// TODO Auto-generated method stub
		
	}

}
