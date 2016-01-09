package com.github.jferard.spreadsheetwrapper.style;

import com.github.jferard.spreadsheetwrapper.Util;

public class Border {
	private double borderLineWidth;
	private WrapperColor color;
	private Object lineType;

	public Border setLineWidth(double borderLineWidth) {
		this.borderLineWidth = borderLineWidth;
		return this;
	}

	public double getLineWidth() {
		return this.borderLineWidth;		
	}

	public Border setLineColor(WrapperColor color) {
		this.color = color;
		return this;
	}

	public WrapperColor getLineColor() {
		return this.color;
	}
	
	public Object getLineType() {
		return this.lineType;
	}
	
	public Border setLineType(Object lineType) {
		this.lineType = lineType;
		return this;
	}
	
	/** {@inheritDoc} */
	@Override
	public boolean equals(final/*@Nullable*/Object obj) {
		if (this == obj)
			return true;
		if (!(obj instanceof Border))
			return false;

		final Border other = (Border) obj;
		return Util.almostEqual(this.borderLineWidth, other.borderLineWidth)
				&& Util.equal(this.color, other.color)
				&& Util.equal(this.lineType, other.lineType);
	}
	
	/** {@inheritDoc} */
	@Override
	public String toString() {
		return new StringBuilder("Border[width=")
		.append(this.borderLineWidth).append(", color=")
		.append(this.color).append(", type=")
		.append(this.lineType).append("]").toString();
	}
}
