package com.github.jferard.spreadsheetwrapper.style;

import java.util.Arrays;

import com.github.jferard.spreadsheetwrapper.Util;

/**
 * The Borders class represents the borders of a cell
 */
public class Borders {
	/**
	 * the top border
	 */
	private Border borderTop;
	/**
	 * the bottom border
	 */
	private Border borderBottom;
	/**
	 * the left border
	 */
	private Border borderLeft;
	/**
	 * the right border
	 */
	private Border borderRight;

	/**
	 * @param borderLineWidth the width to set
	 * @return the Borders objets (fluent style)
	 */
	public Borders setLineWidth(final double borderLineWidth) {
		this.borderTop = (this.borderTop == null ? new Border()
				: this.borderTop).setLineWidth(borderLineWidth);
		this.borderBottom = (this.borderBottom == null ? new Border()
				: this.borderBottom).setLineWidth(borderLineWidth);
		this.borderLeft = (this.borderLeft == null ? new Border()
				: this.borderLeft).setLineWidth(borderLineWidth);
		this.borderRight = (this.borderRight == null ? new Border()
				: this.borderRight).setLineWidth(borderLineWidth);
		return this;
	}

	public double getLineWidth() {
		double lineWidth = this.borderTop == null ? WrapperCellStyle.NO_LINE
				: this.borderTop.getLineWidth();
		for (final Border border : Arrays.asList(this.borderBottom, this.borderLeft,
				this.borderRight)) {
			if (!(border == null || Util.almostEqual(lineWidth,
					border.getLineWidth())))
				return WrapperCellStyle.NO_LINE;
		}
		return lineWidth;
	}

	public WrapperColor getLineColor() {
		WrapperColor color = this.borderTop == null ? null : this.borderTop
				.getLineColor();
		for (Border border : Arrays.asList(this.borderBottom, this.borderLeft,
				this.borderRight)) {
			if (!(border == null || Util.equal(color, border.getLineColor())))
				return null;
		}
		return color;
	}

	public Border getBorderTop() {
		return this.borderTop;
	}

	public Borders setBorderTop(final Border borderTop) {
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

	public Borders setBorderLeft(final Border borderLeft) {
		this.borderLeft = borderLeft;
		return this;
	}

	public Border getBorderRight() {
		return this.borderRight;
	}

	public Borders setBorderRight(final Border borderRight) {
		this.borderRight = borderRight;
		return this;
	}

	public Borders setLineType(final Object lineType) {
		this.borderTop = (this.borderTop == null ? new Border()
				: this.borderTop).setLineType(lineType);
		this.borderBottom = (this.borderBottom == null ? new Border()
				: this.borderBottom).setLineType(lineType);
		this.borderLeft = (this.borderLeft == null ? new Border() : this.borderLeft)
				.setLineType(lineType);
		this.borderRight = (this.borderRight == null ? new Border() : this.borderRight)
				.setLineType(lineType);
		return this;
	}

	public Borders setLineColor(final WrapperColor color) {
		this.borderTop = (this.borderTop == null ? new Border()
				: this.borderTop).setLineColor(color);
		this.borderBottom = (this.borderBottom == null ? new Border()
				: this.borderBottom).setLineColor(color);
		this.borderLeft = (this.borderLeft == null ? new Border() : this.borderLeft)
				.setLineColor(color);
		this.borderRight = (this.borderRight == null ? new Border() : this.borderRight)
				.setLineColor(color);
		return this;
	}

	/** {@inheritDoc} */
	@Override
	public boolean equals(final/*@Nullable*/Object obj) {
		if (this == obj)
			return true;
		if (!(obj instanceof Borders))
			return false;

		final Borders other = (Borders) obj;
		return Util.equal(this.borderBottom, other.borderBottom)
				&& Util.equal(this.borderLeft, other.borderLeft)
				&& Util.equal(this.borderRight, other.borderRight)
				&& Util.equal(this.borderTop, other.borderTop);
	}

	/** {@inheritDoc} */
	@Override
	public String toString() {
		return new StringBuilder("Borders[top=").append(this.borderTop)
				.append(", bottom=").append(this.borderBottom)
				.append(", left=").append(this.borderLeft).append(", right=")
				.append(this.borderRight).append("]").toString();
	}
}
