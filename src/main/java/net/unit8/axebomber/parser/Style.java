package net.unit8.axebomber.parser;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

public class Style {
	private IndexedColors color;
	private IndexedColors backgroundColor;
	private short borderStyle;
	private IndexedColors borderColor;
	private short innerBorderStyle;

	@Override
	public boolean equals(Object obj) {
		if (!(obj instanceof Style))
			return false;
		Style style = (Style)obj;
		return style.getColor() == this.color
				&& style.getBackgroundColor() == this.backgroundColor
				&& style.getBorderStyle() == this.borderStyle
				&& style.getBorderColor() == this.borderColor
				&& style.getInnerBorderStyle() == this.innerBorderStyle;
	}

	public IndexedColors getColor() {
		return color;
	}

	public void setColor(IndexedColors color) {
		this.color = color;
	}

	public IndexedColors getBackgroundColor() {
		return backgroundColor;
	}

	public void setBackgroundColor(IndexedColors backgroundColor) {
		this.backgroundColor = backgroundColor;
	}

	public void setBackgroundColor(String name) {
		this.backgroundColor = IndexedColors.valueOf(name);
	}

	public short getBorderStyle() {
		return borderStyle;
	}

	public void setBorderStyle(short borderStyle) {
		this.borderStyle = borderStyle;
	}
	public void setBorderStyle(String name) {
		BorderStyle borderStyle = BorderStyle.valueOf(name);
		this.borderStyle = (short) borderStyle.ordinal();
	}

	public IndexedColors getBorderColor() {
		return borderColor;
	}

	public void setBorderColor(IndexedColors borderColor) {
		this.borderColor = borderColor;
	}

	public short getInnerBorderStyle() {
		return innerBorderStyle;
	}

	public void setInnerBorderStyle(short innerBorderStyle) {
		this.innerBorderStyle = innerBorderStyle;
	}

	public void setInnerBorderStyle(String name) {
		BorderStyle borderStyle = BorderStyle.valueOf(name);
		this.innerBorderStyle = (short) borderStyle.ordinal();
	}
}
