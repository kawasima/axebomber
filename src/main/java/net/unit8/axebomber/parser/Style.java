package net.unit8.axebomber.parser;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

public class Style {
	private IndexedColors color;
	private IndexedColors backgroundColor;
	private short borderStyle;
	private IndexedColors borderColor;

	@Override
	public boolean equals(Object obj) {
		if (!(obj instanceof Style))
			return false;
		Style style = (Style)obj;
		return style.getColor() == this.color
				&& style.getBackgroundColor() == this.backgroundColor
				&& style.getBorderStyle() == this.borderStyle
				&& style.getBorderColor() == this.borderColor;
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
}
