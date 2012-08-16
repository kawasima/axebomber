package net.unit8.axebomber.parser;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public class StyleManager {
	public static final int BORDER_TOP    = 1;
	public static final int BORDER_BOTTOM = 2;
	public static final int BORDER_LEFT   = 4;
	public static final int BORDER_RIGHT  = 8;

	private static volatile Map<Workbook, StyleManager> cache = new HashMap<Workbook, StyleManager>();
	private Workbook workbook;
	private volatile Map<Style, List<CellStyle>> styles = new HashMap<Style, List<CellStyle>>();

	protected StyleManager(Workbook workbook) {
		this.workbook = workbook;
	}

	public static StyleManager getInstance(Workbook book) {
		if (!cache.containsKey(book)) {
			synchronized(StyleManager.class) {
				if (!cache.containsKey(book)) {
					cache.put(book, new StyleManager(book));
				}
			}
		}
		return cache.get(book);

	}

	public CellStyle getStyle(Style style, int borderBits) {
		if (!styles.containsKey(style)) {
			List<CellStyle> cellStyles = new ArrayList<CellStyle>();
			for (int i=0; i <= (BORDER_TOP|BORDER_BOTTOM|BORDER_LEFT|BORDER_RIGHT); i++) {
				CellStyle cellStyle = workbook.createCellStyle();
				cellStyle.setFillForegroundColor(style.getBackgroundColor().getIndex());
				cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
				if ((i & BORDER_TOP) != 0) {
					cellStyle.setBorderTop(style.getBorderStyle());
				}
				if ((i & BORDER_BOTTOM) != 0) {
					cellStyle.setBorderBottom(style.getBorderStyle());
				}
				if ((i & BORDER_LEFT) != 0) {
					cellStyle.setBorderLeft(style.getBorderStyle());
				}
				if ((i & BORDER_RIGHT) != 0) {
					cellStyle.setBorderRight(style.getBorderStyle());
				}
				cellStyles.add(cellStyle);
			}
			styles.put(style, cellStyles);
		}
		return styles.get(style).get(borderBits);
	}
}
