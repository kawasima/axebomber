package net.unit8.axebomber.parser;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public class StyleManager {
	public static final int BORDER_TOP    = 1;
	public static final int BORDER_BOTTOM = 3;
	public static final int BORDER_LEFT   = 9;
	public static final int BORDER_RIGHT  = 27;

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
			for (int i=0; i <  3*3*3*3; i++) {
				CellStyle cellStyle = workbook.createCellStyle();
				cellStyle.setFillForegroundColor(style.getBackgroundColor().getIndex());
				cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
				if ((i % 3) != 0) {
					cellStyle.setBorderTop(i % 3==1 ? style.getBorderStyle() : style.getInnerBorderStyle());
				}
				if ((i / 3) % 3 != 0) {
					cellStyle.setBorderBottom((i / 3) % 3 ==1 ? style.getBorderStyle() : style.getInnerBorderStyle());
				}
				if ((i / 9) % 3 != 0) {
					cellStyle.setBorderLeft((i / 9) % 3 ==1 ? style.getBorderStyle() : style.getInnerBorderStyle());
				}
				if ((i / 27) % 3 != 0) {
					cellStyle.setBorderRight((i / 27) % 3 ==1 ? style.getBorderStyle() : style.getInnerBorderStyle());
				}
				cellStyles.add(cellStyle);
			}
			styles.put(style, cellStyles);
		}
		return styles.get(style).get(borderBits);
	}
}
