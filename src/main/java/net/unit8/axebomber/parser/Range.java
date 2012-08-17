package net.unit8.axebomber.parser;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

public class Range {
	private static final Pattern R1C1 = Pattern.compile("(R(\\d+)?)?(C(\\d+)?)?");

	private int x1, x2, y1, y2;
	private org.apache.poi.ss.usermodel.Sheet sheet;

	protected Range() {}

	public Range(Cell beginCell, Cell endCell) {

	}

	public static Range parse(String rangeStr) {
		String[] ranges = StringUtils.split(rangeStr, ":", 2);
		Range range = new Range();
		if (ranges.length == 1) {
			Matcher m = R1C1.matcher(ranges[0]);
			if (m.find()) {
				range.y1 = parseIndex(m, 1, -1);
				range.y2 = parseIndex(m, 1, Integer.MAX_VALUE);
				range.x1 = parseIndex(m, 3, -1);
				range.x2 = parseIndex(m, 3, Integer.MAX_VALUE);
			} else
				throw new IllegalArgumentException(rangeStr);
		} else if (ranges.length == 2) {
			Matcher m = R1C1.matcher(ranges[0]);
			if (m.find()) {
				range.y1 = parseIndex(m, 1, -1);
				range.x1 = parseIndex(m, 3, -1);
			}
			m = R1C1.matcher(ranges[1]);
			if (m.find()) {
				range.y2 = parseIndex(m, 1, Integer.MAX_VALUE);
				range.x2 = parseIndex(m, 3, Integer.MAX_VALUE);
			}
		} else {
			throw new IllegalArgumentException(rangeStr);
		}

		if (range.y1 > range.y2) {
			int tmp = range.y2;
			range.y2 = range.y1;
			range.y1 = tmp;
		}
		if (range.x1 > range.x2) {
			int tmp = range.x2;
			range.x2 = range.x1;
			range.x1 = tmp;
		}

		return range;
	}

	private static int parseIndex(Matcher m, int i, int defaultValue) {
		if (StringUtils.isEmpty(m.group(i))) {
			return defaultValue;
		} else {
			return StringUtils.isEmpty(m.group(i+1)) ? 0 : Integer.parseInt(m.group(i+1));
		}
	}


	public void setStyle(Style style) {
		int top = Math.max(y1, sheet.getFirstRowNum());
		int bottom = Math.min(y2, sheet.getLastRowNum());
		for (int i = top; i <= bottom; i++) {
			Row row = sheet.getRow(i);
			int left  = Math.max(x1, row.getFirstCellNum());
			int right = Math.min(x2, row.getLastCellNum());

			for(int j = left; j <= right; j++) {
				int borderBits = 0;
				if (i == top)
					borderBits += StyleManager.BORDER_TOP;
				else
					borderBits += (StyleManager.BORDER_TOP * 2);

				if (i == bottom)
					borderBits += StyleManager.BORDER_BOTTOM;
				else
					borderBits += (StyleManager.BORDER_BOTTOM * 2);

				if (j == left)
					borderBits += StyleManager.BORDER_LEFT;
				else
					borderBits += (StyleManager.BORDER_LEFT * 2);

				if (j == right)
					borderBits += StyleManager.BORDER_RIGHT;
				else
					borderBits += (StyleManager.BORDER_RIGHT * 2);

				org.apache.poi.ss.usermodel.Cell cell = row.getCell(j, Row.CREATE_NULL_AS_BLANK);
				CellStyle cellStyle = StyleManager.getInstance(sheet.getWorkbook()).getStyle(style, borderBits);
				cell.setCellStyle(cellStyle);
			}
		}
	}
	@Override
	public String toString() {
		return String.format("(%d,%d) to (%d,%d)", x1, y1, x2, y2);
	}

	public void setSheet(org.apache.poi.ss.usermodel.Sheet sheet) {
		this.sheet = sheet;
	}
}
