package net.unit8.axebomber.parser;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

public class Range {
	private static final Pattern R1C1 = Pattern.compile("(R(\\d+)?)?(C(\\d+)?)?");

	private int x1, x2, y1, y2;
	private Sheet sheet;
	private List<Integer> labelIndexes;
	private int regionIndex;

	protected Range() {}

	public Range(Cell beginCell, Cell endCell) {

	}

	public static Range parse(String rangeStr, Sheet sheet) {
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
		range.sheet = sheet;
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
		int top = Math.max(y1, sheet.getSubstance().getFirstRowNum());
		int bottom = Math.min(y2, sheet.getSubstance().getLastRowNum());
		for (int i = top; i <= bottom; i++) {
			Row row = sheet.getRow(i);
			int left  = Math.max(x1, row.getSubstance().getFirstCellNum());
			int right = Math.min(x2, row.getSubstance().getLastCellNum());

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

				if (j == left) {
					borderBits += StyleManager.BORDER_LEFT;
				} else if (labelIndexes.contains(j)) {
						borderBits += (StyleManager.BORDER_LEFT * 2);
				}

				if (j == right)
					borderBits += StyleManager.BORDER_RIGHT;

				Cell cell = row.cell(j);
				CellStyle cellStyle = StyleManager.getInstance(sheet.getSubstance().getWorkbook()).getStyle(style, borderBits);
				cell.getSubstance().setCellStyle(cellStyle);
			}
		}
	}

	public Range gridize() {
		int top = Math.max(y1, sheet.getSubstance().getFirstRowNum());
		Row row = sheet.getRow(top);
		int left  = Math.max(x1, row.getSubstance().getFirstCellNum());
		int right = Math.min(x2, row.getSubstance().getLastCellNum());
		for (int i=left; i <= right; i++) {
			sheet.getSubstance().setColumnWidth(i, 746);
		}
		return this;
	}

	public Range merge() {
		regionIndex = sheet.getSubstance().addMergedRegion(new CellRangeAddress(y1, y2, x1, x2));
		return this;
	}

	public Range unmerge() {
		sheet.getSubstance().removeMergedRegion(regionIndex);
		return this;
	}

	public void setLabelColumns(Map<String, Integer> labelColumns) {
		this.labelIndexes = new ArrayList<Integer>(labelColumns.values());
		Collections.sort(labelIndexes);
		int lastIndex = labelIndexes.get(labelIndexes.size() - 1);
		if (x2 < lastIndex) {
			x2 = lastIndex;
		}
	}

	public void setValue(Object value) {
		sheet.cell(x1, y1).setValue(value);
	}

	public int to_i() {
		return sheet.cell(x1,  y1).to_i();
	}

	@Override
	public String toString() {
		return sheet.cell(x1, y1).toString();
	}
}
