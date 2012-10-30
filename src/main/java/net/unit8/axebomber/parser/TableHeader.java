/*******************************************************************************
 * Copyright 2011 kawasima
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 ******************************************************************************/
package net.unit8.axebomber.parser;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;

public class TableHeader {
	private int labelRowIndex;
	private int labelColumnIndex;
	private Map<String, Integer> labelColumns  = new HashMap<String, Integer>();
	private Map<Integer, String> columnLabels  = new HashMap<Integer, String>();
	private Map<String, List<String>> labelRelations = new HashMap<String, List<String>>();
	private Sheet sheet;
	private Object labelPattern;

	public TableHeader(Cell labelCell, Object labelPattern) {
		this.labelPattern = labelPattern;
		this.labelRowIndex = labelCell.getRowIndex() + 1;
		this.labelColumnIndex = labelCell.getColumnIndex();
		this.sheet = labelCell.getSubstance().getSheet();
		Cell beginCell = getCell(labelCell.getColumnIndex(), labelRowIndex);
		scanColumnLabel(beginCell);
	}

	public TableHeader(Row labelRow) {
		this.labelRowIndex = labelRow.getIndex();
		this.labelColumnIndex = 0;
		this.sheet = labelRow.getSubstance().getSheet();
		Cell beginCell = getCell(labelColumnIndex, labelRowIndex);
		for (labelColumnIndex++; beginCell == null && labelColumnIndex <= labelRow.getSubstance().getLastCellNum(); labelColumnIndex++)
			beginCell = getCell(labelColumnIndex, labelRowIndex);

		scanColumnLabel(beginCell);

	}

	public TableHeader(Range range) {
		this.sheet = range.getSheet().getSubstance();
		for (int rowIndex = range.getFirstRowNum(); rowIndex <= range.getLastRowNum(); rowIndex++) {
			String currentLabel = null;
			int colLastIndex = Math.min(this.sheet.getRow(rowIndex).getLastCellNum(), range.getLastColumnNum());
			int colFirstIndex = Math.max(this.sheet.getRow(rowIndex).getFirstCellNum(), range.getFirstColumnNum());
			for (int colIndex = colFirstIndex; colIndex <= colLastIndex; colIndex++) {
				Cell cell = range.getSheet().cell(colIndex, rowIndex);
				String value = StringUtils.remove(StringUtils.trim(cell.toString()), "\n");
				if (StringUtils.isNotBlank(value)) {
					currentLabel = value;
					labelColumns.put(currentLabel, colIndex);
				}
				if (StringUtils.isNotBlank(currentLabel)) {
					String parentLabel = columnLabels.get(colIndex);
					if (parentLabel != null) {
						List<String> children = labelRelations.get(parentLabel);
						if (children == null) {
							children = new ArrayList<String>();
							labelRelations.put(parentLabel, children);
						}
						if (!children.contains(currentLabel))
							children.add(currentLabel);
					}
					columnLabels.put(colIndex, currentLabel);
				}
			}
		}
		this.labelRowIndex = range.getLastRowNum();
	}

	public List<String> childLabels(String parent) {
		return labelRelations.get(parent);
	}
	private void scanColumnLabel(Cell beginCell) {
		int rowIndex = beginCell.getRowIndex();
		org.apache.poi.ss.usermodel.Row row = sheet.getRow(rowIndex);
		String currentValue = null;
		for (int i=row.getFirstCellNum(); i<row.getLastCellNum(); i++) {
			Cell cell = getCell(i, rowIndex);
			if (cell == null)
				continue;
			String label = StringUtils.remove(StringUtils.trim(cell.toString()), "\n");
			if (!cell.toString().equals("")) {
				labelColumns.put(label, i);
				currentValue = label;
			}
			if (currentValue != null)
				columnLabels.put(i, currentValue);
		}
	}

	private Cell getCell(int columnIndex, int rowIndex) {
		org.apache.poi.ss.usermodel.Row row = sheet.getRow(rowIndex);
		if(row == null) {
			return null;
		}
		org.apache.poi.ss.usermodel.Cell cell = row.getCell(columnIndex);
		if(cell == null)
			return null;
		return new CellImpl(cell);
	}
	public Map<String, Integer> getLabelColumns() {
		return labelColumns;
	}

	public boolean match(Row row) {
		if(labelPattern instanceof String) {
			return (StringUtils.equals((String)labelPattern, row.cell(labelColumnIndex).toString()));
		} else if(labelPattern instanceof Pattern) {
			return ((Pattern)labelPattern).matcher(row.cell(labelColumnIndex).toString()).find();
		}
		return false;
	}

	public int getBodyRowIndex() {
		return labelRowIndex+1;
	}

	public String getLabel() {
		return this.toString();
	}

	public String find(int columnIndex) {
		return columnLabels.get(columnIndex);
	}
	@Override
	public String toString() {
		return getCell(labelColumnIndex, labelRowIndex).toString();
	}
}
