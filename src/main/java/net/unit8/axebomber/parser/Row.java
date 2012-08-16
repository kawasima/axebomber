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

import java.awt.Color;
import java.util.Map;

public class Row {
	org.apache.poi.ss.usermodel.Row row;
	private Map<String, Integer> labelColumns = null;
	private boolean editable = false;

	public Row(org.apache.poi.ss.usermodel.Row row, TableHeader tableHeader) {
		this.row = row;
		if(tableHeader != null)
			this.labelColumns = tableHeader.getLabelColumns();
	}

	public Cell cell(Integer columnIndex) {
		return cell(columnIndex, editable);
	}

	public Cell cell(Integer columnIndex, boolean create) {
		org.apache.poi.ss.usermodel.Cell cell = row.getCell(columnIndex);
		if(create && cell == null) {
			cell = row.createCell(columnIndex);
		}
		if (cell != null) {
			return new CellImpl(cell);
		} else {
			return new BlankCell(columnIndex, row.getRowNum());
		}
	}
	public Cell cell(String name) {
		if(this.labelColumns == null) {
			throw new RuntimeException("labelColumnsを設定しないとラベルは使えません");
		}
		Integer columnIndex = labelColumns.get(name);
		if(columnIndex == null)
			throw new IllegalArgumentException("can't find label:"+name);
		org.apache.poi.ss.usermodel.Cell cell = row.getCell(columnIndex);
		if (cell != null) {
			return new CellImpl(cell);
		} else {
			return new BlankCell(columnIndex, row.getRowNum());
		}
	}

	public boolean isGrayout() {
		Cell cell = cell((int)row.getFirstCellNum());
		Color color = cell.getColor();
		if(color.getRed() == color.getBlue() && color.getBlue() == color.getGreen()
				&& color.getRed() < 192) {
			return true;
		}
		return false;
	}

	@Override
	public String toString() {
		StringBuilder sb = new StringBuilder(this.row.getSheet().getSheetName() + " at " + this.row.getRowNum());
		sb.append(" [");
		for(int i=this.row.getFirstCellNum(); i<this.row.getLastCellNum(); i++) {
			sb.append(cell(i).toString());
			if(i < this.row.getLastCellNum() - 1)
				sb.append(",");
		}
		sb.append("]");
		return sb.toString();
	}

	public Integer getIndex() {
		return row.getRowNum();
	}
	public org.apache.poi.ss.usermodel.Row getSubstance() {
		return row;
	}

	public boolean isEditable() {
		return editable;
	}

	public void setEditable(boolean editable) {
		this.editable = editable;
	}
}
