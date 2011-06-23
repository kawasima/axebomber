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
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Sheet {
	@SuppressWarnings("unused")
	private static final Logger logger = LoggerFactory.getLogger(Sheet.class);
	private org.apache.poi.ss.usermodel.Sheet sheet;
	private Integer rowIndex;
	private TableHeader tableHeader;
	private boolean isReadOnly = false;

	public Sheet(org.apache.poi.ss.usermodel.Sheet sheet) {
		rowIndex = 0;
		if(sheet == null)
			throw new RuntimeException("sheet not found");
		this.sheet = sheet;
	}

	public void setTableLabel(String label) {
		Cell labelCell = findCell(label);
		tableHeader = new TableHeader(labelCell, label);
	}

	public void setTableLabel(Pattern p){
		Cell labelCell = findCell(p);
		tableHeader = new TableHeader(labelCell, p);
	}

	public TableHeader getTableHeader() {
		return this.tableHeader;
	}

	public List<Row> getRows() {
		if (tableHeader != null) {
			// デフォルトではタイトル行っぽいものにぶつかったら次のデータの塊として認識する
			final TableHeader currentTableHeader = tableHeader;
			Map<String, Object> options = new HashMap<String, Object>();
			options.put("stopCondition", new ParseCondition() {
				public boolean stop(Row row) {
					return currentTableHeader.match(row);
				}

			});
			return getRows(options);
		} else {
			return getRows(null);
		}
	}

	public List<Row> getRows(Map<String, Object> options) {
		ParseCondition parseCondition = null;
		Boolean exceptGrayout = false;
		Boolean exceptInvisible = false;
		if(options != null) {
			 parseCondition = (ParseCondition)options.get("stopCondition");
			 exceptGrayout = (Boolean)options.get("exceptGrayout");
			 if(exceptGrayout == null) exceptGrayout = false;
			 exceptInvisible = (Boolean)options.get("exceptInvisible");
			 if(exceptInvisible == null) exceptInvisible = false;
		}
		List<Row> rows = new ArrayList<Row>();
		for(int i=tableHeader.getBodyRowIndex(); i<=sheet.getLastRowNum(); i++) {
			Row row = this.getRow(i);

			if(exceptGrayout && row.isGrayout()) {
				continue;
			}
			if(exceptInvisible) {
				continue;
			}

			if(parseCondition != null && parseCondition.stop(row))
				break;
			rows.add(row);
		}
		return rows;
	}

	public Row getRow(int rowIndex) {
		org.apache.poi.ss.usermodel.Row row = sheet.getRow(rowIndex);
		if(row == null)
			row = this.sheet.createRow(rowIndex);
		return new Row(row, tableHeader);
	}

	public Cell findCell(Pattern p) {
		return findCell(p, false);
	}
	public Cell findCell(Pattern p, boolean scanAll) {
		int initialRowNum = (scanAll || tableHeader==null)?sheet.getFirstRowNum() : tableHeader.getBodyRowIndex();
		for(int i=initialRowNum; i<=sheet.getLastRowNum(); i++) {
			org.apache.poi.ss.usermodel.Row row = sheet.getRow(i);
			if(row == null)
					continue;
			for(short j=row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
				org.apache.poi.ss.usermodel.Cell cell = row.getCell(j);
				if (cell == null)
					continue;
				Matcher m = p.matcher(cell.getStringCellValue());
				//logger.debug(cell.getStringCellValue());
				if (m.find()) {
					return new CellImpl(cell);
				}
			}
		}
		throw new CellNotFoundException(p.pattern() + " is not found");
	}

	public Cell findCell(String value) {
		return findCell(value, false);
	}
	public Cell findCell(String value, boolean scanAll) {
		int initialRowNum = (scanAll || tableHeader==null)?sheet.getFirstRowNum() : tableHeader.getBodyRowIndex();
		for(int i=initialRowNum; i<=sheet.getLastRowNum(); i++) {
			org.apache.poi.ss.usermodel.Row row = sheet.getRow(i);
			if(row == null)
					continue;
			for(short j=row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
				org.apache.poi.ss.usermodel.Cell cell = row.getCell(j);
				if (cell == null)
					continue;
				if (value.equals(cell.getStringCellValue())) {
					return new CellImpl(cell);
				}
			}
		}
		throw new CellNotFoundException(value + " is not found");
	}

	public Cell cell(int columnIndex, int rowIndex) {
		org.apache.poi.ss.usermodel.Row row = sheet.getRow(rowIndex);
		if(row == null) {
			row = sheet.createRow(rowIndex);
		}
		org.apache.poi.ss.usermodel.Cell cell = row.getCell(columnIndex);
		if(cell == null)
			cell = row.createCell(columnIndex);
		return new CellImpl(cell);
	}

	/**
	 * Get cell in current row.
	 *
	 * @param columnIndex
	 * @return cell object
	 */
	public Cell cell(int columnIndex) {
		return cell(columnIndex, this.rowIndex);
	}

	public void nextRow() {
		rowIndex++;
		if(!isReadOnly) {
			if(sheet.getRow(rowIndex) == null) {
				sheet.createRow(rowIndex);
			}
		}
	}

	/**
	 *
	 */
	public void autoStyling() {

	}
}
