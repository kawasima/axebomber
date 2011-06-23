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
import java.util.regex.Pattern;

public abstract class Cell {
	private int rowIndex;
	private int columnIndex;

	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}
	public int getRowIndex() {
		return rowIndex;
	}

	public void setColumnIndex(int columnIndex) {
		this.columnIndex = columnIndex;
	}
	public int getColumnIndex() {
		return columnIndex;
	}

	public abstract void setValue(Object value);
	public abstract Color getColor();
	public abstract void setColor(Object color);
	public abstract int to_i();
	
	public static Pattern pattern(String patternStr) {
		return Pattern.compile(patternStr);
	}
	
	public abstract org.apache.poi.ss.usermodel.Cell getSubstance();

}
