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

public class BlankCell extends Cell {
	public BlankCell(Integer columnIndex, Integer rowIndex) {
		setColumnIndex(columnIndex);
		setRowIndex(rowIndex);
	}

	public Boolean isNotBlank() {
		return false;
	}

	public String toString(){
		return "";
	}

	public int to_i() {
		return 0;
	}
	public Color getColor() {
		return null;
	}
	public void setValue(Object value) {
		// doNothing
		// TODO It's better to throw an Exception?
	}
	public void setColor(Object color) {
		// doNothing
	}
	
	@Override
	public org.apache.poi.ss.usermodel.Cell getSubstance() {
		throw new CellNotFoundException("This blank cell is not substantial");
	}

}
