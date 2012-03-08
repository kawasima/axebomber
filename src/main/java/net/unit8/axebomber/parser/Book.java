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

import org.apache.poi.ss.usermodel.Workbook;

/**
 * The model of Excel book.
 * 
 * @author kawasima
 *
 */
public class Book {
	private Workbook workbook;
	private String path;

	public Book(Workbook workbook) {
		this.workbook = workbook;
	}

	/**
	 * get sheet object.
	 * 
	 * @param name the sheet name
	 * @return sheet object
	 */
	public Sheet sheet(String name) {
		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet(name);
		if(sheet == null) {
			sheet = workbook.createSheet(name);
		}
		return new Sheet(sheet);
	}

	/**
	 * get POI Workbook object.
	 * 
	 * @return POI workbook object
	 */
	public org.apache.poi.ss.usermodel.Workbook getSubstance() {
		return this.workbook;
	}
	
	public String getPath() {
		return this.path;
	}
	public void setPath(String path) {
		this.path = path;
	}
}