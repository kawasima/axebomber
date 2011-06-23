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

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Book {
	private Workbook workbook;
	private String path;

	public Book(Workbook workbook) {
		this.workbook = workbook;
	}

	public Sheet sheet(String name) {
		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet(name);
		if(sheet == null) {
			sheet = workbook.createSheet(name);
		}
		return new Sheet(sheet);
	}

	public org.apache.poi.ss.usermodel.Workbook getSubstance() {
		return this.workbook;
	}
	
	public void save() {
		try {
			FileOutputStream fos = new FileOutputStream(path);
			try {
				workbook.write(fos);
			} finally {
				IOUtils.closeQuietly(fos);
			}
		} catch(IOException e) {
			throw new RuntimeException(e);
		}
	}

	public void setPath(String path) {
		this.path = path;
	}
	
	public static Book create(String path) {
		Book newbook;
		File file = new File(path);
		if(file.exists()) {
			FileInputStream fis = null;
			try {
				 fis = new FileInputStream(file);
				Workbook workbook = WorkbookFactory.create(fis);
				newbook = new Book(workbook);
				newbook.setPath(path);
				return newbook;
			} catch(Exception e) {
				throw new RuntimeException(e);
			} finally {
				IOUtils.closeQuietly(fis);
			}
		}

		String ext = FilenameUtils.getExtension(path);
		if(StringUtils.equals("xls", ext)) {
			newbook = new Book(new HSSFWorkbook());
			newbook.setPath(path);
		} else if(StringUtils.equals("xlsx", ext)){
			newbook = new Book(new XSSFWorkbook());
			newbook.setPath(path);
		} else {
			newbook = new Book(new HSSFWorkbook());
			newbook.setPath(path + ".xls");
		}
		return newbook;
	}
}
