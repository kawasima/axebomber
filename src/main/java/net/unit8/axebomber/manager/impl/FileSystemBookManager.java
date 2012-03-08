package net.unit8.axebomber.manager.impl;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import net.unit8.axebomber.BookIOException;
import net.unit8.axebomber.manager.BookManager;
import net.unit8.axebomber.parser.Book;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileSystemBookManager extends BookManager {

	@Override
	public Book create(String path) {
		Book newbook;

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

	@Override
	public Book open(String path) {
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(path);
			Workbook workbook = WorkbookFactory.create(fis);
			Book book = new Book(workbook);
			book.setPath(path);
			return book;
		} catch(IOException e) {
			throw new BookIOException("", e);
		} catch (InvalidFormatException e) {
			throw new BookIOException("", e);
		} finally {
			IOUtils.closeQuietly(fis);
		}
	}

	@Override
	public void save(Book book) {
		try {
			FileOutputStream fos = new FileOutputStream(book.getPath());
			try {
				book.getSubstance().write(fos);
			} finally {
				IOUtils.closeQuietly(fos);
			}
		} catch(IOException e) {
			throw new BookIOException(e);
		}
	}
}
