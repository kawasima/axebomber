package net.unit8.axebomber.manager.impl;

import java.io.FileInputStream;
import java.io.IOException;

import net.unit8.axebomber.BookIOException;
import net.unit8.axebomber.manager.BookManager;
import net.unit8.axebomber.parser.Book;

import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadOnlyFileSystemBookManager extends BookManager {

	@Override
	public Book create(String path) {
		throw new BookIOException("Can't create a new book.");
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
		// do nothing
	}
}
