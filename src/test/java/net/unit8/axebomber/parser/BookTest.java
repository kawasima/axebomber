package net.unit8.axebomber.parser;

import static org.junit.Assert.*;

import java.io.File;
import java.net.URISyntaxException;

import net.unit8.axebomber.manager.BookManager;
import net.unit8.axebomber.manager.impl.ReadOnlyFileSystemBookManager;

import org.junit.Before;
import org.junit.Test;

public class BookTest {
	Book book;
	@Before
	public void prepareBook() throws URISyntaxException {
		BookManager bookManager = new ReadOnlyFileSystemBookManager();
		book = bookManager.open(new File(getClass().getClassLoader().getResource("sample1.xls").toURI()).getAbsolutePath());
	}

	@Test
	public void findCell_can_search_by_a_cell_value() {
		Sheet sheet = book.getSheet("画面仕様書");
		assertNotNull(sheet);

		Cell cell1 = sheet.findCell("画面仕様書");
		assertNotNull(cell1);

		assertEquals(0, cell1.getColumnIndex());
		assertEquals(0, cell1.getRowIndex());

		Cell cell2 = sheet.findCell("承認者");
		assertEquals(0, cell2.getColumnIndex());
		assertEquals(3, cell2.getRowIndex());

	}
}
