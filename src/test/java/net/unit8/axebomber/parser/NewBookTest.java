package net.unit8.axebomber.parser;

import java.io.FileReader;
import java.io.IOException;

import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;

import org.apache.commons.io.IOUtils;
import org.junit.Test;


public class NewBookTest {
	@Test
	public void test() throws IOException, ScriptException {
		ScriptEngine jruby = new ScriptEngineManager().getEngineByName("jruby");
		FileReader reader = new FileReader("src/test/resources/new_book.rb");
		try {
			jruby.eval(reader);
		} finally {
			IOUtils.closeQuietly(reader);
		}
	}
}
