package net.unit8.axebomber.parser;

import static org.junit.Assert.*;

import org.junit.Test;

public class RangeTest extends Range{

	@Test
	public void test() {
		Range range = Range.parse("R2C2:R3C3");
		System.out.println(range);
		range = Range.parse("R2C2");
		System.out.println(range);
		range = Range.parse("RC2");
		System.out.println(range);
		range = Range.parse("R2");
		System.out.println(range);
	}

}
