package net.unit8.axebomber.parser;

@SuppressWarnings("serial")
public class SheetNotFoundException extends RuntimeException {

	public SheetNotFoundException(String name) {
		super(name);
	}

}
