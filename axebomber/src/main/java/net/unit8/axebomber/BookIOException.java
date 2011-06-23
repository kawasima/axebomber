package net.unit8.axebomber;

public class BookIOException extends RuntimeException {
	private static final long serialVersionUID = -4245737192442513340L;
	
	public BookIOException(String message) {
		super(message);
	}
	
	public BookIOException(Throwable throwable) {
		super(throwable);
	}
	
	public BookIOException(String message, Throwable throwable) {
		super(message, throwable);
	}
}
