require 'java'
java_import "net.unit8.axebomber.manager.impl.FileSystemBookManager"
java_import "net.unit8.axebomber.parser.Style"

manager = FileSystemBookManager.new
book = manager.create("target/new_book.xls");
sheet = book.getSheet('test')

row = sheet.getRow(0);
row.cell(0).value = "No."
row.cell(3).value = "title"
sheet.titleRowIndex = 0

range = sheet.createRange("R0")
style = Style.new
style.backgroundColor = "GOLD"
style.borderStyle = "MEDIUM"
range.style = style

(1..3).each {|i|
	row = sheet.getRow(i)
	row.cell(0).value = i
	row.cell("title").value = "#{i}のテスト"
}
manager.save(book)
