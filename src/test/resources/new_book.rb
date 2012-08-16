require 'java'
java_import "net.unit8.axebomber.manager.impl.FileSystemBookManager"
java_import "net.unit8.axebomber.parser.Style"

manager = FileSystemBookManager.new
book = manager.create("target/new_book.xls");
sheet = book.getSheet('test')

row = sheet.getRow(0);
row.cell(0).value = "No."
row.cell(3).value = "title"

range = sheet.createRange("R0")
style = Style.new
style.backgroundColor = "GOLD"
style.borderStyle = "MEDIUM"
range.style = style

manager.save(book)
