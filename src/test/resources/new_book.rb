require 'java'
java_import "net.unit8.axebomber.manager.impl.FileSystemBookManager"
java_import "net.unit8.axebomber.parser.Style"

manager = FileSystemBookManager.new
book = manager.create("target/new_book.xls");
sheet = book.getSheet('test')

row = sheet.getRow(0);
sheet.createRange("RC:RC2").merge.value = "No."
sheet.createRange("RC3:RC10").merge.value = "title"
sheet.titleRowIndex = 0

range = sheet.createRange("R0")
style = Style.new
style.backgroundColor = "GOLD"
style.borderStyle = "MEDIUM"
style.innerBorderStyle = "DASHED"
range.style = style
range.gridize

(1..3).each {|i|
	sheet.cell(0).value = i
	sheet.cell("title").value = "#{i}のテスト"
	sheet.nextRow
}

range = sheet.createRange("R1:R3")
style = Style.new
style.borderStyle = "MEDIUM"
style.innerBorderStyle = "DASHED"
range.style = style

manager.save(book)
