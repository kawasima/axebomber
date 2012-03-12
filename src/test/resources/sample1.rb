require 'java'
java_import "net.unit8.axebomber.manager.impl.ReadOnlyFileSystemBookManager"
java_import "net.unit8.axebomber.parser.SheetNameFilter"

class NameFilter
  include SheetNameFilter
  def accept(name)
    /^画面/ === name
  end
end

manager = ReadOnlyFileSystemBookManager.new
book = manager.open("src/test/resources/sample1.xls");
book.sheets('name' => NameFilter.new).each {|sheet|
	puts sheet.name
}
