[axebomber](https://github.com/kawasima/axebomber)
========

Mainly in Japan, Excel workbook is used as a grid paper in business scene. So we need more flexible and intelligent parser. only POI or OLE is not enough.
Axebomber makes you relieved from stuggling against "hogan" Excel.

INSTALL (using with jruby)
========

    % mvn package
    % cd target
    % jruby -S gem install axebomber-0.1.0.gem

USAGE (using with jruby)
========

    require 'axebomber'
    include Axebomber

    bookManager = ReadOnlyFileSystemBookManager.new
    book = bookManager.open("copy~issues_20120301.xlsx")
    sheet = book.sheet("Sheet1")



EXAMPLE
========

It's an example reading rows except grayed out.

    sheet.rows({
      'exceptGrayout'=>true
    }).each{|row|
      puts row.cell("No").to_i
      puts row.cell("name")
    }