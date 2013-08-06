# 1-1
# -*- mode:ruby; coding:utf-8 -*-

require 'rubygems'
require 'win32ole'
require 'image_size'



### 画像のファイルサイズを取得する関数.
def file_size(filename)
	open(filename, 'rb') do |f|
	 	img = ImageSize.new(f.read)
	 	return { width: img.get_width, height: img.get_height }
	end
end

### シートに画像を挿入するメソッド.
def insert_picture(filename, sheet)
	size = file_size(filename)
	sheet.Shapes.AddPicture(
		filename,
		false,
		true,
		1,
		1,
		size[:width],
		size[:height])
end

### 対象ディレクトリから画像ファイルのリストを取得する
def files_get
	files = Hash.new
  Dir.glob('./img/*') do |f|
    # support: bmp, gif, jpeg, pbm, pcx, pgm, png, ppm, psd, swf, tiff, xbm, xpm
    if /.*?\.(jpg|jpeg|png)/ =~ f
      if files.has_key? sheet_name(f) then
      	files[sheet_name(f)] << f
      else
      	files[sheet_name(f)] = [f]
      end
    end
  end
  files
end

def sheet_name(filename)
	basename = File.basename(filename, '.*')
	basename.split(".")[0]
end

excel = WIN32OLE.new('Excel.Application')

begin
	book = excel.workbooks.add

	counter = 1

	files_get.each do |key, value|

    # 貼り付けるシートを取得
		book.worksheets.add({ :after =>  book.sheets(book.sheets.count) }) if book.sheets.count < counter
		sheet = book.sheets[counter]

		value.each do |f|
			basename = File.basename(f, '.*')
			sheet.Name = sheet_name basename
		  filepath = File.expand_path(f)
		  insert_picture(filepath, sheet)
		end

    counter += 1

	end

	book.saveAs File.expand_path("./evidence.xls")
ensure
	excel.Workbooks.Close
  excel.quit
end
