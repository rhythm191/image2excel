# フォルダ内の画像をExcelに貼り付けるrubyスクリプト
# -*- mode:ruby; coding:utf-8 -*-

require 'rubygems'
require 'win32ole'
require 'image_size'



class Images2Excel

	IMAGE_SPCE = 50

	def initialize
		@sheet_counter = 1
	end

  
	def convert(book, dir)
		get_images_hash(dir).each do |key, value|

	    # 貼り付けるシートを取得
			sheet = get_next_sheet(book)

			value.each do |f|
				basename = File.basename(f, '.*')
				sheet.Name = sheet_name basename
			  filepath = File.expand_path(f)
			  insert_picture(filepath, sheet)
			end
		end
	end

  # 次のシートを取得するメソッド.
	def get_next_sheet(book)
		book.worksheets.add({ :after =>  book.sheets(book.sheets.count) }) if book.sheets.count < @sheet_counter
		sheet = book.sheets[@sheet_counter]
		@sheet_counter += 1
		@y = 1

		sheet
	end

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
			@y,
			size[:width],
			size[:height])
		@y += size[:height] + IMAGE_SPCE
	end

	### 対象ディレクトリから画像ファイルのリストを取得する
	def get_images_hash(dir)
		files = Hash.new
	  Dir.glob(File.basename(dir) + '/*') do |f|
	    # support: bmp, gif, jpeg, pbm, pcx, pgm, png, ppm, psd, swf, tiff, xbm, xpm
	    if /.*?\.(jpg|jpeg|png|JPG|PNG)/ =~ f
	      if files.has_key? sheet_name(f) then
	      	files[sheet_name(f)] << f
	      else
	      	files[sheet_name(f)] = [f]
	      end
	    end
	  end
	  files
	end

  # ファイル名からシート名を取得
	def sheet_name(filename)
		basename = File.basename(filename, '.*')
		basename.split(".")[0]
	end

end


excel = WIN32OLE.new('Excel.Application')

begin
	book = excel.workbooks.add

  i2e = Images2Excel.new
  i2e.convert(book, "./img")

	book.saveAs File.expand_path("./export.xls")
ensure
	excel.Workbooks.Close
  excel.quit
end
