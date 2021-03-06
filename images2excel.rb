# フォルダ内の画像をExcelに貼り付けるrubyスクリプト
# -*- mode:ruby; coding:utf-8 -*-

require 'rubygems'
require 'win32ole'
require 'image_size'
require 'yaml'
require 'csv'


class Images2Excel

	CELL_WIDTH = 8.38
	CELL_HEIGTH = 13.5

	def initialize(width=0.65, height=0.65, space=50)
		@sheet_counter = 1
		@scale_width = width
		@scale_height = height
		@space = space
	end

  
	def convert(book, dir)
		files = get_files_hash(dir)
		sheetnames = files.keys.sort_by {|x| x.split("-")[0].to_i }
		sheetnames.each do |sheetname|

			# 貼り付けるシートを取得
			sheet = get_next_sheet(book)

			files[sheetname].each do |f|
				basename = File.basename(f, '.*')
				sheet.Name = get_sheet_name basename
				insert_comment(basename, sheet) if get_comment basename
				filepath = File.expand_path(f)
				
				if csv_file?(File.basename(f))
					insert_csv(filepath, sheet)
				else
					insert_picture(filepath, sheet)
				end
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
		 	return { width: img.width, height: img.height }
		end
	end

  ### シートにコメントを挿入するメソッド
  # コメントの分だけ@yを増加させる
	def insert_comment(filename, sheet)
		celRow = (@y /  CELL_HEIGTH).ceil
		c_num = 0
		comment = get_comment filename
		comment.split('_').each do |c|
			cell = sheet.Cells.Item(celRow + c_num, 1)
			cell.Value = c
			c_num += 1
		end
		
		@y = ((@y /  CELL_HEIGTH).ceil + c_num - 1) * CELL_HEIGTH
	end

	### シートに画像を挿入するメソッド.
	def insert_picture(filename, sheet)
		size = file_size(filename)
		sheet.Shapes.AddPicture(
			filename.gsub(/\//, '\\'),
			false,
			true,
			1,
			@y, #(CELL_HEIGTH * ln_num).ceil,
			size[:width] * @scale_width,
			size[:height] * @scale_height)
		@y += size[:height] * @scale_height + @space
	end


	### シートに画像を挿入するメソッド.
	def insert_csv(filename, sheet)
		celRow = (@y /  CELL_HEIGTH).ceil
		
		arrs = CSV.read(filename)
		arrs.each_with_index do |line, l_num|
		
			line.each_with_index do |value, index|
				cell = sheet.Cells.Item(celRow + l_num + 1, index + 1)
				cell.Value = value
			end
		end
		
		@y = @y + (arrs.size) * CELL_HEIGTH  + @space
	end

	### 対象ディレクトリから画像ファイルのリストを取得する
	def get_files_hash(dir)
		files = Hash.new
		Dir.glob(File.basename(dir) + '/*') do |f|
			# support: bmp, gif, jpeg, png, csv
			if /.*?\.(bmp|BMP|jpg|jpeg|JPG|png|PNG|csv|CSV)/ =~ f
				if files.has_key? get_sheet_name(f) then
					files[get_sheet_name(f)] << f
				else
					files[get_sheet_name(f)] = [f]
				end
			end
		end
		files
	end
	
	def csv_file?(filename)
		/.*?\.(csv|CSV)/ =~ filename
	end

	# ファイル名からシート名を取得
	def get_sheet_name(filename)
		basename = File.basename(filename, '.*')
		basename.split(".")[0]
	end

	# ファイル名からコメントを取得
	def get_comment(filename)
		/\.\_(.*)\.?/ =~ filename
		$1
	end

end


#yaml形式のconfig.ymlファイルを読み込む
config = YAML.load_file("config.yml")['config']

excel = WIN32OLE.new('Excel.Application')

begin
	book = excel.workbooks.add

	i2e = Images2Excel.new(config['scale']['width'], config['scale']['height'], config['space'])
	i2e.convert(book, config['img_dir'])

	book.saveAs File.expand_path(config['export']).gsub(/\//, '\\')
	excel.Workbooks.Close
	excel.quit
ensure
	excel.Workbooks.Close
	excel.quit
end
