require 'rubyXL'

RUI_ASSET_KEYS = ["PT_NEN","PT_SHISAN","PT_KOU_RUI","PT_KOU_SON","PT_KYO_RUI","PT_KYO_SON"]
ASSET_KEYS = ["PT_SHN_SY","PT_SHN_ME","PT_HIRITU","PT_SURYO","PT_TANKA","PT_HYOU_GAK","PT_UKE_GAK","PT_HYOU_SON"]
# SOURCE_XLSX_FILE = "./Book2.xlsx"
SOURCE_XLSX_FILE = "./benefit401k_20150204-20150422.xlsx"
DEST_XLSX_FILE = "./benefit401k.xlsx"

source_book = RubyXL::Parser.parse(SOURCE_XLSX_FILE)
dest_book = RubyXL::Parser.parse(DEST_XLSX_FILE)

source_book.each do |source_sheet|
  data = Array.new(11){Array.new(8)}
  puts "+"*30 + "入力" + "+"*30
  puts "sheet_name = #{source_sheet.sheet_name}"
  puts "source_sheet[36][1].value.gsub(/^（|現在）$/,"") + '時点' = #{source_sheet[36][1].value.gsub(/^（|現在）$/,"") + '時点'}"
  puts "source_sheet[37][2].value = #{source_sheet[37][2].value}"
  puts "source_sheet[38][2].value = #{source_sheet[38][2].value}"
  puts "source_sheet[39][2].value = #{source_sheet[39][2].value}"
  puts "source_sheet[41][2].value = #{source_sheet[41][2].value}"
  puts "source_sheet[42][2].value = #{source_sheet[42][2].value}"

  data[0][0] = source_sheet[36][1].value.gsub(/^（|現在）$/,"") + '時点'
  data[0][1] = source_sheet[37][2].value
  data[0][2] = source_sheet[38][2].value
  data[0][3] = source_sheet[39][2].value
  data[0][4] = source_sheet[41][2].value
  data[0][5] = source_sheet[42][2].value
  here = 0
  ii = 1
  source_sheet.each do |row_data|
    if( 46 <= here ) then
      break if( nil == row_data || 0 == row_data.size )
      puts "here = #{here}; #{row_data}"
      if( 0 == (here % 2) ) then
        if( nil == row_data[2] ) then
          puts row_data.class
          puts row_data.size
          row_data.each do |tmp|
            puts tmp
          end
        end
        puts row_data[2].value
        puts "source_sheet[#{here}  ][1].value = #{source_sheet[here  ][1].value}"
        puts "source_sheet[#{here}  ][2].value = #{source_sheet[here  ][2].value}"
        puts "source_sheet[#{here}  ][5].value = #{source_sheet[here  ][5].value}"
        puts "source_sheet[#{here}+1][1].value = #{source_sheet[here+1][1].value}"
        puts "source_sheet[#{here}+1][2].value = #{source_sheet[here+1][2].value}"
        puts "source_sheet[#{here}+1][3].value = #{source_sheet[here+1][3].value}"
        puts "source_sheet[#{here}+1][4].value = #{source_sheet[here+1][4].value}"
        puts "source_sheet[#{here}+1][5].value = #{source_sheet[here+1][5].value}"
        data[ii][0] = source_sheet[here  ][1].value.rstrip
        if( nil == source_sheet[here  ][2].value ) then
          data[ii][1] = "現金"
        else
          # 全角括弧を半角括弧に変換
          data[ii][1] = source_sheet[here  ][2].value.rstrip.tr("（）","()")
        end
        data[ii][2] = source_sheet[here  ][5].value
        data[ii][3] = source_sheet[here+1][1].value
        data[ii][4] = source_sheet[here+1][2].value
        data[ii][5] = source_sheet[here+1][3].value
        data[ii][6] = source_sheet[here+1][4].value
        data[ii][7] = source_sheet[here+1][5].value
        ii += 1
      end
    end
    # puts here
    here += 1
  end
  puts "ii = #{ii}"
  data[0..(ii-1)].each do |tmp|
    puts tmp.join(",\t")
  end
  
  puts "--#--"*5 + "先頭シートへの入力" + "--#--"*5
  here = 0
  already_inputed = false
  dest_book[0].each do |row_data|
    if( 0 != row_data.size ) then
      if( row_data[0].value == data[0][0] ) then
        already_inputed = true
        break
      end
      here += 1
      next
    else
      break
    end
  end
  puts "here = #{here}; already_inputed = #{already_inputed}"
  for jj in 0..5 do
    dest_book[0].add_cell(here,jj,data[0][jj])
  end
  if( !already_inputed ) then
    sum = 0
    for jj in 1..(ii-1) do
      for column in 0..(ASSET_KEYS.size-1) do
        dest_book[0].add_cell(here,RUI_ASSET_KEYS.size+ASSET_KEYS.size*(jj-1)+column,data[jj][column])
      end
      sum += dest_book[0][here][RUI_ASSET_KEYS.size+ASSET_KEYS.size*(jj-1)+5].value
    end
    puts "dest_book[0][here][1].value = #{dest_book[0][here][1].value}"
    puts "sum                         = #{sum}"
    if( dest_book[0][here][1].value != sum ) then
      puts "#{dest_book[0][here][1].value} != #{sum}"
      puts "error"
      exit
    end
  end
  puts "*----"*5 + "商品別シートへの入力" + "----*"*5
  
  count = 0
  sum = 0
  for jj in 1..(ii-1) do
    already_inputed = false
    dest_book.each do |dest_sheet|
      # if( 6 == jj ) then
        # puts "jj = #{jj}"
        # puts "dest_sheet[0][0].value = #{dest_sheet[0][0].value}"
        # puts "data[jj][1]            = #{data[jj][1]}"
      # end
      if( dest_sheet[0][0].value == data[jj][1] ) then
        puts "jj = #{jj} : sheet_name = #{dest_sheet.sheet_name}"
        puts "dest_sheet[0][0].value = #{dest_sheet[0][0].value}"
        puts "data[jj][1]            = #{data[jj][1]}"
        here = 0
        dest_sheet.each do |row_data|
          if( row_data[0].value == data[0][0] ) then
            puts "#{dest_sheet[here][0].value}のデータは登録済みです。"
            already_inputed = true
            sum += dest_sheet[here][7].value
            break
          end
          here += 1
        end
        break if( already_inputed )
        puts "here = #{here}"
        dest_sheet.add_cell(here,0,data[0][0])
        dest_sheet.add_cell(here,1,data[jj][5].to_f/data[0][1].to_f)
        for column in 0..(ASSET_KEYS.size-1) do
          dest_sheet.add_cell(here,2+column,data[jj][column])
        end
        sum += dest_sheet[here][7].value
        break
      end
    end
    count += 1
  end
  puts "count = #{count}"
  puts "source_sheet[37][2].value = #{source_sheet[37][2].value}"
  puts "sum                       = #{sum}"
  if( source_sheet[37][2].value != sum ) then
    puts "#{source_sheet[37][2].value} != #{sum}"
    puts "error"
    exit
  end
  if( (ii-1) != count ) then
    puts "error"
    exit
  end
end
dest_book.write(DEST_XLSX_FILE)
puts :end