require 'rubyXL'


RUI_ASSET_KEYS = ["PT_NEN","PT_SHISAN","PT_KOU_RUI","PT_KOU_SON","PT_KYO_RUI","PT_KYO_SON"]
ASSET_KEYS = ["PT_SHN_SY","PT_SHN_ME","PT_HIRITU","PT_SURYO","PT_TANKA","PT_HYOU_GAK","PT_UKE_GAK","PT_HYOU_SON"]
XLSX_FILE = "./benefit401k.xlsx"


puts <<EOS




     エクセルファイル オープン




EOS
workbook = RubyXL::Parser.parse(XLSX_FILE)
worksheet = workbook[0]

data = Array.new(11){Array.new(8)}

row=0

worksheet.each do |row_data|
  if( 0 == row ) then
    row += 1
    next
  end
  for column in 0..(RUI_ASSET_KEYS.size-1) do
    data[0][column] = row_data[column].value
    # puts data[0][column]
  end
  ii=1
  while( nil != row_data[RUI_ASSET_KEYS.size+ASSET_KEYS.size*(ii-1)] ) do
    for column in 0..(ASSET_KEYS.size-1) do
      data[ii][column] = row_data[RUI_ASSET_KEYS.size+ASSET_KEYS.size*(ii-1)+column].value
      # puts data[ii][column]
    end
    ii += 1
  end
  puts "------------------------------"
  puts "row = #{row}, ii = #{ii}"
  puts row_data[0].value
  # puts data
  count = 0
  sum=0
  for jj in 1..ii-1 do
    already_inputed = false
    workbook.each do |sheet|
      if( sheet[0][0].value == data[jj][1] ) then
        # puts "jj = #{jj}"
        # puts "sheet_name = #{sheet.sheet_name}"
        puts "sheet[0][0].value = #{sheet[0][0].value}"
        # puts "data[jj][1] = #{data[jj][1]}"
        here = 0
        sheet.each do |tmp|
          if( tmp[0].value == data[0][0] ) then
            puts "#{sheet[here][0].value}のデータは登録済みです。"
            already_inputed = true
            break
          end
          here += 1
        end
        break if( already_inputed )
        puts "here = #{here}"
        sheet.add_cell(here,0,data[0][0])
        sheet.add_cell(here,1,data[jj][5].to_f/data[0][1].to_f)
        for column in 0..(ASSET_KEYS.size-1) do
          sheet.add_cell(here,2+column,data[jj][column])
        end
        sum += sheet[here][7].value.to_i
        break
      end
    end
    count += 1
  end
  if( !already_inputed && ( worksheet[row][1].value.to_i != sum ) ) then
    puts "error"
    puts "worksheet[row][1].value.to_i = #{worksheet[row][1].value.to_i}"
    puts "sum = #{sum}"
  end
  puts "count = #{count}"
  if( (ii-1) != count ) then
    puts "error"
  end
  row += 1
end

workbook.write(XLSX_FILE)

puts :end