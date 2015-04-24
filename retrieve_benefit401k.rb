require 'mechanize'
require 'logger'
require 'optparse'
require 'rubyXL'
require 'active_support'
require 'active_support/core_ext'

RUI_ASSET_KEYS = ["PT_NEN","PT_SHISAN","PT_KOU_RUI","PT_KOU_SON","PT_KYO_RUI","PT_KYO_SON"]
ASSET_KEYS = ["PT_SHN_SY","PT_SHN_ME","PT_HIRITU","PT_SURYO","PT_TANKA","PT_HYOU_GAK","PT_UKE_GAK","PT_HYOU_SON"]
XLSX_FILE = "./benefit401k.xlsx"

def acquire_asset(page,data,ii,inputName,keyValues)
  page.search('input[@name='+inputName+']').each do | asset|
    jj=0
    keyValues.each do |target|
      case target
      when "PT_SHN_SY","PT_SHN_ME" then
        data[ii][jj] = asset.parent.search('input[@name='+target+']').map{|value| value["value"]}[0]
      when "PT_HIRITU" then
        data[ii][jj] = asset.parent.search('input[@name='+target+']').map{|value| value["value"]}[0].to_f
      when "PT_SURYO","PT_TANKA","PT_HYOU_GAK","PT_UKE_GAK","PT_HYOU_SON" then
        data[ii][jj] = asset.parent.search('input[@name='+target+']').map{|value| value["value"]}[0].delete(",").to_i
      when "PT_NEN" then
        data[ii][jj] = asset.parent.search('input[@name='+target+']').map{|value| value["value"]}[0]
      when "PT_KYO_RUI","PT_SHISAN","PT_KOU_RUI","PT_KYO_SON","PT_KOU_SON" then
        data[ii][jj] = asset.parent.search('input[@name='+target+']').map{|value| value["value"]}[0].delete(",").to_i
      else
        STDERR.puts "error:target = " + target
        data[ii][jj] = nil
        exit
      end
      jj+=1
    end
    ii+=1
  end
  return ii
end

puts <<EOS




     エクセルファイル オープン




EOS
workbook = RubyXL::Parser.parse(XLSX_FILE)
worksheet = workbook[0]
yesterday = Date.yesterday
row=0
for rowData in worksheet do
  puts "row = " + row.to_s + ":" + rowData[0].value.to_s
  if( "日付" == rowData[0].value.to_s ) then
    row += 1
    next
  end
  date = Date.strptime(rowData[0].value.to_s,'%Y年%m月%d日')
  if( yesterday == date ) then
    puts "既に" + yesterday.to_s + "のデータは登録済みです。"
    exit
  end
  break if( nil == rowData[0].value )
  row += 1
end
puts "new row = " + row.to_s

params = ARGV.getopts('',"account:./account.txt","xlsx:./benefit401k.xlsx")
accountFile = open(params["account"].to_s, "r")
if( !accountFile ) then
  puts "アカウント情報ファイルが開けませんでした。"
  exit
end
UserId = accountFile.gets.chomp
PassWdS = accountFile.gets.chomp
accountFile.close

agent = Mechanize.new
agent.log = Logger.new $stderr
OpenSSL::debug = true

puts "agent.user_agent = " + agent.user_agent
puts "agent.verify_mode = " + agent.verify_mode.to_s
puts "agent.ssl_version = " + agent.ssl_version.to_s
puts "agent.ca_file = " + agent.ca_file.to_s
puts "OpenSSL::X509::DEFAULT_CERT_DIR = " + OpenSSL::X509::DEFAULT_CERT_DIR
puts "OpenSSL::X509::DEFAULT_CERT_FILE = " + OpenSSL::X509::DEFAULT_CERT_FILE
puts "agent.cert = " + agent.cert.to_s
puts "agent.key = " + agent.key.to_s
puts "agent.pass = " + agent.pass.to_s
puts "agent.verify_callback = " + agent.verify_callback.to_s
puts "agent.cookies = " + agent.cookies.to_s

agent.user_agent = 'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0)'
agent.verify_mode = OpenSSL::SSL::VERIFY_PEER
agent.ssl_version = 'TLSv1'
# agent.ca_file = './GTEGlRoot.txt'.to_s
agent.ca_file = './BaltimoreCyberTrustRoot.crt'

url = 'https://www.benefit401k.com/customer/'

puts '設定完了'
puts "agent.user_agent = " + agent.user_agent
puts "agent.verify_mode = " + agent.verify_mode.to_s
puts "agent.ssl_version = " + agent.ssl_version.to_s
puts "agent.ca_file = " + agent.ca_file.to_s
puts "OpenSSL::X509::DEFAULT_CERT_DIR = " + OpenSSL::X509::DEFAULT_CERT_DIR
puts "OpenSSL::X509::DEFAULT_CERT_FILE = " + OpenSSL::X509::DEFAULT_CERT_FILE
puts "agent.cert = " + agent.cert.to_s
puts "agent.key = " + agent.key.to_s
puts "agent.pass = " + agent.pass.to_s
puts "agent.verify_callback = " + agent.verify_callback.to_s
puts "agent.cookies = " + agent.cookies.to_s
puts "url = " + url.to_s

loginPage = agent.get(url)

puts "loginPage.title = " + loginPage.title.to_s
puts "agent.cookies = " + agent.cookies.to_s
homePage = loginPage.form_with(:name => 'signonform') do |form|
  form.UserId = UserId
  form.PassWdS = PassWdS
end.submit
puts "agent.cookies = " + agent.cookies.to_s
puts "homePage.title = " + homePage.title.to_s
# もしログインエラーなら、homePage.title = ログインエラーと表示され、次のclickにてNoMethodErrorで終了する。
# ログインエラーのハンドリングはしない。
puts "損益状況click"
detailPage = homePage.link_with(:text => '損益状況').click
data = Array.new(11){Array.new(8)}
puts "損益状況header"
ii=0
ii = acquire_asset(detailPage,data,ii,"PT_NEN",RUI_ASSET_KEYS)
numberOfPages = detailPage.search('input[@name=PT_PAGE1]').map{|value| value["value"]}[0].to_i
puts "number of pages = " + numberOfPages.to_s
for kk in 1..numberOfPages do
  puts "損益状況" + kk.to_s + "ページ目 読込"
  ii = acquire_asset(detailPage,data,ii,"PT_AST",ASSET_KEYS)
  detailPage = detailPage.form_with(:name => 'N_FORM').submit if kk < numberOfPages
end

puts "ii = " + ii.to_s
puts data[0..(ii-1)]

puts <<EOS




     確かめ




EOS
sum=0
for jj in 1..(ii-1) do
  sum += data[jj][5]
end
puts "sum        = " + sum.to_s
puts "data[0][1] = " + data[0][1].to_s
if sum == data[0][1] then
  puts "確かめOK"
else
  puts "error:確かめNG"
  exit
end

puts <<EOS




     エクセルファイル作成




EOS

for column in 0..(RUI_ASSET_KEYS.size-1) do
  worksheet.add_cell(row,column,data[0][column])
end
for jj in 1..(ii-1) do
  for column in 0..(ASSET_KEYS.size-1) do
    worksheet.add_cell(row,RUI_ASSET_KEYS.size+ASSET_KEYS.size*(jj-1)+column,data[jj][column])
  end
end
workbook.write(XLSX_FILE)

puts :end
