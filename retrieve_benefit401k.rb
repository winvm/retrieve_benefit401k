#!ruby -KU
# coding: utf-8
require 'mechanize'
require 'logger'
require 'optparse'
require 'rubyXL'
require 'active_support'
require 'active_support/core_ext'

RUI_ASSET_KEYS = ["PT_NEN","PT_SHISAN","PT_KOU_RUI","PT_KOU_SON","PT_KYO_RUI","PT_KYO_SON"]
ASSET_KEYS = ["PT_SHN_SY","PT_SHN_ME","PT_HIRITU","PT_SURYO","PT_TANKA","PT_HYOU_GAK","PT_UKE_GAK","PT_HYOU_SON"]
XLSX_FILE = "./benefit401k.xlsx"

def acquire_asset(page,data,ii,input_name,key_values)
  page.search('input[@name='+input_name+']').each do | asset|
    jj=0
    key_values.each do |target|
      case target
      when "PT_SHN_SY","PT_SHN_ME" then
        data[ii][jj] = asset.parent.search('input[@name='+target+']').map{|value| value["value"]}[0].to_s.rstrip.tr("（）","()")
      when "PT_HIRITU" then
        data[ii][jj] = asset.parent.search('input[@name='+target+']').map{|value| value["value"]}[0].to_f
      when "PT_SURYO","PT_TANKA","PT_HYOU_GAK","PT_UKE_GAK","PT_HYOU_SON" then
        data[ii][jj] = asset.parent.search('input[@name='+target+']').map{|value| value["value"]}[0].delete(",").to_i
      when "PT_NEN" then
        data[ii][jj] = asset.parent.search("input[@name=#{target}]").map{|value| value["value"]}[0].to_s + "時点"
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

log = Logger.new("benefit401k.log", "monthly")
log.progname = "benefit401k"
log.level = Logger::DEBUG

params = ARGV.getopts('',"account:./account.txt","xlsx:#{XLSX_FILE}")
log.debug("エクセルファイル オープン")
workbook = RubyXL::Parser.parse(params["xlsx"].to_s)
worksheet = workbook[0]
yesterday = Date.yesterday
row=0
for row_data in worksheet do
  log.debug( "row = " + row.to_s + " : " + row_data[0].value.to_s )
  if( "日付" == row_data[0].value.to_s ) then
    row += 1
    next
  end
  #log.debug(row_data[0].value.to_s
  #log.debug(row_data[0].value
  latest = Date.strptime(row_data[0].value,'%Y年%m月%d日時点')
  if( yesterday == latest ) then
    log.info("既に#{yesterday.to_s}のデータは登録済みです。")
    exit
  end
  break if( nil == row_data[0].value )
  row += 1
end
log.debug("new row = #{row.to_s}")
log.debug("latest = #{latest.to_s}\n")

account_file = open(params["account"].to_s, "r")
if( !account_file ) then
  log.error("アカウント情報ファイルが開けませんでした。")
  exit 
end
UserId = account_file.gets.chomp
PassWdS = account_file.gets.chomp
account_file.close

agent = Mechanize.new
agent.log = log
OpenSSL::debug = false

log.debug("agent.user_agent = " + agent.user_agent)
log.debug("agent.verify_mode = " + agent.verify_mode.to_s)
log.debug("agent.ssl_version = " + agent.ssl_version.to_s)
log.debug("agent.ca_file = " + agent.ca_file.to_s)
log.debug("OpenSSL::X509::DEFAULT_CERT_DIR = " + OpenSSL::X509::DEFAULT_CERT_DIR)
log.debug("OpenSSL::X509::DEFAULT_CERT_FILE = " + OpenSSL::X509::DEFAULT_CERT_FILE)
log.debug("agent.cert = " + agent.cert.to_s)
log.debug("agent.key = " + agent.key.to_s)
log.debug("agent.pass = " + agent.pass.to_s)
log.debug("agent.verify_callback = " + agent.verify_callback.to_s)
log.debug("agent.cookies = " + agent.cookies.to_s)

agent.user_agent = 'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0)'
agent.verify_mode = OpenSSL::SSL::VERIFY_PEER
agent.ssl_version = 'TLSv1'
# agent.ca_file = './GTEGlRoot.txt'.to_s
agent.ca_file = './BaltimoreCyberTrustRoot.crt'

url = 'https://www.benefit401k.com/customer/'

log.debug('設定完了')
log.debug("agent.user_agent = " + agent.user_agent)
log.debug("agent.verify_mode = " + agent.verify_mode.to_s)
log.debug("agent.ssl_version = " + agent.ssl_version.to_s)
log.debug("agent.ca_file = " + agent.ca_file.to_s)
log.debug("OpenSSL::X509::DEFAULT_CERT_DIR = " + OpenSSL::X509::DEFAULT_CERT_DIR)
log.debug("OpenSSL::X509::DEFAULT_CERT_FILE = " + OpenSSL::X509::DEFAULT_CERT_FILE)
log.debug("agent.cert = " + agent.cert.to_s)
log.debug("agent.key = " + agent.key.to_s)
log.debug("agent.pass = " + agent.pass.to_s)
log.debug("agent.verify_callback = " + agent.verify_callback.to_s)
log.debug("agent.cookies = " + agent.cookies.to_s)
log.debug("url = " + url.to_s)

login_page = agent.get(url)

log.debug("\nlogin_page.title = " + login_page.title.to_s)
case login_page.title
  when "確定拠出年金サイトへログイン"
    log.debug("login画面入手OK")
  when "システムメンテナンス中"
    log.info("\nシステムメンテナンス中\n")
    exit
  else
    log.fatal("\nUnknown title\n")
    exit
end
log.debug("agent.cookies = " + agent.cookies.to_s)
home_page = login_page.form_with(:name => 'signonform') do |form|
  form.UserId = UserId
  form.PassWdS = PassWdS
end.submit
log.debug("agent.cookies = " + agent.cookies.to_s)
log.debug("\nhome_page.title = " + home_page.title.to_s)
my_desktop = home_page.search("div[@class='p-tx']").first.text
log.info(my_desktop)
if( !my_desktop.include?("様のデスクトップです。") ) then
  log.error("\nログインNG\n")
  exit
end
# もしログインエラーなら、home_page.title = ログインエラーと表示され、次のclickにてNoMethodErrorで終了する。
# ログインエラーのハンドリングはしない。
log.debug("\n損益状況click")
detail_page = home_page.link_with(:text => '損益状況').click
data = Array.new(11){Array.new(8)}
log.debug("損益状況header")
ii=0
ii = acquire_asset(detail_page,data,ii,"PT_NEN",RUI_ASSET_KEYS)
number_of_pages = detail_page.search('input[@name=PT_PAGE1]').map{|value| value["value"]}[0].to_i
log.debug("number of pages = " + number_of_pages.to_s)
for kk in 1..number_of_pages do
  log.debug("損益状況" + kk.to_s + "ページ目 読込")
  ii = acquire_asset(detail_page,data,ii,"PT_AST",ASSET_KEYS)
  detail_page = detail_page.form_with(:name => 'N_FORM').submit if kk < number_of_pages
end

log.debug("ii = " + ii.to_s)
#log.debug(data[0..(ii-1)].join(",\t"))
data[0..(ii-1)].each do |tmp|
  log.info(tmp.join(",\t"))
end

log.debug("LOGOUT")
detail_page = detail_page.search("img[@alt='ログアウト']").first.parent
log.debug(detail_page)
log.debug(detail_page['href'])
detail_page = agent.get(detail_page['href'])
log.debug(detail_page.title)
if( detail_page.title == "ログアウト" ) then
  log.info("ログアウト完了")
else
  log.fatal("ログアウトNG")
  exit
end

log.info("確かめ")
sum=0
for jj in 1..(ii-1) do
  sum += data[jj][5]
end
log.info("sum        = " + sum.to_s)
log.info("data[0][1] = " + data[0][1].to_s)
if sum == data[0][1] then
  log.info("確かめOK")
else
  log.fatal("error:確かめNG")
  exit
end

if( latest == Date.strptime(data[0][0],'%Y年%m月%d日時点') )then
  log.info("latest.to_s = #{latest.to_s}")
  log.info("Date.strptime(data[0][0],'%Y年%m月%d日時点') = #{Date.strptime(data[0][0],'%Y年%m月%d日時点')}")
  log.info("\nalready inputed\n")
  exit
end
log.info("エクセルファイル作成")
for column in 0..(RUI_ASSET_KEYS.size-1) do
  worksheet.add_cell(row,column,data[0][column])
end
for jj in 1..(ii-1) do
  for column in 0..(ASSET_KEYS.size-1) do
    worksheet.add_cell(row,RUI_ASSET_KEYS.size+ASSET_KEYS.size*(jj-1)+column,data[jj][column])
  end
end

count = 0
for jj in 1..ii-1 do
  sheet_count = 0
  already_inputed = false
  workbook.each do |sheet|
    if( sheet[0][0].value == data[jj][1] ) then
      log.debug("jj = #{jj}")
      log.debug("sheet_name = #{sheet.sheet_name}")
      log.debug("sheet[0][0].value = #{sheet[0][0].value}")
      log.debug("data[jj][1]       = #{data[jj][1]}")
      here = 0
      sheet.each do |tmp|
        if( tmp[0].value == data[0][0] ) then
          log.info("#{sheet[here][0].value}のデータは登録済みです。")
          already_inputed = true
          break
        end
        here += 1
      end
      break if( already_inputed )
      log.debug("here = #{here}")
      sheet.add_cell(here,0,data[0][0])
      sheet.add_cell(here,1,data[jj][5].to_f/data[0][1].to_f)
      for column in 0..(ASSET_KEYS.size-1) do
        sheet.add_cell(here,2+column,data[jj][column])
      end
      break
    end
    sheet_count += 1
  end
  if( (count+1) != sheet_count ) then
    log.debug("sheet_count = #{sheet_count}")
    log.debug("エラー? ダブり?")
  end
  count += 1
end
log.info("count = #{count}")
if( (ii-1) != count ) then
  log.fatal("error")
end

workbook.write(XLSX_FILE)

log.info(:end)
