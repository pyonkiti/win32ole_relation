# 《説明》
# 端末ID割付表.xlsxをサーバーからローカルにコピー
# 端末ID割付表.xlsxを動して、先頭のシートを開く

# 《動かし方》
# Start Command Prompt with Rubyを起動（cmdを起動してもRubyは動きません）
# .\excel_get.batコマンドで、C:\User\NSK\excel_get.batを実行

# 参考になったサイト
# https://www.texcell.co.jp/ruby/excel/RubyExcel.html
# http://nodding-off-programmer.blogspot.com/2017/09/ruby-win32oleexcel.html

require "fileutils"
require 'win32ole'

PATH_S = '\\\192.168.19.12\share\ASNET\DevelopEnvironment\SofinetCloud\\'
PATH_L = 'D:\\C\\1\\環境\\端末ID管理表取得\\'
FILE_S = '端末ID割付表.xlsx'
FILE_L = '_端末ID割付表.xlsx'
SHEET = 'ユーザー一覧'

module Excel; end                                               # Excel VBA定数のロード

def excel_init?

    begin
        excel = WIN32OLE.new('Excel.Application')                   # Excelオブジェクト生成
        excel.visible = true                                        # 表示可
        excel.displayAlerts = false                                 # メッセージを非表示
        WIN32OLE.const_load(excel, Excel)                           # Excelが起動

        puts    "#{PATH_L}#{FILE_L}"
        book = excel.Workbooks.Open("#{PATH_L}#{FILE_L}")           # ExcelをOpen
        
        sheet = book.sheets("#{SHEET}")
        sheet.Activate                                              # 指定したシートをアクティブにする
        true

    rescue => e
        puts "例外エラーが発生しました => " + e.message.to_s
        false
    end
    # book.close
    # excel.quit()
end

def get_file?

    unless File.exist?("#{PATH_S}#{FILE_S}")
        false
    else
        FileUtils.cp("#{PATH_S}#{FILE_S}","#{PATH_L}#{FILE_L}")
        true
    end
    
end

unless get_file?
    puts "コピー元ファイルが存在しません。"
else
    puts "コピーOK"
end

puts "正常終了しました。" if excel_init?
exit