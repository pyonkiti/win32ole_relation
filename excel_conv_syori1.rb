# 《説明》
# 端末ID割付表.xlsxを読み込む。
# ユーザー名とユーザーキーを抜き出し、端末ID割付表.txtファイルを作成する

# サンプルURL
# http://nodding-off-programmer.blogspot.com/2017/09/ruby-win32oleexcel.html

require 'win32ole'
require 'csv'

# PATH = '\\\192.168.19.12\share\ASNET\DevelopEnvironment\SofinetCloud'
PATH = 'C:\vagrant\appsuite'
FILE = '端末ID割付表'
SHET = 'ユーザー一覧'
HEAD = 'ユーザー名, ユーザーキー'

module Excel; end                                       # Excel VBA定数のロード

def excel_init?()
    
    # puts "#{PATH}\\#{FILE}.xlsx"

    unless File.exist?("#{PATH}\\#{FILE}.xlsx")
        false
    else
        @excel = WIN32OLE.new('Excel.Application')       # Excelオブジェクト生成
        @excel.visible = false                           # Excelを非表示
        @excel.displayAlerts = false                     # メッセージを非表示
        WIN32OLE.const_load(@excel, Excel)               # Excelを起動
        true
    end
end

def excel_read?()

    begin

        book = @excel.Workbooks.Open("#{PATH}\\#{FILE}.xlsx")
        # puts book

        sheet = book.Worksheets("#{SHET}")
        # puts sheet

        File.delete(".\\#{FILE}.txt") if File.exist?(".\\#{FILE}.txt")
        # puts ".\\#{FILE}.txt"

        CSV.open(".\\#{FILE}.txt", "w") do |csv| 

            # puts "#{HEAD.encode!("Windows-31J").split(',')}"

            csv.puts(HEAD.encode!("Windows-31J").split(','))                # ヘッダを書き込む
            iRowCnt = 0

            loop do
                break if iRowCnt == 300
                
                csvRow = []
                csvRow << sheet.Cells.item(2 + iRowCnt, 2).value.to_s       # ユーザー名
                csvRow << sheet.Cells.item(2 + iRowCnt, 14).value.to_s      # ユーザーキー

                csv.puts(csvRow)

                iRowCnt += 1
            end
        end

        book.close; @excel.quit
        true

    rescue => e
        puts "Error = " + e.message
        false
    end
end

ret = excel_init?
if ret
    puts "処理が正常終了しました" if excel_read?
else
    puts "取り込み元のExcelファイルが存在しません"
end

exit