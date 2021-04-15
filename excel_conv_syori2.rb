# 《説明》
# コルソスシリアル番号管理.xlsxを読み込む。
# 全シートを取得して、ユーザー.txtに書き込む
# 各シートの施設情報を取得して、端末ID割付表.txtに書き込む

# 《このプログラムを動かすための条件》
#   ・施設コードの先頭行はA5からスタートしている事
#   ・空白になっている施設名がない事
#   ・日付はyyyy/mm/dd形式で登録されている事

# ループしながら、標準出力をする場合、これをすると、１件づつ表示してくれる。
# ないと、バッファに溜まった標準出力が一気にドサッと表示される事になるので見難くなる。
# ループしながら大量データを標準出力する時は必須
STDOUT.sync = true

require 'win32ole'
require 'csv'

# Excel VBA定数のロード（お約束）
module Excel; end

# 下記４行はお約束
def init_excel()
    
    excel = WIN32OLE.new('Excel.Application')                   # Excelオブジェクト生成
    excel.visible = false                                       # Excelを見えるように起動する
    excel.displayAlerts = false                                 # メッセージを表示させない
    WIN32OLE.const_load(excel, Excel)                           # ここでExcelが起動される
    return excel
end

def read_excel(excel, file, sheet_num = 1)

    book = excel.Workbooks.Open(file)
    @sheet_cnt = book.WorkSheets.count                             # シート数を取得
    
    file_create("ユーザー.txt","施設情報.txt")                      # txtの空ファイルを作成

    CSV.open("施設情報.txt", "w") do |csv| 

        # ヘッダ情報を更新
        _head = "ユーザー,施設コード,施設名,通報(親),電源,機器１,通報(子),日付".encode("Windows-31J")
        csv.puts(_head.split(','))

        # ヘッダ情報を更新
        _head = "ユーザー名".encode("Windows-31J")
        File.open("ユーザー.txt", "a") { |f| f.puts(_head.split(',')) }  # シート名をAppend

        # １シートづつ処理
        book.WorkSheets.each do |_sheet|

            # 明細単位で更新
            File.open("ユーザー.txt", "a") { |f| f.puts(_sheet.Name) }  # シート名をAppend
            sheet = book.Worksheets(_sheet.Name)

            i = 0
            # シートに登録されている施設を取得
            loop do
                break if (sheet.Cells.item(5+i,2).value.to_s.empty? || i == 100)

                _row = []
                _row << _sheet.Name
                _row << sheet.Cells.item(5+i,1).value.to_i      # 施設コード
                _row << sheet.Cells.item(5+i,2).value.to_s      # 施設名
                _row << sheet.Cells.item(5+i,3).value.to_s      # コルソス通報（親）
                _row << sheet.Cells.item(5+i,4).value.to_s      # コルソス電源
                
                # s-Jisでエンコードしないとエラーになる事象の回避
                _sonotakiki = sheet.Cells.item(5+i,5).value.to_s.empty? ? "" : "コルソス子機１".encode!("Windows-31J")
                _row << _sonotakiki

                _row << sheet.Cells.item(5+i,5).value.to_s      # コルソス通報（子）
                
                str1 = sheet.Cells.item(5+i,6).value.to_s.match(/^[0-9]{4}-[0-9]{2}-[0-9]{2}/)
                str2 = str1.to_s.gsub('-', '/') 
                _row << str2                                    # 日付
                
                # 明細単位で更新
                csv.puts(_row)
                i += 1
            end
        end
        book.close
    end
end

def main()

    fso = WIN32OLE.new('Scripting.FileSystemObject')                    # OLE32用FileSystemObject生成
    file = fso.GetAbsolutePathName('コルソスシリアル番号管理.xlsx')       # 要は絶対パスを取得する
    # これでもOKではある
    # file = 'C:\vagrant\appsuite\コルソスシリアル番号管理.xlsx'

    unless File.exist?(file)
        puts "取り込み元のExcelファイルが存在しません" .encode!("Windows-31J")
        false
    else
        excel = init_excel()                                    # 初期処理
        read_excel(excel, file)                                 # Excel読み込み
        excel.quit()                                            # Excel終了
        true
    end
end

def file_create(*filname)

    filname.each do |fil|
        File.open(fil, "w").close
        File.delete(fil) if File.exist?(fil)
    end
end

start_time = Time.now

unless main
    puts "処理が中断されました"
else
    puts "*** 処理結果 ***"
    puts "処理時間 : " + "#{Time.at(Time.now - start_time).strftime('%M:%S')}"
    puts "シート数 : " + "#{@sheet_cnt}"
end
