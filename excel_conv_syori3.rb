# 《説明》
# ユーザー.txtを読み込む。
# ユーザー名をキーに端末ID割付表.txtを読み込み、ユーザー名に該当するユーザーキーを取得する。
# ユーザー名とユーザーキーを元に、ユーザー_new.txtを新規作成する
# aaaaa

require "csv"

# PATH = 'C:\vagrant\appsuite'
FIOT1 = '端末ID割付表'
FIOT2 = 'ユーザー'

def upd_file?()

    begin
        # 新ファイルを削除
        File.delete(".\\#{FIOT2}_new.txt") if File.exist?(".\\#{FIOT2}_new.txt")

        CSV.open(".\\#{FIOT2}_new.txt", "w") do |new|   # 新ファイルをオープン

            CSV.foreach(".\\#{FIOT2}.txt") do |csv|     # OutPutファイルを読み込み

                row = []
                row << csv[0].to_s
                str = search_file(csv[0].to_s)          # ユーザキーを検索
                row << str
                
                new.puts(row)                           # 書き込み
            end
        end

        File.delete(".\\#{FIOT2}.txt") if File.exist?(".\\#{FIOT2}.txt")
        File.rename(".\\#{FIOT2}_new.txt", ".\\#{FIOT2}.txt") if File.exist?(".\\#{FIOT2}_new.txt")     # ファイル名をリネーム
        true

    rescue => e
        puts "Error = " + e.message
        false
    end
end

def search_file(user)

    ret = ""

    CSV.read(".\\#{FIOT1}.txt").map do |csv|            # OutPutファイルを読み込み

        if csv[0].to_s.include?("#{user}")
            ret = csv[1].to_s
            break
        end
    end
    ret
end

puts "処理が正常終了しました" if upd_file?
exit