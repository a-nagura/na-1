# Description:
#   Example scripts for you to examine and try out.
#
# Notes:
#   They are commented out by default, because most of them are pretty silly and
#   wouldn't be useful and amusing enough for day to day huboting.
#   Uncomment the ones you want to try and experiment with.
#
#   These are from the scripting documentation: https://github.com/github/hubot/blob/master/docs/scripting.md
cron = require('cron').CronJob
#channel = '#notifications'
channel = '#nagura_test'

module.exports = (robot) ->

    cron = new cron(
     cronTime: "0 00 9 * * *"     # 実行時間
     start:    true              # すぐにcronのjobを実行するか
     timeZone: "Asia/Tokyo"      # タイムゾーン指定
     onTick: ->                  # 時間が来た時に実行する処理
        #robot.send {room: "#nagura_test"}, "cron test!"


     #robot.respond /holiday?/i, (msg) ->
        #Excelを取得
        XLSX = require('xlsx')

        #Excel情報
        Excel_row = ['F','G','H','I','J','K','L','M','N','P','Q','S','T']
        sheet_name = '予定表'
        file_name = '/vagrant_data/holidaybot/Workingday.xlsx'
        

        #シリアル値→unix時間への変換用
        COEFFICIENT = 24 * 60 * 60 * 1000
        DATES_OFFSET = 70 * 365 + 17 + 1
        MILLIS_DIFFERENCE = 9 * 60 * 60 * 1000

        #Excelファイル取得
        workbook = XLSX.readFile(file_name)

        #シートを取得
        worksheet = workbook.Sheets[sheet_name]


        #本日日付を（yyyy/mm/dd)形式で出力
        now_d = new Date
        now_year = now_d.getFullYear()     # 年
        now_month = now_d.getMonth() + 1    # 月
        now_day = now_d.getDate()       # 日
        now_data = "#{now_year}/#{now_month}/#{now_day}"
        
        #本日日付をholiday_kekkaに入れる
        holiday_kekka = """#{now_data} の出勤情報
        """


        #Excelシートから本日日付の休み予定を取得
        for n in [530..2027]
         cell_data = worksheet['C'+ n].v

         #シリアル値をUNIXtimeへ変換、UNIXtimeをyyyy/mm/ddに変換
         cell_data_unix = (cell_data - DATES_OFFSET) * COEFFICIENT - MILLIS_DIFFERENCE
         Excel_d = new Date(cell_data_unix)
         Excel_year  = Excel_d.getFullYear()
         Excel_month = Excel_d.getMonth() + 1
         Excel_day = Excel_d.getDate()
         Excel_data = "#{Excel_year}/#{Excel_month}/#{Excel_day}"

         #Excelの日付と本日の日付を比較、セルの行番号をday_celに入れる
         if now_data == Excel_data then nowdata_cell = n
         


        #休みの予定を取得
        for i in [0..12]
         name = worksheet[Excel_row[i]+'3']
         day = worksheet[Excel_row[i]+nowdata_cell]

         #セルの値を取得
         a = if name then name.v else undefined
         b = if day then day.v else '出勤'
         
         #セルの値をholiday_kekkaへ入れる
         holiday_kekka = """#{holiday_kekka}
         #{a} : #{b}
         """
         
        #日付と休みの予定を足したholiday_kekkaを出力
        robot.send {room: "#notifications"}, "#{holiday_kekka}"
    )
    robot.respond /holiday?/i, (msg) ->
        #Excelを取得
        XLSX = require('xlsx')

        #Excel情報
        Excel_row = ['F','G','H','I','J','K','L','M','N','P','Q','S','T']
        sheet_name = '予定表'
        file_name = '/vagrant_data/holidaybot/Workingday.xlsx'
        

        #シリアル値→unix時間への変換用
        COEFFICIENT = 24 * 60 * 60 * 1000
        DATES_OFFSET = 70 * 365 + 17 + 1 + 1
        MILLIS_DIFFERENCE = 9 * 60 * 60 * 1000

        #Excelファイル取得
        workbook = XLSX.readFile(file_name)

        #シートを取得
        worksheet = workbook.Sheets[sheet_name]


        #本日日付を（yyyy/mm/dd)形式で出力
        now_d = new Date
        now_unix = now_d.getTime()
        now_year = now_d.getFullYear()     # 年
        now_month = now_d.getMonth() + 1    # 月
        now_day = now_d.getDate()       # 日
        now_data = "#{now_year}/#{now_month}/#{now_day}"
        
        now_data_serial = (now_unix + MILLIS_DIFFERENCE) / COEFFICIENT + DATES_OFFSET
        
        #1、文字列変換　2、先頭5行を抽出　3、数値に変換
        now_data_serial = now_data_serial + ''
        now_data_serial = now_data_serial.slice(0,5)
        now_data_serial = now_data_serial - 0

        console.log now_data_serial
        #本日日付をholiday_kekkaに入れる
        holiday_kekka = """#{now_data} の出勤情報
        """

        


        #Excelシートから本日日付の休み予定を取得
        gyou = 500
        cell_data = worksheet['C'+ gyou].v

        nowdata_cell = now_data_serial - cell_data + gyou 
        
        console.log nowdata_cell
        

        #Excelの日付と本日の日付を比較、セルの行番号をday_celに入れる
        #if now_data == Excel_data then nowdata_cell = n
         

        title = now_data + 'の出勤予定'
        holiday_kekka_name = []
        holiday_kekka_day = []
        #休みの予定を取得
        for i in [0..12]
         name = worksheet[Excel_row[i]+'3']
         day = worksheet[Excel_row[i]+nowdata_cell]

         #セルの値を取得
         a = if name then name.v else undefined
         b = if day then day.v else '出勤'
         
         #セルの値をholiday_kekkaへ入れる
         holiday_kekka_name.push(a)
         holiday_kekka_day.push(b)
         
        #日付と休みの予定を足したholiday_kekkaを出力
        #robot.send {room: "#{channel}"}, "#{holiday_kekka}"


        # https://api.slack.com/docs/message-attachments
        
        attachments = [
            {
                fallback: "#{title}",
                color: "#FF0000",
                pretext: "#{title}",
                fields: [
                    {
                        title: 'SSDMember',
                        value: " #{holiday_kekka_name[0]}\n #{holiday_kekka_name[1]}\n #{holiday_kekka_name[2]}\n #{holiday_kekka_name[3]}\n #{holiday_kekka_name[4]}\n #{holiday_kekka_name[5]}\n #{holiday_kekka_name[6]}\n #{holiday_kekka_name[7]}\n #{holiday_kekka_name[8]}",
                        short: true
                    },
                    {
                        title: 'status',
                        value: " #{holiday_kekka_day[0]}\n #{holiday_kekka_day[1]}\n #{holiday_kekka_day[2]}\n #{holiday_kekka_day[3]}\n #{holiday_kekka_day[4]}\n #{holiday_kekka_day[5]}\n #{holiday_kekka_day[6]}\n #{holiday_kekka_day[7]}\n #{holiday_kekka_day[8]}",
                        short: true
                    }
                ]
            },
            {
                color: "#00FF00",
                fields: [
                    {
                        title: 'OTRMember',
                        value: "#{holiday_kekka_name[9]}\n #{holiday_kekka_name[10]}\n #{holiday_kekka_name[11]}\n #{holiday_kekka_name[12]}",
                        short: true
                    },
                    {
                        title: 'status',
                        value: "#{holiday_kekka_day[9]}\n #{holiday_kekka_day[10]}\n #{holiday_kekka_day[11]}\n #{holiday_kekka_day[12]}",
                        short: true
                    }
                ]
            }
        ]

        options = { as_user: true, link_names: 1, attachments: attachments }
        robot.send {room: "#{channel}"}, options






 

 