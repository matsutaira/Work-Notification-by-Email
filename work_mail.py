#ライブラリの読み込み
import tkinter
import win32com.client
import configparser
import datetime

#configparseの宣言とiniファイルの読み込み
config_ini = configparser.ConfigParser()
config_ini.read('./config.ini', encoding='utf-8')

#ラジオボックスの定義
def rdo_input():
    #ラジオボタンのクラス設定
    tki = tkinter.Tk()
    tki.geometry('300x200')
    tki.title('通知の選択')

    #ラジオボタンのラベルをリスト化
    rdo_txt = ['在宅勤務開始', '在宅勤務終了']
    result = 'NotSelected'
    rdo_var = tkinter.StringVar(value=result)

    #ラジオボタンをrdo_txtリストから作成して配置
    for i in rdo_txt:
        rdo = tkinter.Radiobutton(tki, value=i, variable=rdo_var, text=i)
        rdo.pack(anchor = tkinter.CENTER, padx=30, pady=3)
    
    #ボタンクリックイベント
    def btn_click():
        nonlocal result
        result = rdo_var.get()
        tki.destroy()
    
    #決定ボタンの作成
    btn = tkinter.Button(tki, text='決定', command=btn_click)
    btn.pack(fill ='x', padx=40, pady=3)
    tki.mainloop()
    return result

#ラジオボックスで選択した値を入力
subject = rdo_input()

#Outlookのオブジェクト設定
outlook = win32com.client.Dispatch('Outlook.Application')
mymail = outlook.CreateItem(0)

#datetimeのフォーマット設定
today = datetime.date.today().strftime('%#m/%#d')

#署名
sign = config_ini['ADDRESS']['NAME']

#メールの設定
mymail.BodyFormat = 1
mymail.To = config_ini['ADDRESS']['TO']
mymail.Cc = config_ini['ADDRESS']['CC']
mymail.Bcc = config_ini['ADDRESS']['BCC']
mymail.Subject = today + ' ' + subject + ' ' + config_ini['ADDRESS']['USER']

if subject == '在宅勤務開始':
    mymail.Body = config_ini['START']['BODY1'] + '\n\n' + config_ini['START']['BODY2'] + '\n' + config_ini['START']['NAME'] + 'です。' + '\n\n' + config_ini['START']['BODY3'] + '\n\n' + config_ini['START']['BODY4'] + '\n\n'+ config_ini['START']['BODY5'] + '\n\n'+ config_ini['START']['BODY6'] + '\n\n' + sign

if subject ==  '在宅勤務終了':
    mymail.Body = config_ini['FINISH']['BODY1'] + '\n\n' + config_ini['FINISH']['BODY2'] + '\n' + config_ini['FINISH']['NAME'] + 'です。' + '\n\n' + config_ini['FINISH']['BODY3'] + '\n\n' + config_ini['FINISH']['BODY4'] + '\n\n' + config_ini['FINISH']['BODY5'] + '\n\n' + config_ini['FINISH']['BODY6'] + '\n\n' + sign

#出来上がったメール確認
mymail.Display(True)
#確認せず送信する場合は、mymail.Display(True)を消して、下記コードを使用する
#mymail.Send()