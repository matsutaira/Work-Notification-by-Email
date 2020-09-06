#ライブラリの読み込み
import tkinter
import win32com.client
import configparser
import datetime
import win32gui, win32con

#コンソールを非表示にする
hide = win32gui.GetForegroundWindow()
win32gui.ShowWindow(hide, win32con.SW_HIDE)

#configparseの宣言とiniファイルの読み込み
config_ini = configparser.ConfigParser()
config_ini.read('./config.ini', encoding='utf-8')

#ラジオボックスの定義
def rdo_input():
    #ラジオボタンのクラス設定
    tki = tkinter.Tk()
    tki.title('就業連絡の作成')

    #ラジオボタンのラベルをリスト化
    rdo_txt = ['在宅勤務開始', '在宅勤務終了']
    result = 'NotSelected'
    rdo_var = tkinter.StringVar(value=result)

    #ラジオボタンをrdo_txtリストから作成して配置
    frame1 =tkinter.Frame(tki)
    frame1.pack()
    label1 = tkinter.Label(frame1, text='【開始/終了】')
    label1.pack(side='left')
    for i in rdo_txt:
        #rdo = tkinter.Radiobutton(tki, value=i, variable=rdo_var, text=i)
        #rdo.pack(anchor = tkinter.CENTER, padx=30, pady=3)
        rdo = tkinter.Radiobutton(frame1, value=i, variable=rdo_var, text=i)
        rdo.pack(side='left')
    
    #時刻入力ボックスの配置
    frame2 = tkinter.Frame(tki)
    frame2.pack()
    label2 = tkinter.Label(frame2, text='【業務時間帯】')
    label2.pack(side='left')
    timebox = tkinter.Entry(frame2)
    timebox.insert(tkinter.END, config_ini['WORK']['TIME'])
    timebox.pack(side='left')

    #業務内容ボックスの配置
    frame3 = tkinter.Frame(tki)
    frame3.pack()
    label3 = tkinter.Label(frame3, text='【業務内容】')
    label3.pack(side='left')
    jobbox = tkinter.Text(frame3)
    jobbox.insert(tkinter.END, config_ini['WORK']['CONTENTS'])
    jobbox.pack(side='left')
        
    #ボタンクリックイベント
    def btn_click():
        nonlocal result
        result = rdo_var.get()
        config_ini['WORK']['TIME'] = timebox.get()
        config_ini['WORK']['CONTENTS'] = jobbox.get('1.0', 'end -1c')
        tki.destroy()
    
    #決定ボタンの作成
    btn = tkinter.Button(tki, text='作成！', command=btn_click)
    btn.pack(fill ='x', padx=40, pady=3)
    tki.mainloop()
    return result

#ラジオボックスで選択した値を入力
subject = rdo_input()

#ラジオボックスの値確認
if subject == 'NotSelected':
    print('メール作成は実行されませんでした。')
    exit()

#メールを作成
else:
    #Outlookのオブジェクト設定
    outlook = win32com.client.Dispatch('Outlook.Application')
    mymail = outlook.CreateItem(0)
    
    #datetimeのフォーマット設定
    today = datetime.date.today().strftime('%#m/%#d')
    
    #署名
    sign = config_ini['ADDRESS']['NAME']
    
    #メールの作成
    mymail.BodyFormat = 1
    mymail.To = config_ini['ADDRESS']['TO']
    mymail.Cc = config_ini['ADDRESS']['CC']
    mymail.Bcc = config_ini['ADDRESS']['BCC']
    mymail.Subject = today + ' ' + subject + ' ' + config_ini['ADDRESS']['USER']

    #メール本文の作成
    body1 = config_ini['DEFAULT']['BODY1'] + '\n\n' + config_ini['DEFAULT']['BODY2'] + '\n' + config_ini['DEFAULT']['NAME'] + 'です。' + '\n\n'
    body2 = '\n\n' + config_ini['DEFAULT']['BODY4'] + '\n' + config_ini['WORK']['TIME'] + '\n\n'+ config_ini['DEFAULT']['BODY5'] + '\n' + config_ini['WORK']['CONTENTS'] + '\n\n'+ config_ini['DEFAULT']['BODY6'] + '\n\n' + sign

    if subject == '在宅勤務開始':
        mymail.Body = body1 + config_ini['START']['BODY3'] + body2
    
    if subject ==  '在宅勤務終了':
        mymail.Body = body1 + config_ini['FINISH']['BODY3'] + body2
    
    #出来上がったメール確認
    mymail.Display(True)
    #確認せず送信する場合は、mymail.Display(True)を消して、下記コードを使用する
    #mymail.Send()
    
    #config.ini の更新
    config_update = configparser.RawConfigParser()
    config_update.read('./config.ini', encoding='utf-8')
    config_update.set('WORK', 'TIME', config_ini['WORK']['TIME'])
    config_update.set('WORK', 'CONTENTS', config_ini['WORK']['CONTENTS'])
    with open('./config.ini', 'w', encoding='utf-8') as f:
        config_update.write(f)