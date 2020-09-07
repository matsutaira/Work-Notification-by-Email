#ライブラリの読み込み
import tkinter
import win32com.client
import configparser
import datetime
import win32gui, win32con

#configparseの宣言とiniファイルの読み込み
config_ini = configparser.ConfigParser()
config_ini.read('./config.ini', encoding='utf-8')

#ホップアップメニューの定義
def rdo_input():
    #ホップアップメニューのクラス設定
    tki = tkinter.Tk()
    tki.title('就業連絡の作成')
    
    #UserID入力ボックスの配置
    frame_user = tkinter.Frame(tki)
    frame_user.pack()
    label_user = tkinter.Label(frame_user, text='【ユーザID】')
    label_user.pack(side='left')
    userbox = tkinter.Entry(frame_user, width=80)
    userbox.insert(tkinter.END, config_ini['INFO']['user'])
    userbox.pack(side='left')

    #名前入力ボックスの配置
    frame_name = tkinter.Frame(tki)
    frame_name.pack()
    label_name = tkinter.Label(frame_name, text='【名前】')
    label_name.pack(side='left')
    namebox = tkinter.Entry(frame_name, width=80)
    namebox.insert(tkinter.END, config_ini['INFO']['name'])
    namebox.pack(side='left')

    #Toアドレス入力ボックスの配置
    frame_to = tkinter.Frame(tki)
    frame_to.pack()
    label_to = tkinter.Label(frame_to, text='【To】')
    label_to.pack(side='left')
    tobox = tkinter.Entry(frame_to, width=80)
    tobox.insert(tkinter.END, config_ini['ADDRESS']['to'])
    tobox.pack(side='left')

    #Ccアドレス入力ボックスの配置
    frame_cc = tkinter.Frame(tki)
    frame_cc.pack()
    label_cc = tkinter.Label(frame_cc, text='【Cc】')
    label_cc.pack(side='left')
    ccbox = tkinter.Entry(frame_cc, width=80)
    ccbox.insert(tkinter.END, config_ini['ADDRESS']['cc'])
    ccbox.pack(side='left')

    #Bccアドレス入力ボックスの配置
    frame_bcc = tkinter.Frame(tki)
    frame_bcc.pack()
    label_bcc = tkinter.Label(frame_bcc, text='【Bcc】')
    label_bcc.pack(side='left')
    bccbox = tkinter.Entry(frame_bcc, width=80)
    bccbox.insert(tkinter.END, config_ini['ADDRESS']['bcc'])
    bccbox.pack(side='left')

    #宛名入力ボックスの配置
    frame_addname = tkinter.Frame(tki)
    frame_addname.pack()
    label_addname = tkinter.Label(frame_addname, text='【宛名】')
    label_addname.pack(side='left')
    addnamebox = tkinter.Entry(frame_addname, width=80)
    addnamebox.insert(tkinter.END, config_ini['ADDRESS']['addname'])
    addnamebox.pack(side='left')

    #ラジオボタンのラベルをリスト化
    rdo_txt = ['在宅勤務開始', '在宅勤務終了']
    result = 'NotSelected'
    rdo_var = tkinter.StringVar(value=result)

    #ラジオボタンをrdo_txtリストから作成して配置
    frame1 = tkinter.Frame(tki)
    frame1.pack()
    label1 = tkinter.Label(frame1, text='【開始/終了】')
    label1.pack(side='left')
    for i in rdo_txt:
        rdo = tkinter.Radiobutton(frame1, value=i, variable=rdo_var, text=i)
        rdo.pack(side='left')
    
    #時刻入力ボックスの配置
    frame2 = tkinter.Frame(tki)
    frame2.pack()
    label2 = tkinter.Label(frame2, text='【業務時間帯】')
    label2.pack(side='left')
    timebox = tkinter.Entry(frame2)
    timebox.insert(tkinter.END, config_ini['WORK']['time'])
    timebox.pack(side='left')

    #業務内容ボックスの配置
    frame3 = tkinter.Frame(tki)
    frame3.pack()
    label3 = tkinter.Label(frame3, text='【業務内容】')
    label3.pack(side='left')
    jobbox = tkinter.Text(frame3, width=80)
    jobbox.insert(tkinter.END, config_ini['WORK']['contents'])
    jobbox.pack(side='left')
        
    #ボタンクリックイベント
    def btn_click():
        nonlocal result
        result = rdo_var.get()
        config_ini['INFO']['user'] = userbox.get()
        config_ini['INFO']['name'] = namebox.get()
        config_ini['ADDRESS']['to'] = tobox.get()
        config_ini['ADDRESS']['cc'] = ccbox.get()
        config_ini['ADDRESS']['bcc'] = bccbox.get()
        config_ini['ADDRESS']['addname'] = addnamebox.get()
        config_ini['WORK']['time'] = timebox.get()
        config_ini['WORK']['contents'] = jobbox.get('1.0', 'end -1c')
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
    sign = config_ini['INFO']['name']
    
    #メールの作成
    mymail.BodyFormat = 1
    mymail.To = config_ini['ADDRESS']['to']
    mymail.Cc = config_ini['ADDRESS']['cc']
    mymail.Bcc = config_ini['ADDRESS']['bcc']
    mymail.Subject = today + ' ' + subject + ' ' + config_ini['INFO']['user']

    #メール本文の作成
    body1 = config_ini['ADDRESS']['addname'] + '\n\n' + config_ini['DEFAULT']['body1'] + '\n' + config_ini['INFO']['name'] + 'です。' + '\n\n'
    body2 = '\n\n' + config_ini['DEFAULT']['body2'] + '\n' + config_ini['WORK']['time'] + '\n\n'+ config_ini['DEFAULT']['body3'] + '\n' + config_ini['WORK']['contents'] + '\n\n'+ config_ini['DEFAULT']['body4'] + '\n\n' + sign

    if subject == '在宅勤務開始':
        mymail.Body = body1 + config_ini['START']['body_start'] + body2
    
    if subject ==  '在宅勤務終了':
        mymail.Body = body1 + config_ini['FINISH']['body_finish'] + body2
    
    #出来上がったメール確認
    mymail.Display(True)
    #確認せず送信する場合は、mymail.Display(True)を消して、下記コードを使用する
    #mymail.Send()
    
    #config.ini の更新
    config_update = configparser.RawConfigParser()
    config_update.read('./config.ini', encoding='utf-8')
    config_update.set('INFO', 'user', config_ini['INFO']['user'])
    config_update.set('INFO', 'name', config_ini['INFO']['name'])
    config_update.set('ADDRESS', 'to', config_ini['ADDRESS']['to'])
    config_update.set('ADDRESS', 'cc', config_ini['ADDRESS']['cc'])
    config_update.set('ADDRESS', 'bcc', config_ini['ADDRESS']['bcc'])
    config_update.set('ADDRESS', 'addname', config_ini['ADDRESS']['addname'])
    config_update.set('WORK', 'time', config_ini['WORK']['time'])
    config_update.set('WORK', 'contents', config_ini['WORK']['contents'])
    with open('./config.ini', 'w', encoding='utf-8') as f:
        config_update.write(f)