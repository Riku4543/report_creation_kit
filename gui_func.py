# GUIを扱うモジュール
import tkinter
from tkinter import filedialog
from tkinter import ttk
# 日付を扱うモジュール
import datetime
from pdf_func import Pdfoperation
from word_func import Wordopration
from excel_func import Excelopration


class Authentication(tkinter.Frame):
    # 簡易的な認証システム
    def __init__(self,root):
        super().__init__(root)
        self.root = root
        # mainのスタイル
        self.s = ttk.Style()
        self.s.theme_use('vista')
        self.user_name = None
        self.pass_word = None
        self.check = False # 認証できたかの判断
        self.error = None
        self.pack()
        self.title()
        self.username()
        self.password()
        self.button()
        
    def title(self):
        #コンポーネントの作成
        title_frame = ttk.Frame(self.root) # 枠組みの作成
        title_mes = tkinter.Message(title_frame,text='ログイン認証画面',width=200) # 表示メッセージの作成
        self.error = tkinter.Message(title_frame,text='認証失敗',width=200,foreground = 'red')
        # コンポーネントの配置
        title_frame.pack(expand=True)
        title_mes.pack()
    
    def username(self):
        # username記入欄
        user_frame = ttk.Frame(self.root)
        self.user_name = ttk.Entry(user_frame,width=20) # 入力欄の作成
        user_mes = tkinter.Message(user_frame,text='UserName：',width=200)
        user_frame.pack(expand=True)
        user_mes.pack(side='left',expand=True,anchor='center')
        self.user_name.pack(side='left',expand=True,anchor='center')
        
    def password(self):
        # password記入欄
        pass_frame = ttk.Frame(self.root)
        self.pass_word=ttk.Entry(pass_frame,show='*',width=20)
        pass_mes = tkinter.Message(pass_frame,text='PassWord：',width=200)
        pass_frame.pack(expand=True)
        pass_mes.pack(side='left',expand=True,anchor='center')
        self.pass_word.pack(side='left',expand=True,anchor='center')
    
    def button(self):
        # 認証、終了させるためのボタン
        but_frame = ttk.Frame(self.root)
        button1 = ttk.Button(but_frame,text='認証',command=self.system) # ボタンの作成
        button2 = ttk.Button(but_frame,text='終了',command=self.root.destroy)
        but_frame.pack(expand=True)
        button1.pack(side='left',expand=True,anchor='center')
        button2.pack(side='left',expand=True,anchor='center')
                                 
    def system(self):
        # 認証できてるかのシステム
        User=self.user_name.get() # 入力欄から値を取得
        Pass=self.pass_word.get() # 入力欄から値を取得
        if User=='name' and Pass=='pass':
            #認証できたときの処理
            self.check = True
            self.root.destroy()
        else:
            self.error.pack(expand=True)
            self.user_name.delete(0,tkinter.END) # 入力欄の初期化
            self.pass_word.delete(0,tkinter.END) # 入力欄の初期化



class MainFunction(tkinter.Frame):
    def __init__(self,root):
        super().__init__(root)
        self.root = root
        self.s = ttk.Style()
        self.s.theme_use('vista')
        
        # エラーメッセージ
        self.error_mes1 = None
        self.error_mes2 = None
        
        # failnameの保存
        self.pdffile = ''
        self.textfile = ''
        self.wordfile = ''
        self.excelfile = ''
        self.outfile = ''
        
        # pdfから取得したテキストの保存先
        self.textlist = {}
        
        self.wordtitle1 = None 
        self.wordtitle2 = None
        self.group = None
        
        # 学籍番号、名前の保存先
        self.member = []
        self.member_num = []
        
        # word表紙記入項目
        self.contentlist = ['実験コース名：','実験タイトル：','グループ名：','メンバー1：','メンバー2：','メンバー3：','メンバー1の学籍番号：','メンバー2の学籍番号：','メンバー3の学籍番号：']
        
        # 記入項目の保存先
        self.worddata1 = []
        self.worddata2 = []
        
        self.pack()
        self.pdfopration()
        self.wordopration()
        self.excelopration()
        self.close()
        
        
    def pdfopration(self):
        def pdfbutton1():
            # ファイル取得機能
            self.pdffile = filedialog.askopenfilename()
            if self.pdffile[-4:] == '.pdf': # pdfファイルかどうかの確認
                pdffile.insert(0,self.pdffile)
                pdf_button2['state'] = 'normal' # ボタンクリックできる状態
                pdf_button1['state'] = 'disabled' # ボタンクリックできない状態
                self.error_mes1.pack_forget()
            else:
                self.error_mes1.pack(expand=True)
                
        def pdfbutton2():
            pdf_button2['state'] = 'disabled'            
            pdf_button3['state'] = 'normal'
            pdf_frame_top.pack_forget()
            pdf_mes2.pack(expand=True)
            outputfile.pack(side='left',expand=True)
            pdf_button3.pack(side='left',expand=True)
        
        def pdfbutton3():
            #pdfテキスト読み込みと保存
            self.textfile = outputfile.get()+'.txt' #保存先のtextファイル
            pdf_system = Pdfoperation()
            pdf_system.readpdf(self.pdffile,self.textfile)
            pdf_system.filemodify('./'+self.textfile)
            pdf_system.convert()
            self.textlist = pdf_system.paragraphlist
            
            pdf_frame_bot.pack_forget()
            pdf_mes3.pack(expand=True)
            
        # コンポーネントの作成
        pdf_frame = ttk.Frame(self.root)
        pdf_frame_top = ttk.Frame(pdf_frame)
        pdf_frame_bot = ttk.Frame(pdf_frame)
        pdf_button1 = ttk.Button(pdf_frame_top,text='pdf参照',command=pdfbutton1)
        pdf_button2 = ttk.Button(pdf_frame_top,text='実行',state='disabled',command=pdfbutton2)
        pdf_button3 = ttk.Button(pdf_frame_bot,text='保存&実行',state='disabled',command=pdfbutton3)
        pdffile = ttk.Entry(pdf_frame_top,width=50)
        outputfile = ttk.Entry(pdf_frame_bot,width=53)
        pdf_mes1 = tkinter.Message(pdf_frame_top,text='pdf読み取り',width=200)
        pdf_mes2 = tkinter.Message(pdf_frame_bot,text='出力先ファイル名を入力してください',width=200,pady=10)
        pdf_mes3 = tkinter.Message(pdf_frame,text='succcesfully',width=200,pady=10,font=('',24))
        self.error_mes1 = tkinter.Message(pdf_frame_bot,text='pdfファイルを指定してください。',width=200,foreground = 'red')
        
        # コンポーネントの配置
        pdf_frame.pack(expand=True)
        pdf_frame_top.pack(expand=True)
        pdf_frame_bot.pack(expand=True)
        pdf_mes1.pack(expand=True)
        pdffile.pack(side='left',expand=True)
        pdf_button1.pack(side='left',expand=True)
        pdf_button2.pack(side='left',expand=True)
        pdf_mes2.pack_forget()
        outputfile.pack_forget()
        pdf_button3.pack_forget()
        

    def wordopration(self):        
        def word_button1():
            # 記入欄からデータ保存
            self.wordtitle1 = self.worddata1[0].get() 
            self.wordtitle2 = self.worddata1[1].get()
            self.group = self.worddata1[2].get()
            self.member = self.worddata1[3:6]
            self.member_num = self.worddata1[6:]
            word_frame_top.pack_forget()
            word_frame_mid.pack_forget()
            word_button1.pack_forget()
            word_mes2.pack(expand=True)
            wordfile.pack(side='left',expand=True)
            word_button2.pack(side='left',expand=True)
            

        def word_button2():
            # wordに書き込み処理
            self.wordfile = wordfile.get()+'.docx' #保存先のwordファイル
            
            # 表紙の作成
            word_system = Wordopration()
            word_system.add_title(self.wordtitle1,24)
            word_system.add_title(self.wordtitle2,18,True)
            date = datetime.datetime.now()
            year=date.year 
            month=date.month
            day=date.day
            word_system.add_date(f'''
            実験開始日　　　　{year}年　　　{month}月　　　{day}日
            実験終了日　　　　{year}年　　　{month}月　　　{day}日
            報告書提出日　　　{year}年　　　{month}月　　　{day}日
            ''')
                
            word_system.add_table(self.group,self.member,self.member_num)
            
            #pdfから読み取ったtextの書き込み
            for i in self.textlist:
                for num,j in enumerate(self.textlist[i]):
                    if num==0:
                        word_system.new_document.add_page_break()
                        word_system.add_text(j,kaigyou=True)
                    else:
                        word_system.add_text(j)
                        
            word_system.new_document.save('./'+self.wordfile)
            
            word_frame_bot.pack_forget()
            word_mes3.pack(expand=True)
        
        # コンポーネントの作成
        word_frame = ttk.Frame(self.root)
        word_frame_top = ttk.Frame(word_frame)
        word_frame_mid = ttk.Frame(word_frame)
        word_frame_mid_L = ttk.Frame(word_frame_mid)
        word_frame_mid_R = ttk.Frame(word_frame_mid)
        word_frame_bot = ttk.Frame(word_frame)
        word_button1 = ttk.Button(word_frame_bot,text='保存',command=word_button1)
        word_button2 = ttk.Button(word_frame_bot,text='保存&実行',command=word_button2)
        wordfile = ttk.Entry(word_frame_bot,width=50)
        word_mes1 = tkinter.Message(word_frame_top,text='文書作成',width=200)
        word_mes2 = tkinter.Message(word_frame_bot,text='保存先ファイル名を入力してください',width=200)
        word_mes3 = tkinter.Message(word_frame,text='succcesfully',width=200,pady=10,font=('',24))
       
        # コンポーネントの配置
        word_frame.pack(expand=True)
        word_frame_top.pack(expand=True)
        word_frame_mid.pack(expand=True)
        word_frame_mid_L.pack(side='left',expand=True,anchor='center')
        word_frame_mid_R.pack(side='left',expand=True,anchor='center') 
        word_frame_bot.pack(expand=True)
        
        # 表紙の記入項目の入力欄の作成と配置  
        for num,i in enumerate(self.contentlist):
            content_name = tkinter.Message(word_frame_mid_L,text=i,width=200)
            content = ttk.Entry(word_frame_mid_R,width=20)
            content_name.pack(pady=1,expand=True)
            content.pack(pady=2,expand=True)
            self.worddata1.append(content)
            self.worddata2.append(content_name)
            
        word_mes1.pack(expand=True)
        word_button1.pack(expand=True,pady=10)
        
    def excelopration(self):
        def excelbutton1():
            # ファイル取得機能
            self.excelfile = filedialog.askopenfilename()
            if self.excelfile[-5:] == '.xlsx': # excelファイルかどうかの確認
                excelfile.insert(0,self.excelfile)
                excel_button2['state'] = 'normal'
                excel_button1['state'] = 'disabled'
                self.error_mes2.pack_forget()
            else:
                self.error_mes2.pack(expand=True)
                
        def excelbutton2():
            excel_frame_top.pack_forget()
            excel_mes2.pack(expand=True)
            outputfile.pack(side='left',expand=True)
            excel_button3.pack(side='left',expand=True)
            
        def excelbutton3():
            self.outfile = outputfile.get()+'.xlsx' #保存先のexcelファイル
            excel_frame_mid.pack_forget()
            excel_mes3.pack(expand=True)
            excel_rb1.pack(side='left',padx=5)
            excel_rb2.pack(side='left',padx=5)
            excel_rb3.pack(side='left',padx=5)
            excel_rb4.pack(side='left',padx=5)
        
        def excelbutton4(trend):
            #近似タイプを選択し、グラフ作成
            excel_system = Excelopration()
            excel_system.excel(self.excelfile,self.outfile,trend)
            excel_frame_bot.pack_forget()
            excel_mes4.pack(expand=True)
            
        # コンポーネントの作成      
        excel_frame = ttk.Frame(self.root)
        excel_frame_top = ttk.Frame(excel_frame)
        excel_frame_mid = ttk.Frame(excel_frame)
        excel_frame_bot = ttk.Frame(excel_frame)
        excel_button1 = ttk.Button(excel_frame_top,text='excel参照',command=excelbutton1)
        excel_button2 = ttk.Button(excel_frame_top,text='実行',state='disabled',command=excelbutton2)
        excel_button3 = ttk.Button(excel_frame_mid,text='保存&実行',command=excelbutton3)
        excelfile = ttk.Entry(excel_frame_top,width=50)
        outputfile = ttk.Entry(excel_frame_mid,width=53)
        excel_mes1 = tkinter.Message(excel_frame_top,text='excel読み取り',width=200)
        excel_mes2 = tkinter.Message(excel_frame_mid,text='出力先ファイル名を入力してください',width=200)
        excel_mes3 = tkinter.Message(excel_frame_bot,text='近似タイプを選択してください',width=200)
        excel_mes4 = tkinter.Message(excel_frame,text='succcesfully',width=200,pady=10,font=('',24))
        self.error_mes2 = tkinter.Message(excel_frame_mid,text='excelファイルを指定してください。',width=200,foreground = 'red')
        excel_rb1 = ttk.Radiobutton(excel_frame_bot,text='None',value='None',command=lambda: excelbutton4('None'))
        excel_rb2 = ttk.Radiobutton(excel_frame_bot,text='exponential',value='exponential',command=lambda: excelbutton4('exponential'))
        excel_rb3 = ttk.Radiobutton(excel_frame_bot,text='linear',value='linear',command=lambda: excelbutton4('linear'))
        excel_rb4 = ttk.Radiobutton(excel_frame_bot,text='log',value='log',command=lambda: excelbutton4('log'))
        
        # コンポーネントの配置
        excel_frame.pack(expand=True)
        excel_frame_top.pack(expand=True)
        excel_frame_mid.pack(expand=True)
        excel_frame_bot.pack(expand=True)
        excel_mes1.pack(expand=True)
        excelfile.pack(side='left',expand=True)
        excel_button1.pack(side='left',expand=True)
        excel_button2.pack(side='left',expand=True)
        
    def close(self):
        close_frame = ttk.Frame(self.root)
        button = ttk.Button(close_frame,text='終了',command=self.root.destroy)
        close_frame.pack(expand=True)
        button.pack(expand=True)
