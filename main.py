import tkinter
from gui_func import Authentication
from gui_func import MainFunction

def main():
    
    root = tkinter.Tk()
    root.title('レポート作成キット') # タイトル
    root.geometry('300x150') # windowサイズ
    Application = Authentication(root=root)
    Application.mainloop()


    if Application.check==True:
        root = tkinter.Tk()
        root.title('レポート作成キット')
        root.geometry('600x650')
        MainApplication = MainFunction(root=root)            
        MainApplication.mainloop()
        
        
if __name__ == '__main__':
    main()
