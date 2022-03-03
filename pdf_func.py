# pdfテキスト解析するモジュール
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams


class Pdfoperation:
    def __init__(self):
        
        # 全角と半角に対応のため二つ用意
        self.paragraph1 = ['１．','２．','３．','４．','５．','６．','７．']
        self.paragraph2 = ['1 ','2 ','3 ','4 ','5 ','6 ','7 ']
        
        # テキストをparagraphごとにわけた保存先
        self.paragraphlist = {
            'title': [],
            'text1': [],
            'text2': [],
            'text3': [],
            'text4': [],
            'text5': [],
            'extra': [],
        }

    def readpdf(self,pdffile,filename):
        
        # 標準組込み関数open()でモード指定をbinaryでFileオブジェクトを取得
        fp = open(pdffile,'rb')
        # 出力先をfileオブジェクトにする
        outfp = open(filename, 'w', encoding='utf-8')
        rmgr = PDFResourceManager() # PDFResourceManagerオブジェクトの取得
        lprms = LAParams()          # LAParamsオブジェクトの取得
        device = TextConverter(rmgr, outfp, laparams=lprms)    # TextConverterオブジェクトの取得
        iprtr = PDFPageInterpreter(rmgr, device) # PDFPageInterpreterオブジェクトの取得

        # PDFファイルから1ページずつ解析(テキスト抽出)処理する
        for page in PDFPage.get_pages(fp):
            iprtr.process_page(page)
            
        outfp.close()
        device.close() # TextConverterオブジェクトの解放
        fp.close()     


    def filemodify(self,file_path):
        # paragraphごとにテキストを保存
        with open(file_path,encoding='utf-8') as f:
            row_num=0
            for row in f:
                row_shu = row.strip()
                if row_shu == '' or len(row_shu)<4:
                    continue
                if (row_shu[:2] in self.paragraph1[0]) or (row_shu[:2] in self.paragraph2[0]):
                    self.paragraph1.pop(0)
                    self.paragraph2.pop(0)
                    row_num += 1
                if row_num == 0:
                    self.paragraphlist['title'].append(row_shu)
                    continue

                self.paragraphlist['text'+str(row_num)].append(row_shu)

    def convert(self):
        # 規定テキストに加工、修正箇所の指定
        for i in self.paragraphlist:
            for num,content in enumerate(self.paragraphlist[i]):
                new_content = content.replace('　',' ').replace('。','.').replace('、',',').replace('こと.','.').replace('\x0c',' ').replace('cid:','{{修正}}')
                self.paragraphlist[i][num] = new_content
                
                
                
