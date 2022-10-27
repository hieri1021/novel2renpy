#tkinterについて調べる場合は、https://qiita.com/nnahito/items/ad1428a30738b3d93762 がおすすめ
#tkinterdnd2について調べる場合は、https://office54.net/python/tkinter/file-drag-drop がおすすめ
#その他については、頑張って調べてください

from tkinterdnd2 import *
from tkinter import *
import docx
import os


# 注意書き
page1 = ['・これは、原稿に書いてある名前の部分を書き換え、Renpy仕様のスクリプトを作成する手助けをするものです',
         '・このウィンドウにテキストファイル(docxファイルでも可)をドラッグアンドドロップしてください',
         '・次に出て来るウィンドウに書き換える名前を打ち込んでください',
         '・名前を打ち間違えた場合は、"back"ボタンを押せばひとつ前に戻る事ができます',
         '・すべての名前を打ち込んだら、終了します',
         '・もし原稿の名前に誤字があり、その名前がウィンドウに表示された場合には、その名前と思われるものを打ち込んでください'
        ]


# 不要な空白文字を消去する関数
def delete_space(i, list):

    list[i] = list[i].replace('\n','')
    list[i] = list[i].replace('\t','')
    list[i] = list[i].replace(' ','')
    list[i] = list[i].replace('\u3000','')
    #消去する文字に不備があれば、ここに追加


# 名前を変更する関数
def change_name(i, list, name_list):
    global before_name
    
    for name in name_list: # 元の名前に対応する新しい名前を探す
        if list[i][0:len(name[0])+1] == name[0]+'「':
            list[i] = name[1] + ' "' + list[i].removeprefix(name[0]) + '"'
            before_name = name[1] # 一旦名前を保存しておく(次の行で使うことがある)
            return
            
    if list[i] == '': # 行の中身が空の時
        list[i] = '\n\n'

    elif list[i][0] == '「': # その行の話し手が書かれていないとき
        list[i] = '\n' + before_name + ' "' + list[i] + '"'
    
    else: # その他(ナレーションが話しているとき)
        list[i] = '\n' + name[-1] + ' "' + list[i] + '"'


# ファイルのディレクトリを取得する関数
def text_view(event):
    global file

    file = event.data # Entryの中身を取得
    
    root1.quit()
    root1.destroy()

# 新しい名前を取得する関数
def get():
    global name_list,page

    text = txt.get()
    name_list[page].append(text)

    page += 1

    if page >= len(name_list):
        root2.destroy()
        return
    
    canvas.delete("t")
    txt.delete(0,END)
    canvas.create_text(105,50,text=f'{name_list[page][0]}',tag="t")


# ひとつ前の名前に戻る関数
def back():
    global name_list,page

    if page > 0:
        page -= 1
        name_list[page].pop(1)
        
        canvas.delete("t")
        canvas.create_text(105,50,text=f'{name_list[page][0]}',tag="t")
    
    txt.delete(0,END)


# 名前を取得
def get_name(file):
    global buf,attention_root,attention_canvas

    # ファイルの中身を読み込む
    if file.endswith('.docx'):
        buf = []

        doc = docx.Document(file)
        for para in doc.paragraphs:
            buf.append(para.text)

    elif file.endswith('.txt'):

        with open(file,'r',encoding='utf-8') as f:
            buf = f.readlines()
              
    else:
        attention_root = Tk()
        height = attention_root.winfo_screenheight()
        width = attention_root.winfo_screenwidth()
        attention_canvas = Canvas(attention_root,width=400,height=200)
        attention_root.title('set names')
        attention_root.geometry('400x200+'+str(width//3)+'+'+str(height//3))
        attention_root.resizable(0,0)
        attention_root.config(bg='#66ffff')
        attention_canvas.create_text(200,100,text='Error: connot use this file',activefill="red",anchor=CENTER)

        attention_canvas.pack()
        attention_root.mainloop()
    
    # 変更する名前を取得する
    for j in range(len(buf)):

        delete_space(j,buf)
        
        target = '「'
        idx = buf[j].find(target)
        ori_name = buf[j][:idx]
        
        if idx != -1 and len(ori_name) > 0:
            if [ori_name] not in name_list or len(name_list) == 0:
                name_list.append([ori_name])

    name_list.append(['ナレーション'])

    return name_list


# 新しいファイルを作る
def writefile(buf, name_list):
    for i in range(len(buf)):
        # 閉じるボタンが押されたときの対処
        try:
            change_name(i,buf,name_list)
        
        except IndexError:
            return

        texts.append(buf[i])

    # ディレクトリの一部を抜粋して変数に代入
    base_file = os.path.splitext(os.path.basename(file))[0]
    dir_file = os.path.dirname(file)

    # 新しいファイルに書き込みをする
    with open(dir_file+'/edited_'+base_file+'.txt','w',encoding='utf-8') as f:
        f.writelines(texts)


# メインウィンドウの生成
def mainwindow():
    global root1,canvas,root2,txt

    root1 = TkinterDnD.Tk()
    height = root1.winfo_screenheight()
    width = root1.winfo_screenwidth()
    root1.title('drag and drop')
    root1.geometry('400x300+'+str(width//3)+'+'+str(height//3))
    root1.resizable(0,0)
    root1.config(bg='#66ffff')

    # ドラッグアンドドロップ機能の追加
    root1.drop_target_register(DND_FILES)
    root1.dnd_bind('<<Drop>>', text_view)

    canvas = Canvas(root1,width=400,height=300)
    canvas.create_text(200,40,text='*注意書き*',anchor=CENTER)

    i = 0
    for t in page1:
        canvas.create_text(40,60+i*15,text=t,anchor=NW,width=320)

        if len(t) > 35:
            i += 1
        
        i += 1

    # ウィジェットの配置
    canvas.pack()
    root1.mainloop()

    # ドラッグアンドドロップする前に閉じるボタンが押された時の対処
    try:
        name_list = get_name(file)

    except NameError:
        return

    # メインウィンドウ2の生成
    root2 = Tk()
    canvas = Canvas(root2,width=400,height=200)
    root2.title('set names')
    root2.geometry('400x200+'+str(width//3)+'+'+str(height//3))
    root2.resizable(0,0)
    root2.config(bg='#66ffff')

    lbl = Label(text='name')
    lbl.place(x=40, y=68)

    canvas.create_text(105,50,text=f'{name_list[page][0]}',tag="t")

    txt = Entry(width=20)
    txt.place(x=90, y=70)
    txt.bind("<Return>",lambda event: get())

    done_btn = Button(root2,text='Done',command=get)
    done_btn.place(x=220,y=65)

    back_btn = Button(root2,text='back',command=back)
    back_btn.place(x=260,y=65)

    canvas.pack()
    root2.mainloop()

    writefile(buf, name_list)


# main関数の実行
if __name__ == '__main__':

    texts = []
    name_list = []
    before_name = ''
    page = 0

    mainwindow()

# exeファイルを作るフォルダと同じ場所に、hook-tkinterdnd2.pyを入れておく(そうしないとexe化したときにエラーをはく)
# pyinstallerでexe化するときはターミナル(もしくはコマンドプロンプト,etc...)にて、
# "pyinstaller -F -w --additional-hooks-dir . replace.py"と打つ
# 詳しくは、https://juu7g.hatenablog.com/entry/Python/csv/viewer