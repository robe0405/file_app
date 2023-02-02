import os
import pathlib
import pandas as pd
import tkinter as tk
from tkinter import ttk
import tkinter.filedialog
import tkinter.messagebox

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("700x500")
        self.root.configure(bg="white")
        self.root.title("file_app")
        # ttkのスタイル
        self.s = ttk.Style()
        self.s.theme_use('classic')
        self.s.configure('MyWidget.TButton', background='white', relief="flat")

    """機能"""
    def Form(self):
        form = tk.Entry(master=self.root, width=40)
        self.text = tk.StringVar()
        form.configure(textvariable=self.text)
        form.pack(side="top")

    def Static(self, master, text):
        static = tk.Label(master=master)
        static.configure(text=text)
        static.pack(side="top", pady=10)

    def Button(self, master, text, next, side):
        # tk.Buttonでも問題ない
        button = ttk.Button(master=master, text=text, width=20, padding=-3, command=next, style='MyWidget.TButton')
        button.pack(side=side, pady=10)

    def Frame_top(self):
        self.frame_t = tk.Frame(master=self.root, relief="groove", bd=1, width=420, height=250)
        self.frame_t.pack(side="top")
        self.frame_t.pack_propagate(0)

    def Frame_bottom(self):
        self.frame_b = tk.Frame(master=self.root, relief="groove", width=420, height=100)
        self.frame_b.pack(side="top")
        self.frame_b.pack_propagate(0)

    # ラジオボタン
    def Radiobutton(self):
        self.radio_value = tk.IntVar()
        self.radio_value.set(0)
        r1 = tk.Radiobutton(master=self.newWindow, text="縦", value=0, variable=self.radio_value)
        r1.pack()
        r2 = tk.Radiobutton(master=self.newWindow, text="横", value=1, variable=self.radio_value)
        r2.pack()

    """データの表示"""
    def View(self, master, column):
        self.tree = ttk.Treeview(master=master)# ツリービューの作成
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)# 選択された値が得られる
        self.tree["columns"] = (1)# 列インデックスの作成
        for col in column:# SELECT文で取得した各レコードを繰り返し取得
            self.tree.insert("", "end", text=col)
        self.tree.pack()# ツリービューの配置

    # よく分からないがTreeviewの仕様
    def on_tree_select(self, event):
        print("selected items:")
        self.selected_col = []
        for item in self.tree.selection():
            item_text = self.tree.item(item,"text")
            print(item_text)
            self.selected_col.append(item_text)

    """サブウィンドウ"""
    def NewWindow(self, column, filepath):
        self.newWindow = tk.Toplevel(master=self.root)
        self.newWindow.geometry("500x400+420+200")
        self.View(self.newWindow, column)
        self.Static(self.newWindow, "結合する方向をお選びください")
        self.Radiobutton()
        self.Button(self.newWindow, "実行", lambda : self.Pandas(filepath, column), "top")

    """ファイルの結合"""
    def Pandas(self, filepath, column):
        value = self.radio_value.get()

        dir_path = os.path.dirname(filepath)
        path = pathlib.Path(dir_path)
        df_list = []
        for file in path.iterdir():
            if file.match("*.xlsx"):
                df = pd.read_excel(file)
                df = df.loc[:, column]
                df_list.append(df)
        # 結合
        new_df = pd.concat(df_list, axis=value)
        # 保存
        save_path = tk.filedialog.asksaveasfilename(title="名前を付けて保存", defaultextension="xlsx")
        new_df.to_excel(save_path)
        # 終了
        self.root.destroy()


"""メインコード"""
class Main:
    def __init__(self):
        self.col_list = []

        self.app = App()
        self.app.Static(self.app.root, "結合したいフォルダ内のファイルを１つ選択してください")
        self.app.Form()
        self.app.text.set("# 選択したファイルのURL")
        self.app.Button(self.app.root, "ファイル", lambda : self.Fileselect(), "top")
        self.app.Static(self.app.root, "結合する列を選択してください")
        self.app.Frame_top()
        self.app.Static(self.app.frame_t, "列要素")
        self.app.Frame_bottom()
        self.app.Button(
                        self.app.frame_b, "決定",
                        lambda : self.app.NewWindow(self.app.selected_col, self.path), "left"
                        )
        self.app.Button(self.app.frame_b, "リセット", lambda : self.Reset(), "right")

    """ファイル選択"""
    def Fileselect(self):
        # ファイルの選択
        self.path = tk.filedialog.askopenfilename()
        # Formのテキストと連動している
        self.app.text.set(self.path)
        # pandas データ編集
        df = pd.read_excel(self.path)
        self.col_list = list(df.columns)
        self.app.View(self.app.frame_t, self.col_list)

    """リセット"""
    def Reset(self):
        self.app.text.set("# 選択したファイルのURL")
        self.app.tree.destroy()

    def Run(self):
        self.app.root.mainloop()

if __name__ == "__main__":
    main = Main()
    main.Run()
