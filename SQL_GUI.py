import tkinter as tk
from tkinter import messagebox, ttk
import pyodbc
from datetime import datetime
from tkcalendar import DateEntry

def connect_db():
    server = 'NTC18'
    database = 'TP'
    username = 'sa'
    password = 'Ntc002611'
    conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                          f'SERVER={server};'
                          f'DATABASE={database};'
                          f'UID={username};'
                          f'PWD={password}')
    return conn

def submit_data():
    conn = connect_db()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO TP_2024 (測定日, 測定者, グレード, ロット番号, 枝番, 枝番2, 本数, 
                               [厚さ(mm)], [幅(mm)], [長さ(mm)], [重量(g)], [電圧(mv)], 
                               SH1, SH2, SH3, SH4, SH5, 
                               [破壊荷重(Kgf)], [灰分(%)], 灰化色, 備考) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            entry_date.get(), entry_measurement_by.get(), entry_grade.get(), entry_lot_number.get(), 
            entry_branch.get(), entry_branch2.get(), entry_num.get(), entry_thickness.get(), 
            entry_width.get(), entry_length.get(), entry_weight.get(), 
            entry_voltage.get(), entry_sh1.get(), entry_sh2.get(), entry_sh3.get(), 
            entry_sh4.get(), entry_sh5.get(), entry_break_load.get(), 
            entry_ash.get(), entry_ash_color.get(), entry_remarks.get()
        ))
        conn.commit()
        messagebox.showinfo("成功", "データが登録されました")
        refresh_data_list()
    except Exception as e:
        messagebox.showerror("エラー", str(e))
    finally:
        cursor.close()
        conn.close()
    clear_entries()

def refresh_data_list(sort_by='測定日', filter_value=None):
    conn = connect_db()
    cursor = conn.cursor()
    
    try:
        # SQLクエリの更新
        query = f"""
            SELECT 
                [測定日], 
                MAX([測定者]) AS 測定者, 
                MAX([グレード]) AS グレード, 
                MAX([ロット番号]) AS ロット番号, 
                MAX([枝番]) AS 枝番, 
                MAX([枝番2]) AS 枝番2, 
                MAX([本数]) AS 本数, 
                AVG([厚さ(mm)]) AS 平均厚さ, 
                AVG([幅(mm)]) AS 平均幅, 
                AVG([長さ(mm)]) AS 平均長さ, 
                AVG([重量(g)]) AS 平均重量, 
                AVG([BD]) AS 平均BD, 
                AVG([電圧(mv)]) AS 平均電圧, 
                AVG([SR(μΩｍ)]) AS 平均SR, 
                AVG([SH1]) AS 平均SH1, 
                AVG([SH2]) AS 平均SH2, 
                AVG([SH3]) AS 平均SH3, 
                AVG([SH4]) AS 平均SH4, 
                AVG([SH5]) AS 平均SH5, 
                AVG([SH(平均)]) AS 平均SH, 
                AVG([破壊荷重(Kgf)]) AS 平均破壊荷重, 
                CAST(AVG([BS(MPa)]) AS FLOAT) AS [平均BS], 
                MAX([灰分(%)]) AS 平均灰分, 
                MAX([灰化色]) AS 灰化色, 
                MAX([備考]) AS 備考
            FROM TP_2024
        """
        
        # フィルタ条件の追加
        if filter_value:
            query += f" WHERE [測定者] LIKE '%{filter_value}%'"
        
        # GROUP BY句の修正
        query += f"""
            GROUP BY 
                [測定日], 
                [測定者], 
                [グレード], 
                [ロット番号], 
                [枝番], 
                [枝番2]
        """
        
        # ソート条件の追加
        query += f" ORDER BY [{sort_by}] "
        
        cursor.execute(query)
        rows = cursor.fetchall()

        # データツリーの更新
        for item in data_tree.get_children():
            data_tree.delete(item)

        for row in rows:
            formatted_row = (
                row[0].strftime('%Y-%m-%d'),  # 日付のフォーマット
                row[1],  # 測定者
                row[2],  # グレード
                row[3],  # ロット番号
                row[4],  # 枝番
                row[5],  # 枝番2
                f"{row[6]:.0f}",  # 本数
                f"{row[7]:.3f}",  # 平均厚さ
                f"{row[8]:.3f}",  # 平均幅
                f"{row[9]:.3f}",  # 平均長さ
                f"{row[10]:.4f}",  # 平均重量
                f"{row[11]:.2f}",  # 平均BD
                f"{row[12]:.2f}",  # 平均電圧
                f"{row[13]:.2f}",  # 平均SR
                f"{row[14]:.0f}",  # 平均SH1
                f"{row[15]:.0f}",  # 平均SH2
                f"{row[16]:.0f}",  # 平均SH3
                f"{row[17]:.0f}",  # 平均SH4
                f"{row[18]:.0f}",  # 平均SH5
                f"{row[19]:.1f}",  # 平均SH
                f"{row[20]:.4f}",  # 平均破壊荷重
                f"{row[21]:.2f}",  # 平均BS
                f"{row[22]:.2f}",  # 灰分
                row[23],  # 灰化色
                row[24]   # 備考
            )
            data_tree.insert("", "end", values=formatted_row)

    finally:
        cursor.close()  # 確実にカーソルを閉じる
        conn.close()    # 確実に接続を閉じる

    # 列幅を自動調整
    auto_adjust_column_width(data_tree)

def auto_adjust_column_width(treeview):
    for col_index, col in enumerate(treeview["columns"]):
        max_width = max([len(str(treeview.item(item)["values"][col_index])) for item in treeview.get_children()])
        max_width = max(max_width, len(col))  # ヘッダーの長さも考慮
        treeview.column(col, width=max_width * 10)  # 余白を追加

def on_item_double_click(event):
    selected_items = data_tree.selection()
    if not selected_items:  # 選択されたアイテムがない場合
        return  # 処理を終了

    selected_item = selected_items[0]  # 選択したアイテムを取得
    values = data_tree.item(selected_item)['values']  # アイテムの値を取得
    
    if not values:  # もしvaluesが空だった場合
        return  # 処理を終了

    # 各エントリーに値を設定
    entry_date.delete(0, tk.END)
    entry_date.insert(0, values[0])  # 測定日
    entry_measurement_by.delete(0, tk.END)
    entry_measurement_by.insert(0, values[1])  # 測定者
    entry_grade.delete(0, tk.END)
    entry_grade.insert(0, values[2])  # グレード
    entry_lot_number.delete(0, tk.END)
    entry_lot_number.insert(0, values[3])  # ロット番号
    entry_branch.delete(0, tk.END)
    entry_branch.insert(0, values[4])  # 枝番
    entry_branch2.delete(0, tk.END)
    entry_branch2.insert(0, values[5])  # 枝番2
    entry_num.delete(0, tk.END)
    entry_num.insert(0, values[6])  # 本数
    entry_thickness.delete(0, tk.END)
    entry_thickness.insert(0, values[7])  # 厚さ
    entry_width.delete(0, tk.END)
    entry_width.insert(0, values[8])  # 幅
    entry_length.delete(0, tk.END)
    entry_length.insert(0, values[9])  # 長さ
    entry_weight.delete(0, tk.END)
    entry_weight.insert(0, values[10])  # 重量
    entry_voltage.delete(0, tk.END)
    entry_voltage.insert(0, values[12])  # 電圧
    entry_sh1.delete(0, tk.END)
    entry_sh1.insert(0, values[14])  # SH1
    entry_sh2.delete(0, tk.END)
    entry_sh2.insert(0, values[15])  # SH2
    entry_sh3.delete(0, tk.END)
    entry_sh3.insert(0, values[16])  # SH3
    entry_sh4.delete(0, tk.END)
    entry_sh4.insert(0, values[17])  # SH4
    entry_sh5.delete(0, tk.END)
    entry_sh5.insert(0, values[18])  # SH5
    entry_break_load.delete(0, tk.END)
    entry_break_load.insert(0, values[20])  # 破壊荷重
    entry_ash.delete(0, tk.END)
    entry_ash.insert(0, values[22])  # 灰分
    entry_ash_color.delete(0, tk.END)
    entry_ash_color.insert(0, values[23])  # 灰化色
    entry_remarks.delete(0, tk.END)
    entry_remarks.insert(0, values[24])  # 備考

def update_data():
    conn = connect_db()
    cursor = conn.cursor()
    try:
        selected_item = data_tree.selection()[0]
        selected_values = data_tree.item(selected_item, 'values')
        
        cursor.execute("""
            UPDATE TP_2024 
            SET 測定日=?, 測定者=?, グレード=?, ロット番号=?, 枝番=?, 枝番2=?, 本数=?, 
                [厚さ(mm)]=?, [幅(mm)]=?, [長さ(mm)]=?, [重量(g)]=?, 
                [電圧(mv)]=?, SH1=?, SH2=?, SH3=?, SH4=?, SH5=?, 
                [破壊荷重(Kgf)]=?, [灰分(%)]=?, 灰化色=?, 備考=? 
            WHERE [測定日]=? AND 測定者=? AND グレード=? AND ロット番号=? 
            AND 枝番=? AND 枝番2=?
        """, (
            entry_date.get(), entry_measurement_by.get(), entry_grade.get(), entry_lot_number.get(), 
            entry_branch.get(), entry_branch2.get(), entry_num.get(), entry_thickness.get(), 
            entry_width.get(), entry_length.get(), entry_weight.get(), 
            float(entry_voltage.get()), entry_sh1.get(), entry_sh2.get(), 
            entry_sh3.get(), entry_sh4.get(), entry_sh5.get(), 
            entry_break_load.get(), entry_ash.get(), entry_ash_color.get(), 
            entry_remarks.get(),
            *selected_values[:6]  # WHERE句の値を設定
        ))
        
        conn.commit()
        messagebox.showinfo("成功", "データが更新されました")
        refresh_data_list()
    except Exception as e:
        messagebox.showerror("エラー", str(e))
    finally:
        cursor.close()
        conn.close()
    clear_entries()

def search_data():
    conn = connect_db()
    cursor = conn.cursor()
    try:
        # 表示したい列の順番と集計方法を指定
        query = """
        SELECT 
            [測定日], 
            MAX([測定者]) AS 測定者, 
            MAX([グレード]) AS グレード, 
            MAX([ロット番号]) AS ロット番号, 
            MAX([枝番]) AS 枝番, 
            MAX([枝番2]) AS 枝番2, 
            MAX([本数]) AS 本数, 
            AVG([厚さ(mm)]) AS 平均厚さ, 
            AVG([幅(mm)]) AS 平均幅, 
            AVG([長さ(mm)]) AS 平均長さ, 
            AVG([重量(g)]) AS 平均重量, 
            AVG([BD]) AS 平均BD, 
            AVG([電圧(mv)]) AS 平均電圧, 
            AVG([SR(μΩｍ)]) AS 平均SR, 
            AVG([SH1]) AS 平均SH1, 
            AVG([SH2]) AS 平均SH2, 
            AVG([SH3]) AS 平均SH3, 
            AVG([SH4]) AS 平均SH4, 
            AVG([SH5]) AS 平均SH5, 
            AVG([SH(平均)]) AS 平均SH, 
            AVG([破壊荷重(Kgf)]) AS 平均破壊荷重, 
            CAST(AVG([BS(MPa)]) AS FLOAT) AS 平均BS, 
            MAX([灰分(%)]) AS 平均灰分, 
            MAX([灰化色]) AS 灰化色, 
            MAX([備考]) AS 備考
        FROM TP_2024
        WHERE 1=1
        """
        params = []

        # 検索条件を追加
        if search_date.get():
            query += " AND 測定日=?"
            params.append(search_date.get())
        if search_grade.get():
            query += " AND グレード=?"
            params.append(search_grade.get())
        if search_lot_number.get():
            query += " AND ロット番号=?"
            params.append(search_lot_number.get())

        # グループ化を追加
        query += " GROUP BY [測定日]"

        cursor.execute(query, params)
        rows = cursor.fetchall()
        
        # 既存のデータをクリア
        for item in data_tree.get_children():
            data_tree.delete(item)
        
        # 新しいデータを挿入
        for row in rows:
            formatted_row = (
                row[0].strftime('%Y-%m-%d'),  # 測定日
                row[1],  # 測定者
                row[2],  # グレード
                row[3],  # ロット番号
                row[4],  # 枝番
                row[5],  # 枝番2
                f"{float(row[6]):.0f}" if isinstance(row[6], (int, float)) else row[6],  # 本数
                f"{float(row[7]):.3f}" if isinstance(row[7], (int, float)) else row[7],  # 平均厚さ
                f"{float(row[8]):.3f}" if isinstance(row[8], (int, float)) else row[8],  # 平均幅
                f"{float(row[9]):.3f}" if isinstance(row[9], (int, float)) else row[9],  # 平均長さ
                f"{float(row[10]):.4f}" if isinstance(row[10], (int, float)) else row[10],  # 平均重量
                f"{float(row[11]):.2f}" if isinstance(row[11], (int, float)) else row[11],  # 平均BD
                f"{float(row[12]):.2f}" if isinstance(row[12], (int, float)) else row[12],  # 平均電圧
                f"{float(row[13]):.2f}" if isinstance(row[13], (int, float)) else row[13],  # 平均SR
                f"{float(row[14]):.0f}" if isinstance(row[14], (int, float)) else row[14],  # 平均SH1
                f"{float(row[15]):.0f}" if isinstance(row[15], (int, float)) else row[15],  # 平均SH2
                f"{float(row[16]):.0f}" if isinstance(row[16], (int, float)) else row[16],  # 平均SH3
                f"{float(row[17]):.0f}" if isinstance(row[17], (int, float)) else row[17],  # 平均SH4
                f"{float(row[18]):.0f}" if isinstance(row[18], (int, float)) else row[18],  # 平均SH5
                f"{float(row[19]):.0f}" if isinstance(row[19], (int, float)) else row[19],  # 平均SH
                f"{float(row[20]):.4f}" if isinstance(row[20], (int, float)) else row[20],  # 平均破壊荷重
                f"{float(row[21]):.2f}" if isinstance(row[21], (int, float)) else row[21],  # 平均BS
                f"{float(row[22]):.2f}" if isinstance(row[22], (int, float)) else row[22],  # 平均灰分
                row[23],  # 灰化色
                row[24]   # 備考
            )
            data_tree.insert("", "end", values=formatted_row)
        
        # messagebox.showinfo("成功", "データが検索されました")
    except Exception as e:
        messagebox.showerror("エラー", str(e))
    finally:
        cursor.close()
        conn.close()

def clear_entries():
    entry_date.delete(0, tk.END)
    entry_measurement_by.delete(0, tk.END)
    entry_grade.delete(0, tk.END)
    entry_lot_number.delete(0, tk.END)
    entry_branch.delete(0, tk.END)
    entry_branch2.delete(0, tk.END)
    entry_num.delete(0, tk.END)
    entry_thickness.delete(0, tk.END)
    entry_width.delete(0, tk.END)
    entry_length.delete(0, tk.END)
    entry_weight.delete(0, tk.END)
    entry_voltage.delete(0, tk.END)
    entry_sh1.delete(0, tk.END)
    entry_sh2.delete(0, tk.END)
    entry_sh3.delete(0, tk.END)
    entry_sh4.delete(0, tk.END)
    entry_sh5.delete(0, tk.END)
    entry_break_load.delete(0, tk.END)
    entry_ash.delete(0, tk.END)
    entry_ash_color.delete(0, tk.END)
    entry_remarks.delete(0, tk.END)
    
def change_sort_option(*args):
    sort_by = sort_var.get()
    refresh_data_list(sort_by=sort_by, filter_value=filter_var.get())

def apply_filter(*args):
    filter_value = filter_var.get()
    refresh_data_list(sort_by=sort_var.get(), filter_value=filter_value)

sort_order = {}

# ソート関数
def sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    l.sort(reverse=reverse)

    # リストを並び替えた順にTreeviewを更新
    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)

    # ソート順を反転させる
    sort_order[col] = not reverse

def setup_treeview(tree):
    columns = tree["columns"]

    for col in columns:
        tree.heading(col, text=col, command=lambda c=col: sort_column(tree, c, sort_order.get(c, False)))


root = tk.Tk()
root.title('データ登録アプリ')
root.geometry('1024x768')

label_date = tk.Label(root, text='測定日:')
label_date.place(x=40, y=10)

# スタイルの設定
style = ttk.Style(root)
style.theme_use('clam')  # 使用するテーマを設定

# Entryの代わりにDateEntryを使用し、配色を指定する
entry_date = DateEntry(root, width=15,  borderwidth=2, background='#7f7fff', 
                       foreground='black', 
                       headersbackground='#7fbfff', 
                       headersforeground='black', 
                       weekendbackground='#ff7fbf', 
                       weekendforeground='darkred', 
                       othermonthbackground='#bcffdd', 
                       othermonthforeground='#a0a0a0',
                       relief='sunken', # 境界線のテーマ
                       font=('Meiryo', 8)) # フォントの指定

entry_date.place(x=90, y=10) # 日付入力の場所指定

# 測定者ラベルとエントリ
label_measurement_by = tk.Label(root, text='測定者:')
label_measurement_by.place(x=250, y=10)
entry_measurement_by = tk.Entry(root, width=15)
entry_measurement_by.place(x=300, y=10)

# グレードラベルとエントリ
label_grade = tk.Label(root, text='グレード:')
label_grade.place(x=10, y=50)
entry_grade = tk.Entry(root, width=15)
entry_grade.place(x=60, y=50)

# ロット番号ラベルとエントリ
label_lot_number = tk.Label(root, text='ロット番号:')
label_lot_number.place(x=170, y=50)
entry_lot_number = tk.Entry(root, width=15)
entry_lot_number.place(x=235, y=50)

label_branch = tk.Label(root, text='枝番:')
label_branch.place(x=340, y=50)
entry_branch = tk.Entry(root, width=15)
entry_branch.place(x=380, y=50)

label_branch2 = tk.Label(root, text='枝番2:')
label_branch2.place(x=490, y=50)
entry_branch2 = tk.Entry(root, width=15)
entry_branch2.place(x=535, y=50)

label_num = tk.Label(root, text='本数:')
label_num.place(x=640, y=50)
entry_num = tk.Entry(root, width=15)
entry_num.place(x=680, y=50)

label_thickness = tk.Label(root, text='厚さ(mm):')
label_thickness.place(x=10, y=80)
entry_thickness = tk.Entry(root, width=15)
entry_thickness.place(x=70, y=80)

label_width = tk.Label(root, text='幅(mm):')
label_width.place(x=180, y=80)
entry_width = tk.Entry(root, width=15)
entry_width.place(x=230, y=80)

label_length = tk.Label(root, text='長さ(mm):')
label_length.place(x=340, y=80)
entry_length = tk.Entry(root, width=15)
entry_length.place(x=400, y=80)

label_weight = tk.Label(root, text='重量(g):')
label_weight.place(x=510, y=80)
entry_weight = tk.Entry(root, width=15)
entry_weight.place(x=560, y=80)

label_voltage = tk.Label(root, text='電圧(mv):')
label_voltage.place(x=670, y=80)
entry_voltage = tk.Entry(root, width=15)
entry_voltage.place(x=730, y=80)

label_sh1 = tk.Label(root, text='SH1:')
label_sh1.place(x=10, y=110)
entry_sh1 = tk.Entry(root, width=15)
entry_sh1.place(x=45, y=110)

label_sh2 = tk.Label(root, text='SH2:')
label_sh2.place(x=150, y=110)
entry_sh2 = tk.Entry(root, width=15)
entry_sh2.place(x=185, y=110)

label_sh3 = tk.Label(root, text='SH3:')
label_sh3.place(x=290, y=110)
entry_sh3 = tk.Entry(root, width=15)
entry_sh3.place(x=325, y=110)

label_sh4 = tk.Label(root, text='SH4:')
label_sh4.place(x=430, y=110)
entry_sh4 = tk.Entry(root, width=15)
entry_sh4.place(x=465, y=110)

label_sh5 = tk.Label(root, text='SH5:')
label_sh5.place(x=570, y=110)
entry_sh5 = tk.Entry(root, width=15)
entry_sh5.place(x=605, y=110)

label_break_load = tk.Label(root, text='破壊荷重(Kgf):')
label_break_load.place(x=10, y=140)
entry_break_load = tk.Entry(root, width=15)
entry_break_load.place(x=95, y=140)

label_ash = tk.Label(root, text='灰分(%):')
label_ash.place(x=210, y=140)
entry_ash = tk.Entry(root, width=15)
entry_ash.place(x=265, y=140)

label_ash_color = tk.Label(root, text='灰化色:')
label_ash_color.place(x=380, y=140)
entry_ash_color = tk.Entry(root, width=15)
entry_ash_color.place(x=430, y=140)

label_remarks = tk.Label(root, text='備考:')
label_remarks.place(x=560, y=140)
entry_remarks = tk.Entry(root, width=15)
entry_remarks.place(x=600, y=140)

submit_button = tk.Button(root, text='登録', command=submit_data, background='#7fffbf')
submit_button.place(x=10, y=180)

# 更新ボタンを追加
update_button = tk.Button(root, text='更新', command=update_data, background='#ffff7f')
update_button.place(x=60, y=180)  # 登録ボタンの隣に配置

# 検索処理
search_date = tk.Label(root, text='データ検索')
search_date.place(x=140, y=182)

# フレームを作成
frame = tk.Frame(root, bd=2, relief="groove")
frame.place(x=200, y=170, width=620, height=45)

# フレーム内にウィジェットを配置
search_date = tk.Label(frame, text='測定日:')
search_date.place(x=10, y=10)

search_date = tk.Entry(frame, width=15)
search_date.place(x=70, y=10)

label_grade = tk.Label(frame, text='グレード:')
label_grade.place(x=200, y=10)
search_grade = tk.Entry(frame, width=15)
search_grade.place(x=260, y=10)

label_lot_number = tk.Label(frame, text='ロット番号:')
label_lot_number.place(x=370, y=10)
search_lot_number = tk.Entry(frame, width=15)
search_lot_number.place(x=440, y=10)

# 検索ボタンを追加
search_button = tk.Button(root, text='検索', command=search_data, background='#7fbfff')
search_button.place(x=770, y=180)  # 登録ボタンの隣に配置

# 初期化ボタンを追加
search_button = tk.Button(root, text='初期化', command=refresh_data_list, background='#7f7fff')
search_button.place(x=850, y=180)  

# データリスト表示部分
columns = ('測定日', '測定者', 'グレード', 'ロット番号', '枝番', '枝番2', '本数', '厚さ(mm)', '幅(mm)', 
           '長さ(mm)', '重量(g)', 'BD', '電圧(mv)', 'SR(μΩｍ)', 'SH1', 'SH2', 'SH3', 'SH4', 'SH5', 
           'SH(平均)', '破壊荷重(Kgf)', 'BS(MPa)', '灰分(%)', '灰化色', '備考')

data_tree = ttk.Treeview(root, columns=columns, show='headings', height=20) # relheightが設定されているとheightは無効
data_tree.place(x=10, y=230, relwidth=0.98, relheight=0.74) # データリスト表示サイズ調整
sortable_columns = ['測定日', '測定者', 'グレード', 'ロット番号']  # ソート可能な列を指定

for col in columns:
    if col in sortable_columns:
        data_tree.heading(col, text=col, command=lambda _col=col: refresh_data_list(sort_by=_col))
    else:
        data_tree.heading(col, text=col)

# スクロールバーを追加
scrollbar = ttk.Scrollbar(root, orient="vertical", command=data_tree.yview)
data_tree.configure(yscroll=scrollbar.set)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

sort_var = tk.StringVar(value='測定日')
sort_option = tk.OptionMenu(root, sort_var, '測定日', '測定者', 'グレード', 'ロット番号')
sort_option.place(x=915, y=177)
sort_var.trace_add('write', change_sort_option)

filter_var = tk.StringVar()
filter_entry = tk.Entry(root, textvariable=filter_var)
# filter_entry.grid(row=27, column=1)　#フィルター条件入力エントリ
filter_var.trace_add('write', apply_filter)

# ウィンドウサイズ変更に対応
root.grid_rowconfigure(26, weight=1)
root.grid_columnconfigure(3, weight=1)

refresh_data_list()  # データを初期ロード
auto_adjust_column_width(data_tree)  # 列幅を初期調整

# ソート機能をTreeviewに設定
setup_treeview(data_tree)
   
# ダブルクリックイベントをバインド
data_tree.bind("<Double-1>", on_item_double_click)

root.mainloop()