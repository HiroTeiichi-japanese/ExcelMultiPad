import tkinter as tk
from tkinter import ttk
import pyautogui
import time
import threading
import pystray
from PIL import Image, ImageDraw
import pygetwindow as gw

# --- 1. 全言語のフルデータ（完全収録） ---
DATA_ALL = {
    "Japanese": {
        "menu": {"show": "表示", "lang": "言語切替", "exit": "終了"},
        "items": [
            ("✨ 基本中の基本", [("上書き保存 (Ctrl+S)", ['ctrl', 's']), ("名前を付けて保存 (F12)", ['f12']), ("元に戻す (Ctrl+Z)", ['ctrl', 'z']), ("やり直し (Ctrl+Y)", ['ctrl', 'y']), ("印刷 (Ctrl+P)", ['ctrl', 'p'])]),
            ("🎨 セルの書式設定", [("セルの書式設定 (Ctrl+1)", ['ctrl', '1']), ("桁区切り (C+S+!)", ['ctrl', 'shift', '1']), ("パーセント (C+S+%)", ['ctrl', 'shift', '5']), ("通貨形式 (C+S+$)", ['ctrl', 'shift', '4']), ("日付形式 (C+S+#)", ['ctrl', 'shift', '3']), ("外枠をつける (C+S+&)", ['ctrl', 'shift', '6']), ("外枠を消す (C+S+_)", ['ctrl', 'shift', '_']), ("太字 (C+B)", ['ctrl', 'b']), ("下線 (C+U)", ['ctrl', 'u']), ("斜体 (C+I)", ['ctrl', 'i']), ("取り消し線 (C+5)", ['ctrl', '5'])]),
            ("📋 コピペ・形式選択", [("コピー (Ctrl+C)", ['ctrl', 'c']), ("貼り付け (Ctrl+V)", ['ctrl', 'v']), ("形式を選択して貼付 (C+A+V)", ['ctrl', 'alt', 'v']), ("値のみ貼り付け (Alt,E,S,V)", ['alt', 'e', 's', 'v']), ("書式のみ貼り付け (Alt,E,S,T)", ['alt', 'e', 's', 't']), ("数式のみ貼り付け (Alt,E,S,F)", ['alt', 'e', 's', 'f'])]),
            ("➕ 行・列・セルの編集", [("セル編集 (F2)", ['f2']), ("行/列の挿入 (C+[+])", ['ctrl', 'shift', ';']), ("行/列の削除 (C+[-])", ['ctrl', '-']), ("行全体を選択 (S+Space)", ['shift', 'space']), ("列全体を選択 (C+Space)", ['ctrl', 'space']), ("行を隠す (C+9)", ['ctrl', '9']), ("列を隠す (C+0)", ['ctrl', '0']), ("上のセルをコピー (C+D)", ['ctrl', 'd']), ("左のセルをコピー (C+R)", ['ctrl', 'r'])]),
            ("🧮 数式・データ", [("オートSUM (Alt+Shift+=)", ['alt', 'shift', '=']), ("今日の日付 (Ctrl+;)", ['ctrl', ';']), ("現在の時刻 (Ctrl+:)", ['ctrl', ':']), ("フラッシュフィル (Ctrl+E)", ['ctrl', 'e']), ("再計算 (F9)", ['f9']), ("数式のみ表示切替 (Ctrl+`)", ['ctrl', '`'])]),
            ("🔍 検索・移動・フィルター", [("検索 (Ctrl+F)", ['ctrl', 'f']), ("置換 (Ctrl+H)", ['ctrl', 'h']), ("フィルターON/OFF (C+S+L)", ['ctrl', 'shift', 'l']), ("ジャンプ (Ctrl+G)", ['ctrl', 'g']), ("可視セルのみ選択 (Alt+;)", ['alt', ';']), ("表の端まで移動 (Ctrl+矢印)", ['ctrl', 'down']), ("表の端まで選択 (C+S+矢印)", ['ctrl', 'shift', 'down'])]),
            ("📑 シート・ウィンドウ", [("次のシート (C+PgDn)", ['ctrl', 'pagedown']), ("前のシート (C+PgUp)", ['ctrl', 'pageup']), ("新しいシート (S+F11)", ['shift', 'f11']), ("ウィンドウ固定 (Alt,W,F,F)", ['alt', 'w', 'f', 'f']), ("全画面表示 (Ctrl+F1)", ['ctrl', 'f1'])]),
            ("🛠️ その他便利機能", [("動作の繰り返し (F4)", ['f4']), ("名前の管理 (Ctrl+F3)", ['ctrl', 'f3']), ("コメント挿入 (S+F2)", ['shift', 'f2']), ("テーブル作成 (Ctrl+T)", ['ctrl', 't'])])
        ]
    },
    "English": {
        "menu": {"show": "Show", "lang": "Language", "exit": "Exit"},
        "items": [
            ("✨ The Basics", [("Save (Ctrl+S)", ['ctrl', 's']), ("Save As (F12)", ['f12']), ("Undo (Ctrl+Z)", ['ctrl', 'z']), ("Redo (Ctrl+Y)", ['ctrl', 'y']), ("Print (Ctrl+P)", ['ctrl', 'p'])]),
            ("🎨 Cell Formatting", [("Format Cells (Ctrl+1)", ['ctrl', '1']), ("Comma (C+S+!)", ['ctrl', 'shift', '1']), ("Percent (C+S+%)", ['ctrl', 'shift', '5']), ("Currency (C+S+$)", ['ctrl', 'shift', '4']), ("Date (C+S+#)", ['ctrl', 'shift', '3']), ("Add Border (C+S+&)", ['ctrl', 'shift', '6']), ("Remove Border (C+S+_)", ['ctrl', 'shift', '_']), ("Bold (C+B)", ['ctrl', 'b']), ("Underline (C+U)", ['ctrl', 'u']), ("Italic (C+I)", ['ctrl', 'i']), ("Strikethrough (C+5)", ['ctrl', '5'])]),
            ("📋 Copy & Paste Special", [("Copy (Ctrl+C)", ['ctrl', 'c']), ("Paste (Ctrl+V)", ['ctrl', 'v']), ("Paste Special (C+A+V)", ['ctrl', 'alt', 'v']), ("Values (Alt,E,S,V)", ['alt', 'e', 's', 'v']), ("Formats (Alt,E,S,T)", ['alt', 'e', 's', 't']), ("Formulas (Alt,E,S,F)", ['alt', 'e', 's', 'f'])]),
            ("➕ Edit Rows/Cols/Cells", [("Edit (F2)", ['f2']), ("Insert (C+[+])", ['ctrl', 'shift', ';']), ("Delete (C+[-])", ['ctrl', '-']), ("Select Row (S+Space)", ['shift', 'space']), ("Select Col (C+Space)", ['ctrl', 'space']), ("Hide Row (C+9)", ['ctrl', '9']), ("Hide Col (C+0)", ['ctrl', '0']), ("Fill Down (C+D)", ['ctrl', 'd']), ("Fill Right (C+R)", ['ctrl', 'r'])]),
            ("🧮 Data & Formulas", [("AutoSUM (Alt+S+=)", ['alt', 'shift', '=']), ("Today (Ctrl+;)", ['ctrl', ';']), ("Time (Ctrl+:)", ['ctrl', ':']), ("Flash Fill (Ctrl+E)", ['ctrl', 'e']), ("Recalculate (F9)", ['f9']), ("Toggle Formula (Ctrl+`)", ['ctrl', '`'])]),
            ("🔍 Find & Filter", [("Find (Ctrl+F)", ['ctrl', 'f']), ("Replace (Ctrl+H)", ['ctrl', 'h']), ("Filter (C+S+L)", ['ctrl', 'shift', 'l']), ("Go To (Ctrl+G)", ['ctrl', 'g']), ("Select Visible (Alt+;)", ['alt', ';']), ("To Edge (Ctrl+Arr)", ['ctrl', 'down']), ("Select Edge (C+S+Arr)", ['ctrl', 'shift', 'down'])]),
            ("📑 Sheet & Window", [("Next Sheet (C+PgDn)", ['ctrl', 'pagedown']), ("Prev Sheet (C+PgUp)", ['ctrl', 'pageup']), ("New Sheet (S+F11)", ['shift', 'f11']), ("Freeze Panes (Alt,W,F,F)", ['alt', 'w', 'f', 'f']), ("Full Screen (C+F1)", ['ctrl', 'f1'])]),
            ("🛠️ Others", [("Repeat (F4)", ['f4']), ("Name Mgr (Ctrl+F3)", ['ctrl', 'f3']), ("Comment (Shift+F2)", ['shift', 'f2']), ("Table (Ctrl+T)", ['ctrl', 't'])])
        ]
    },
    "Korean": {
        "menu": {"show": "표시", "lang": "언어 변경", "exit": "종료"},
        "items": [
            ("✨ 기본 기능", [("저장 (Ctrl+S)", ['ctrl', 's']), ("다른 이름으로 저장 (F12)", ['f12']), ("실행 취소 (Ctrl+Z)", ['ctrl', 'z']), ("다시 실행 (Ctrl+Y)", ['ctrl', 'y']), ("인쇄 (Ctrl+P)", ['ctrl', 'p'])]),
            ("🎨 셀 서식 설정", [("셀 서식 (Ctrl+1)", ['ctrl', '1']), ("쉼표 (C+S+!)", ['ctrl', 'shift', '1']), ("백분율 (C+S+%)", ['ctrl', 'shift', '5']), ("통화 (C+S+$)", ['ctrl', 'shift', '4']), ("날짜 (C+S+#)", ['ctrl', 'shift', '3']), ("테두리 추가 (C+S+&)", ['ctrl', 'shift', '6']), ("테두리 제거 (C+S+_)", ['ctrl', 'shift', '_']), ("굵게 (Ctrl+B)", ['ctrl', 'b']), ("밑줄 (Ctrl+U)", ['ctrl', 'u']), ("기울임꼴 (Ctrl+I)", ['ctrl', 'i']), ("취소선 (Ctrl+5)", ['ctrl', '5'])]),
            ("📋 복사 및 붙여넣기", [("복사 (Ctrl+C)", ['ctrl', 'c']), ("붙여넣기 (Ctrl+V)", ['ctrl', 'v']), ("선택하여 붙여넣기 (C+A+V)", ['ctrl', 'alt', 'v']), ("값 붙여넣기 (Alt,E,S,V)", ['alt', 'e', 's', 'v']), ("서식 붙여넣기 (Alt,E,S,T)", ['alt', 'e', 's', 't']), ("수식 붙여넣기 (Alt,E,S,F)", ['alt', 'e', 's', 'f'])]),
            ("➕ 행/열/セル 편집", [("셀 편집 (F2)", ['f2']), ("행/열 삽입 (C+[+])", ['ctrl', 'shift', ';']), ("행/열 삭제 (C+[-])", ['ctrl', '-']), ("행 전체 선택 (S+Space)", ['shift', 'space']), ("열 전체 선택 (C+Space)", ['ctrl', 'space']), ("행 숨기기 (Ctrl+9)", ['ctrl', '9']), ("열 숨기기 (Ctrl+0)", ['ctrl', '0']), ("아래로 채우기 (Ctrl+D)", ['ctrl', 'd']), ("오른쪽으로 채우기 (Ctrl+R)", ['ctrl', 'r'])]),
            ("🧮 데이터 및 수식", [("자동 합계 (Alt+S+=)", ['alt', 'shift', '=']), ("오늘 날짜 (Ctrl+;)", ['ctrl', ';']), ("현재 시간 (Ctrl+:)", ['ctrl', ':']), ("플래시 채우기 (Ctrl+E)", ['ctrl', 'e']), ("재계산 (F9)", ['f9']), ("수식 표시 전환 (Ctrl+`)", ['ctrl', '`'])]),
            ("🔍 검색 및 필터", [("찾기 (Ctrl+F)", ['ctrl', 'f']), ("바꾸기 (Ctrl+H)", ['ctrl', 'h']), ("필터 (C+S+L)", ['ctrl', 'shift', 'l']), ("이동 (Ctrl+G)", ['ctrl', 'g']), ("화면 셀만 선택 (Alt+;)", ['alt', ';']), ("끝으로 (Ctrl+방향)", ['ctrl', 'down']), ("끝까지 (C+S+방향)", ['ctrl', 'shift', 'down'])]),
            ("📑 시트 및 창", [("다음 시트 (C+PgDn)", ['ctrl', 'pagedown']), ("이전 시트 (C+PgUp)", ['ctrl', 'pageup']), ("새 시트 (S+F11)", ['shift', 'f11']), ("틀 고정 (Alt,W,F,F)", ['alt', 'w', 'f', 'f']), ("전체 화면 (C+F1)", ['ctrl', 'f1'])]),
            ("🛠️ 기타 기능", [("반복 (F4)", ['f4']), ("이름 관리 (Ctrl+F3)", ['ctrl', 'f3']), ("메모 삽입 (S+F2)", ['shift', 'f2']), ("표 만들기 (Ctrl+T)", ['ctrl', 't'])])
        ]
    },
    "Chinese": {
        "menu": {"show": "显示", "lang": "语言切换", "exit": "退出"},
        "items": [
            ("✨ 基础功能", [("保存 (Ctrl+S)", ['ctrl', 's']), ("另存为 (F12)", ['f12']), ("撤销 (Ctrl+Z)", ['ctrl', 'z']), ("重做 (Ctrl+Y)", ['ctrl', 'y']), ("打印 (Ctrl+P)", ['ctrl', 'p'])]),
            ("🎨 单元格格式", [("单元格格式 (Ctrl+1)", ['ctrl', '1']), ("千位分隔符 (C+S+!)", ['ctrl', 'shift', '1']), ("百分比 (C+S+%)", ['ctrl', 'shift', '5']), ("货币 (C+S+$)", ['ctrl', 'shift', '4']), ("日期格式 (C+S+#)", ['ctrl', 'shift', '3']), ("添加边框 (C+S+&)", ['ctrl', 'shift', '6']), ("清除边框 (C+S+_)", ['ctrl', 'shift', '_']), ("加粗 (Ctrl+B)", ['ctrl', 'b']), ("下划线 (Ctrl+U)", ['ctrl', 'u']), ("斜体 (Ctrl+I)", ['ctrl', 'i']), ("删除线 (Ctrl+5)", ['ctrl', '5'])]),
            ("📋 复制与选择性粘贴", [("复制 (Ctrl+C)", ['ctrl', 'c']), ("粘贴 (Ctrl+V)", ['ctrl', 'v']), ("选择性粘贴 (C+A+V)", ['ctrl', 'alt', 'v']), ("粘贴数值 (Alt,E,S,V)", ['alt', 'e', 's', 'v']), ("粘贴格式 (Alt,E,S,T)", ['alt', 'e', 's', 't']), ("粘贴公式 (Alt,E,S,F)", ['alt', 'e', 's', 'f'])]),
            ("➕ 行/列/单元格编辑", [("编辑 (F2)", ['f2']), ("插入 (C+[+])", ['ctrl', 'shift', ';']), ("删除 (C+[-])", ['ctrl', '-']), ("选中整行 (S+Space)", ['shift', 'space']), ("选中整列 (C+Space)", ['ctrl', 'space']), ("隐藏行 (Ctrl+9)", ['ctrl', '9']), ("隐藏列 (Ctrl+0)", ['ctrl', '0']), ("向下填充 (Ctrl+D)", ['ctrl', 'd']), ("向右填充 (Ctrl+R)", ['ctrl', 'r'])]),
            ("🧮 数据与公式", [("自动求和 (Alt+S+=)", ['alt', 'shift', '=']), ("输入日期 (Ctrl+;)", ['ctrl', ';']), ("输入时间 (Ctrl+:)", ['ctrl', ':']), ("快速填充 (Ctrl+E)", ['ctrl', 'e']), ("重算 (F9)", ['f9']), ("切换显示公式 (Ctrl+`)", ['ctrl', '`'])]),
            ("🔍 查找与筛选", [("查找 (Ctrl+F)", ['ctrl', 'f']), ("替换 (Ctrl+H)", ['ctrl', 'h']), ("筛选 (C+S+L)", ['ctrl', 'shift', 'l']), ("定位 (Ctrl+G)", ['ctrl', 'g']), ("选中可见单元格 (Alt+;)", ['alt', ';']), ("移至边缘 (Ctrl+方向)", ['ctrl', 'down']), ("选至边缘 (C+S+方向)", ['ctrl', 'shift', 'down'])]),
            ("📑 工作表与窗口", [("下一张表 (C+PgDn)", ['ctrl', 'pagedown']), ("上一张表 (C+PgUp)", ['ctrl', 'pageup']), ("新建表 (S+F11)", ['shift', 'f11']), ("冻结窗格 (Alt,W,F,F)", ['alt', 'w', 'f', 'f']), ("全屏显示 (C+F1)", ['ctrl', 'f1'])]),
            ("🛠️ 其他功能", [("重复操作 (F4)", ['f4']), ("名称管理 (Ctrl+F3)", ['ctrl', 'f3']), ("插入批注 (S+F2)", ['shift', 'f2']), ("创建表 (Ctrl+T)", ['ctrl', 't'])])
        ]
    }
}

# --- 2. 実行制御 ---
def press(keys):
    root.withdraw()
    time.sleep(0.1)
    excel_wins = [w for w in gw.getWindowsWithTitle('Excel') if w.visible]
    if excel_wins:
        try:
            excel_wins[0].activate()
            time.sleep(0.05)
        except: pass

    if isinstance(keys, list):
        pyautogui.hotkey(*keys)
    else:
        pyautogui.write(keys)

    time.sleep(0.1)
    root.deiconify()
    root.attributes("-topmost", True)

# --- 3. メイン画面の構築 ---
root = tk.Tk()
root.title("Excel Launcher")
root.geometry("380x880")
root.attributes("-topmost", True)

lang_bar = tk.Frame(root, bg="#eeeeee")
lang_bar.pack(fill="x")

# --- 数字キーエリア（最上部固定） ---
num_area = tk.LabelFrame(root, text="🔢 NumPad", font=("Meiryo", 9, "bold"), fg="#1D6F42")
num_area.pack(fill="x", padx=5, pady=5)

np_btns = [
    ('7', 0, 0), ('8', 0, 1), ('9', 0, 2), ('/', 0, 3),
    ('4', 1, 0), ('5', 1, 1), ('6', 1, 2), ('*', 1, 3),
    ('1', 2, 0), ('2', 2, 1), ('3', 2, 2), ('-', 2, 3),
    ('0', 3, 0), ('.', 3, 1), ('=', 3, 2), ('+', 3, 3)
]

for (text, r, c) in np_btns:
    if text == '+': k = ['shift', ';']
    elif text == '*': k = ['shift', ':']
    elif text == '/': k = ['/']
    elif text == '=': k = ['=']
    elif text == '-': k = ['-']
    else: k = text
    tk.Button(num_area, text=text, width=6, height=2, font=("Arial", 10, "bold"), 
              command=lambda v=k: press(v)).grid(row=r, column=c, padx=2, pady=2)

tk.Button(num_area, text="Enter", height=2, font=("Meiryo", 9, "bold"), 
          command=lambda: press(['enter'])).grid(row=4, column=0, columnspan=4, sticky="we", pady=2)

# --- リストエリア ---
canvas = tk.Canvas(root)
scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
scroll_inner = ttk.Frame(canvas)
scroll_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.create_window((0, 0), window=scroll_inner, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")
canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

# --- 言語切り替え・トレイメニュー ---
def update_tray_menu(lang_name):
    m = DATA_ALL[lang_name]["menu"]
    l_sub = [pystray.MenuItem(ln, (lambda n=ln: lambda: change_lang(n))()) for ln in DATA_ALL.keys()]
    icon.menu = pystray.Menu(
        pystray.MenuItem(m["show"], lambda: root.deiconify(), default=True),
        pystray.MenuItem(m["lang"], pystray.Menu(*l_sub)),
        pystray.MenuItem(m["exit"], lambda i: (i.stop(), root.destroy()))
    )

def change_lang(name):
    for w in scroll_inner.winfo_children(): w.destroy()
    data = DATA_ALL[name]
    for cat, items in data["items"]:
        lf = tk.LabelFrame(scroll_inner, text=cat, font=("Meiryo", 9, "bold"), fg="#1D6F42")
        lf.pack(fill="x", padx=5, pady=3)
        for label, keys in items:
            ttk.Button(lf, text=label, command=lambda k=keys: press(k)).pack(fill="x")
    update_tray_menu(name)

# 起動用ボタン
for ln in DATA_ALL.keys():
    tk.Button(lang_bar, text=ln, font=("Arial", 8), command=lambda n=ln: change_lang(n)).pack(side="left", padx=2)

# トレイアイコン初期化
def create_icon():
    img = Image.new('RGB', (64, 64), (33, 115, 70))
    ImageDraw.Draw(img).text((15, 20), "EX", fill=(255, 255, 255))
    return img

icon = pystray.Icon("ExcelLauncher", create_icon(), "Excel Tool")
threading.Thread(target=icon.run, daemon=True).start()

root.protocol('WM_DELETE_WINDOW', lambda: root.withdraw())
change_lang("Japanese") # 初期言語設定とともにトレイメニューも生成される
root.mainloop()