"""
GUI関連の共通ユーティリティ関数

このモジュールは rename_mp3.py と write_mp3_tags.py で共通して使用される
GUI機能を提供します。
"""

# tkinterのインポートを試行し、失敗した場合はCUIモードで動作
try:
    import tkinter as tk
    from tkinter import ttk
    from tkinter import messagebox

    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False


def sortby(tree, col, descending):
    """Treeviewの並べ替えを行う"""
    data = [(tree.set(child, col), child) for child in tree.get_children("")]
    data.sort(reverse=descending)
    for ix, item in enumerate(data):
        tree.move(item[1], "", ix)
    tree.heading(col, command=lambda: sortby(tree, col, int(not descending)))


def on_double_click(event):
    """セルをダブルクリックで部分選択とコピーを可能にする"""
    item_id = event.widget.identify_row(event.y)
    column = event.widget.identify_column(event.x)
    value = event.widget.item(item_id, "values")[int(column[1:]) - 1]
    show_copy_popup(value)


def show_copy_popup(value):
    """部分選択とコピーのためのポップアップを表示"""
    popup = tk.Toplevel()
    popup.title("部分選択とコピー")
    text = tk.Text(popup, wrap="word")
    text.insert("1.0", value)
    text.pack(expand=True, fill="both")
    text.bind("<Control-c>", lambda e: popup.clipboard_append(text.selection_get()))
    close_button = ttk.Button(popup, text="閉じる", command=popup.destroy)
    close_button.pack()


def setup_window_size(root, scale_factor=1.5):
    """ウィンドウサイズを設定する"""
    default_width = 800
    default_height = 600
    window_width = int(default_width * scale_factor)
    window_height = int(default_height * scale_factor)
    root.geometry(f"{window_width}x{window_height}")


def setup_scrollbar_for_tree(frame, tree):
    """Treeview用のスクロールバーを設定する"""
    scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=0, column=1, sticky=tk.N + tk.S)
    return scrollbar


def setup_button_frame(root, row, buttons_config):
    """
    ボタンフレームを設定する

    Args:
        root: 親ウィンドウ
        row: 配置する行
        buttons_config: [(text, command), ...] のリスト

    Returns:
        作成されたボタンフレーム
    """
    button_frame = ttk.Frame(root, padding=10)
    button_frame.grid(row=row, column=0, sticky=tk.W + tk.E + tk.N + tk.S)

    for i, (text, command) in enumerate(buttons_config):
        button = ttk.Button(button_frame, text=text, command=command)
        button.grid(row=0, column=i, padx=5, pady=5)

    return button_frame


def setup_frame_resize_behavior(frame, tree):
    """フレームとTreeviewのリサイズ動作を設定する"""
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)
    tree.columnconfigure(0, weight=1)
    tree.columnconfigure(1, weight=1)
