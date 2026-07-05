# ----------------------------------------------------------------------
# NowPlayingPoster v1.1.0 (2026-07-05)
#
# Windows版iTunesで現在再生中の曲情報を取得し、
# X (旧Twitter) への投稿を支援するデスクトップアプリケーション
#
# Copyright (c) 2026 HMago6
#
# Released under the MIT License.
# See LICENSE file for details.
# Author: HMago6
# GitHub: https://github.com/hmago6/NowPlayingPoster
# ----------------------------------------------------------------------
# Version History
# v1.0.0 (2026-06-28)
# - 初版公開
#
# v1.1.0 (2026-07-05)
# - ウィンドウ位置保存機能を追加
# - 起動時に前回終了位置へ復元する機能を追加
# - 保存位置が画面外の場合の補正機能を追加
# - 「常に手前に表示」チェックボックスを追加
# - 画面設定を setting.json に保存するよう変更
#
# v1.1.1 (2026-07-05)
# - チェックボックスの位置を調整
# ----------------------------------------------------------------------

import json
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox
from urllib.parse import quote
import webbrowser

import win32com.client


STATE_FILE = Path("state.json")
SETTING_FILE = Path("setting.json")
ICON_FILE = "NowPlayingPoster.ico"

WINDOW_WIDTH = 420
WINDOW_HEIGHT = 160


def resource_path(relative_path):
    """
    PyInstallerでexe化した場合でも、同梱ファイルを参照できるようにする。
    通常実行時は、このPythonファイルと同じフォルダを参照する。
    """
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path

    return Path(__file__).parent / relative_path


class NowPlayingGUI:

    def __init__(self, root):

        self.root = root

        self.state = self.load_state()
        self.setting = self.load_setting()

        root.title("NowPlayingPoster")

        self.set_initial_geometry()

        root.resizable(False, False)

        root.protocol("WM_DELETE_WINDOW", self.close_app)

        self.topmost_var = tk.BooleanVar(
            value=self.setting["topmost"]
        )

        root.attributes(
            "-topmost",
            self.topmost_var.get()
        )

        # タイトルバー左上のアイコンを設定
        icon_path = resource_path(ICON_FILE)

        if icon_path.exists():
            try:
                root.iconbitmap(str(icon_path))
            except Exception as e:
                print("アイコン設定エラー:", e)
        else:
            print("アイコンファイルが見つかりません:", icon_path)

        self.itunes = win32com.client.Dispatch(
            "iTunes.Application"
        )

        self.last_track = self.state["last_track"]
        self.last_album = self.state["last_album"]

        self.text = tk.Text(
            root,
            height=4,
            width=48,
            wrap="word",
            font=("Meiryo", 10)
        )

        self.text.pack(
            padx=8,
            pady=(8, 4),
            fill="both"
        )

        tk.Checkbutton(
            root,
            text="常に手前に表示",
            variable=self.topmost_var,
            command=self.toggle_topmost
        ).pack(
            anchor="w",
            padx=8,
            pady=(0, 2)
        )

        button_frame = tk.Frame(root)
        button_frame.pack(pady=(2, 6))

        tk.Button(
            button_frame,
            text="更新",
            width=10,
            command=self.refresh_display
        ).pack(side="left", padx=5)

        tk.Button(
            button_frame,
            text="Xへ投稿",
            width=10,
            command=self.post_to_x
        ).pack(side="left", padx=5)

        tk.Button(
            button_frame,
            text="終了",
            width=10,
            command=self.close_app
        ).pack(side="left", padx=5)

        self.refresh_display()

        print("NowPlaying Monitor 起動")

        self.monitor()

    # --------------------------
    # state管理
    # --------------------------

    def load_state(self):

        if STATE_FILE.exists():

            with open(
                STATE_FILE,
                "r",
                encoding="utf-8"
            ) as f:

                return json.load(f)

        return {
            "last_track": None,
            "last_album": None
        }

    def save_state(self):

        with open(
            STATE_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                self.state,
                f,
                ensure_ascii=False,
                indent=2
            )

    # --------------------------
    # setting管理
    # --------------------------

    def load_setting(self):

        if SETTING_FILE.exists():

            with open(
                SETTING_FILE,
                "r",
                encoding="utf-8"
            ) as f:

                setting = json.load(f)

            return {
                "window_x": setting.get("window_x"),
                "window_y": setting.get("window_y"),
                "topmost": setting.get("topmost", False)
            }

        return {
            "window_x": None,
            "window_y": None,
            "topmost": False
        }

    def save_setting(self):

        self.setting["window_x"] = self.root.winfo_x()
        self.setting["window_y"] = self.root.winfo_y()
        self.setting["topmost"] = self.topmost_var.get()

        with open(
            SETTING_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                self.setting,
                f,
                ensure_ascii=False,
                indent=2
            )

    # --------------------------
    # ウィンドウ位置
    # --------------------------

    def set_initial_geometry(self):

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x = self.setting.get("window_x")
        y = self.setting.get("window_y")

        if x is None or y is None:

            x = (screen_width - WINDOW_WIDTH) // 2
            y = (screen_height - WINDOW_HEIGHT) // 2

        else:

            x, y = self.safe_window_position(
                x,
                y,
                screen_width,
                screen_height
            )

        self.root.geometry(
            f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{x}+{y}"
        )

    def safe_window_position(
        self,
        x,
        y,
        screen_width,
        screen_height
    ):

        if x < 0:
            x = 0

        if y < 0:
            y = 0

        if x + WINDOW_WIDTH > screen_width:
            x = screen_width - WINDOW_WIDTH

        if y + WINDOW_HEIGHT > screen_height:
            y = screen_height - WINDOW_HEIGHT

        if x < 0:
            x = 0

        if y < 0:
            y = 0

        return x, y

    # --------------------------
    # 常に手前に表示
    # --------------------------

    def toggle_topmost(self):

        self.root.attributes(
            "-topmost",
            self.topmost_var.get()
        )

        self.setting["topmost"] = self.topmost_var.get()

    # --------------------------
    # 投稿文保存
    # --------------------------

    def save_tweet(self, track):

        tweet_text = (
            f"♪ {track.Name} - "
            f"{track.Album} by "
            f"{track.Artist} "
            f"#NowPlaying"
        )

        with open(
            "current_tweet.txt",
            "w",
            encoding="utf-8"
        ) as f:

            f.write(tweet_text)

        return tweet_text

    # --------------------------
    # アートワーク保存
    # --------------------------

    def save_artwork(self, track):

        artworks = track.Artwork

        if artworks.Count == 0:
            return

        artwork = artworks.Item(1)

        fmt = artwork.Format

        for filename in [
            "current_artwork.jpg",
            "current_artwork.png"
        ]:

            p = Path(filename)

            if p.exists():
                p.unlink()

        if fmt == 1:

            outfile = (
                Path.cwd()
                / "current_artwork.jpg"
            )

        else:

            outfile = (
                Path.cwd()
                / "current_artwork.png"
            )

        artwork.SaveArtworkToFile(
            str(outfile)
        )

        print(
            "アートワーク保存:",
            outfile.name
        )

    # --------------------------
    # 表示更新
    # --------------------------

    def refresh_display(self):

        tweet_file = Path(
            "current_tweet.txt"
        )

        if tweet_file.exists():

            text = tweet_file.read_text(
                encoding="utf-8"
            )

        else:

            text = (
                "まだ投稿候補がありません"
            )

        self.text.delete(
            "1.0",
            tk.END
        )

        self.text.insert(
            tk.END,
            text
        )

    # --------------------------
    # iTunes監視
    # --------------------------

    def monitor(self):

        try:

            track = self.itunes.CurrentTrack

            if track:

                current_track = track.Name
                current_album = track.Album

                if current_track != self.last_track:

                    tweet = self.save_tweet(
                        track
                    )

                    print()
                    print("=" * 60)
                    print(tweet)

                    if (
                        current_album
                        != self.last_album
                    ):

                        print(
                            "アルバム変更検知"
                        )

                        self.save_artwork(
                            track
                        )

                    else:

                        print(
                            "同一アルバム"
                        )

                    self.last_track = (
                        current_track
                    )

                    self.last_album = (
                        current_album
                    )

                    self.state[
                        "last_track"
                    ] = current_track

                    self.state[
                        "last_album"
                    ] = current_album

                    self.save_state()

                    self.refresh_display()

        except Exception as e:

            print(
                "監視エラー:",
                e
            )

        self.root.after(
            5000,
            self.monitor
        )

    # --------------------------
    # X投稿
    # --------------------------

    def post_to_x(self):

        tweet_text = self.text.get(
            "1.0",
            tk.END
        ).strip()

        if not tweet_text:

            messagebox.showwarning(
                "投稿文なし",
                "投稿する内容がありません。"
            )

            return

        encoded_text = quote(
            tweet_text
        )

        url = (
            "https://x.com/intent/post?text="
            + encoded_text
        )

        webbrowser.open(url)

    # --------------------------
    # 終了処理
    # --------------------------

    def close_app(self):

        result = messagebox.askyesno(
            "終了確認",
            "NowPlayingPosterを終了します。\nよろしいですか？"
        )

        if result:
            self.save_setting()
            self.root.destroy()


root = tk.Tk()

app = NowPlayingGUI(root)

root.mainloop()
