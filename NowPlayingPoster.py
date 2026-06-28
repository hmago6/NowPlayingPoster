# ----------------------------------------------------------------------
# NowPlaying Poster v1.0.0 (2026-06-28)
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
ICON_FILE = "NowPlayingPoster.ico"


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

        root.title("NowPlaying Poster")
        root.geometry("420x170")
        root.resizable(False, False)

        root.protocol("WM_DELETE_WINDOW", self.close_app)

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

        self.state = self.load_state()

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
            "NowPlaying Posterを終了します。\nよろしいですか？"
        )

        if result:
            self.root.destroy()


root = tk.Tk()

app = NowPlayingGUI(root)

root.mainloop()
