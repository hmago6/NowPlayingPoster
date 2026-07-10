# ----------------------------------------------------------------------
# NowPlayingPoster v1.2.0 (2026-07-10)
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
#
# v1.2.0 (2026-07-10)
# - アルバムアートワークのウィンドウ内表示に対応
# - 左側に96×96のアートワーク、右側に投稿文を配置
# - JPEG／PNG形式のアートワーク表示に対応
# - アートワークがない場合に「No Artwork」を表示
# - アートワークがない曲へ移動した際、以前の画像を削除
# - Pillowを使用した画像表示処理を追加
# - アプリケーション情報を定数として一元管理
# - タイトルバーにバージョン番号を表示
# ----------------------------------------------------------------------

import ctypes
import json
import sys
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import messagebox
from urllib.parse import quote

import win32com.client
from PIL import Image, ImageDraw, ImageOps, ImageTk


# --------------------------
# アプリケーション情報
# --------------------------

APP_NAME = "NowPlayingPoster"
APP_VERSION = "1.2.0"
APP_TITLE = f"{APP_NAME} v{APP_VERSION}"
APP_AUTHOR = "HMago6"
APP_COPYRIGHT = "Copyright (c) 2026 HMago6"
APP_GITHUB = "https://github.com/hmago6/NowPlayingPoster"


# --------------------------
# ファイル・画面設定
# --------------------------

ICON_FILE = "NowPlayingPoster.ico"

WINDOW_WIDTH = 560
WINDOW_HEIGHT = 190

ARTWORK_WIDTH = 96
ARTWORK_HEIGHT = 96

MONITOR_INTERVAL_MS = 5000


def get_app_directory():
    """
    通常実行時はPythonファイルの保存先、
    PyInstallerによるexe実行時はexeの保存先を返す。
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent


APP_DIRECTORY = get_app_directory()

STATE_FILE = APP_DIRECTORY / "state.json"
SETTING_FILE = APP_DIRECTORY / "setting.json"
TWEET_FILE = APP_DIRECTORY / "current_tweet.txt"

ARTWORK_JPG_FILE = APP_DIRECTORY / "current_artwork.jpg"
ARTWORK_PNG_FILE = APP_DIRECTORY / "current_artwork.png"


def resource_path(relative_path):
    """
    PyInstallerでexe化した場合でも、
    同梱したリソースファイルを参照できるようにする。
    """
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path

    return Path(__file__).resolve().parent / relative_path


class NowPlayingGUI:

    def __init__(self, root):

        self.root = root

        self.state = self.load_state()
        self.setting = self.load_setting()

        self.root.title(APP_TITLE)
        self.set_initial_geometry()
        self.root.resizable(False, False)

        # 終了ボタンと右上の×を同じ終了処理へ接続
        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.close_app
        )

        self.topmost_var = tk.BooleanVar(
            value=self.setting["topmost"]
        )

        self.root.attributes(
            "-topmost",
            self.topmost_var.get()
        )

        # タイトルバー左上のアイコン
        icon_path = resource_path(ICON_FILE)

        if icon_path.exists():

            try:
                self.root.iconbitmap(
                    str(icon_path)
                )

            except Exception as error:
                print(
                    "アイコン設定エラー:",
                    error
                )

        else:
            print(
                "アイコンファイルが見つかりません:",
                icon_path
            )

        # iTunes COM接続
        self.itunes = win32com.client.Dispatch(
            "iTunes.Application"
        )

        # 起動後の初回監視では必ず現在曲を取得する
        self.last_track_id = None
        self.last_album = self.state.get(
            "last_album"
        )

        # Tkinterの画像を保持する変数
        self.artwork_photo = (
            self.create_placeholder_image()
        )

        # --------------------------
        # アートワーク＋投稿文
        # --------------------------

        content_frame = tk.Frame(
            self.root
        )

        content_frame.pack(
            padx=8,
            pady=(8, 4),
            fill="x"
        )

        # 左側：96×96アートワーク表示欄
        artwork_frame = tk.Frame(
            content_frame,
            width=ARTWORK_WIDTH,
            height=ARTWORK_HEIGHT
        )

        artwork_frame.pack(
            side="left",
            padx=(0, 8)
        )

        # 子ウィジェットのサイズでフレームが変形しないようにする
        artwork_frame.pack_propagate(
            False
        )

        self.artwork_label = tk.Label(
            artwork_frame,
            image=self.artwork_photo,
            relief="sunken",
            borderwidth=1,
            anchor="center"
        )

        self.artwork_label.pack(
            fill="both",
            expand=True
        )

        # 右側：投稿文表示・編集欄
        self.text = tk.Text(
            content_frame,
            height=6,
            width=51,
            wrap="word",
            font=("Meiryo", 10)
        )

        self.text.pack(
            side="left",
            fill="both",
            expand=True
        )

        # --------------------------
        # 常に手前に表示
        # --------------------------

        tk.Checkbutton(
            self.root,
            text="常に手前に表示",
            variable=self.topmost_var,
            command=self.toggle_topmost
        ).pack(
            anchor="w",
            padx=8,
            pady=(0, 2)
        )

        # --------------------------
        # 操作ボタン
        # --------------------------

        button_frame = tk.Frame(
            self.root
        )

        button_frame.pack(
            pady=(0, 6)
        )

        tk.Button(
            button_frame,
            text="更新",
            width=10,
            command=self.manual_refresh
        ).pack(
            side="left",
            padx=5
        )

        tk.Button(
            button_frame,
            text="Xへ投稿",
            width=10,
            command=self.post_to_x
        ).pack(
            side="left",
            padx=5
        )

        tk.Button(
            button_frame,
            text="終了",
            width=10,
            command=self.close_app
        ).pack(
            side="left",
            padx=5
        )

        # 保存済みの表示内容を読み込む
        self.refresh_display()
        self.refresh_artwork_display()

        print(
            f"{APP_TITLE} 起動"
        )

        # iTunes監視開始
        self.monitor()

    # --------------------------
    # JSON共通処理
    # --------------------------

    def read_json_file(
        self,
        file_path,
        default_value
    ):

        if not file_path.exists():
            return default_value.copy()

        try:
            with file_path.open(
                "r",
                encoding="utf-8"
            ) as file:

                data = json.load(
                    file
                )

            if not isinstance(
                data,
                dict
            ):
                raise ValueError(
                    "JSONの内容が辞書形式ではありません。"
                )

            return data

        except (
            OSError,
            json.JSONDecodeError,
            ValueError
        ) as error:

            print(
                f"{file_path.name} 読み込みエラー:",
                error
            )

            return default_value.copy()

    def write_json_file(
        self,
        file_path,
        data
    ):

        try:
            with file_path.open(
                "w",
                encoding="utf-8"
            ) as file:

                json.dump(
                    data,
                    file,
                    ensure_ascii=False,
                    indent=2
                )

        except OSError as error:

            print(
                f"{file_path.name} 保存エラー:",
                error
            )

    # --------------------------
    # state管理
    # --------------------------

    def load_state(self):

        default_state = {
            "last_track": None,
            "last_album": None,
            "last_artist": None
        }

        state = self.read_json_file(
            STATE_FILE,
            default_state
        )

        return {
            "last_track": state.get(
                "last_track"
            ),
            "last_album": state.get(
                "last_album"
            ),
            "last_artist": state.get(
                "last_artist"
            )
        }

    def save_state(self):

        self.write_json_file(
            STATE_FILE,
            self.state
        )

    # --------------------------
    # setting管理
    # --------------------------

    def load_setting(self):

        default_setting = {
            "window_x": None,
            "window_y": None,
            "topmost": False
        }

        setting = self.read_json_file(
            SETTING_FILE,
            default_setting
        )

        return {
            "window_x": setting.get(
                "window_x"
            ),
            "window_y": setting.get(
                "window_y"
            ),
            "topmost": bool(
                setting.get(
                    "topmost",
                    False
                )
            )
        }

    def save_setting(self):

        self.setting["window_x"] = (
            self.root.winfo_x()
        )

        self.setting["window_y"] = (
            self.root.winfo_y()
        )

        self.setting["topmost"] = (
            self.topmost_var.get()
        )

        self.write_json_file(
            SETTING_FILE,
            self.setting
        )

    # --------------------------
    # ウィンドウ位置
    # --------------------------

    def get_virtual_screen_bounds(self):

        """
        Windowsの仮想デスクトップ全体の座標を取得する。
        外部モニターが左側にある場合の負の座標にも対応する。
        """
        user32 = ctypes.windll.user32

        virtual_x = user32.GetSystemMetrics(
            76
        )

        virtual_y = user32.GetSystemMetrics(
            77
        )

        virtual_width = user32.GetSystemMetrics(
            78
        )

        virtual_height = user32.GetSystemMetrics(
            79
        )

        return (
            virtual_x,
            virtual_y,
            virtual_width,
            virtual_height
        )

    def set_initial_geometry(self):

        (
            screen_x,
            screen_y,
            screen_width,
            screen_height
        ) = self.get_virtual_screen_bounds()

        x = self.setting.get(
            "window_x"
        )

        y = self.setting.get(
            "window_y"
        )

        if x is None or y is None:

            # 初回起動時は仮想画面全体の中央
            x = (
                screen_x
                + (
                    screen_width
                    - WINDOW_WIDTH
                ) // 2
            )

            y = (
                screen_y
                + (
                    screen_height
                    - WINDOW_HEIGHT
                ) // 2
            )

        else:

            x, y = (
                self.safe_window_position(
                    x,
                    y,
                    screen_x,
                    screen_y,
                    screen_width,
                    screen_height
                )
            )

        self.root.geometry(
            f"{WINDOW_WIDTH}x"
            f"{WINDOW_HEIGHT}"
            f"{x:+d}{y:+d}"
        )

    def safe_window_position(
        self,
        x,
        y,
        screen_x,
        screen_y,
        screen_width,
        screen_height
    ):

        maximum_x = (
            screen_x
            + screen_width
            - WINDOW_WIDTH
        )

        maximum_y = (
            screen_y
            + screen_height
            - WINDOW_HEIGHT
        )

        # 画面よりウィンドウの方が大きい場合にも対応
        maximum_x = max(
            screen_x,
            maximum_x
        )

        maximum_y = max(
            screen_y,
            maximum_y
        )

        x = min(
            max(x, screen_x),
            maximum_x
        )

        y = min(
            max(y, screen_y),
            maximum_y
        )

        return x, y

    # --------------------------
    # 常に手前に表示
    # --------------------------

    def toggle_topmost(self):

        topmost = (
            self.topmost_var.get()
        )

        self.root.attributes(
            "-topmost",
            topmost
        )

        self.setting["topmost"] = (
            topmost
        )

        # チェック変更時にも設定を保存する
        self.save_setting()

    # --------------------------
    # 投稿文関連
    # --------------------------

    def create_tweet_text(
        self,
        track
    ):

        return (
            f"♪ {track.Name} - "
            f"{track.Album} by "
            f"{track.Artist} "
            f"#NowPlaying"
        )

    def save_tweet(
        self,
        track
    ):

        tweet_text = (
            self.create_tweet_text(
                track
            )
        )

        try:
            TWEET_FILE.write_text(
                tweet_text,
                encoding="utf-8"
            )

        except OSError as error:

            print(
                "投稿文保存エラー:",
                error
            )

        return tweet_text

    def refresh_display(self):

        if TWEET_FILE.exists():

            try:
                text = (
                    TWEET_FILE.read_text(
                        encoding="utf-8"
                    )
                )

            except OSError as error:

                print(
                    "投稿文読み込みエラー:",
                    error
                )

                text = (
                    "投稿文を読み込めませんでした"
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
    # アートワーク関連
    # --------------------------

    def create_placeholder_image(self):

        image = Image.new(
            "RGB",
            (
                ARTWORK_WIDTH,
                ARTWORK_HEIGHT
            ),
            (
                240,
                240,
                240
            )
        )

        drawer = ImageDraw.Draw(
            image
        )

        placeholder_text = (
            "No Artwork"
        )

        text_box = drawer.textbbox(
            (
                0,
                0
            ),
            placeholder_text
        )

        text_width = (
            text_box[2]
            - text_box[0]
        )

        text_height = (
            text_box[3]
            - text_box[1]
        )

        text_x = (
            ARTWORK_WIDTH
            - text_width
        ) // 2

        text_y = (
            ARTWORK_HEIGHT
            - text_height
        ) // 2

        drawer.text(
            (
                text_x,
                text_y
            ),
            placeholder_text,
            fill=(
                90,
                90,
                90
            )
        )

        return ImageTk.PhotoImage(
            image
        )

    def delete_old_artwork(self):

        for artwork_file in (
            ARTWORK_JPG_FILE,
            ARTWORK_PNG_FILE
        ):

            try:
                if artwork_file.exists():
                    artwork_file.unlink()

            except OSError as error:

                print(
                    "古いアートワーク削除エラー:",
                    error
                )

    def get_current_artwork_path(self):

        if ARTWORK_JPG_FILE.exists():
            return ARTWORK_JPG_FILE

        if ARTWORK_PNG_FILE.exists():
            return ARTWORK_PNG_FILE

        return None

    def save_artwork(
        self,
        track
    ):

        try:
            artworks = (
                track.Artwork
            )

            # 前のアルバム画像が残らないように先に削除
            self.delete_old_artwork()

            if artworks.Count == 0:

                print(
                    "アートワークなし"
                )

                self.show_no_artwork()

                return

            artwork = artworks.Item(
                1
            )

            artwork_format = (
                artwork.Format
            )

            if artwork_format == 1:

                output_file = (
                    ARTWORK_JPG_FILE
                )

            else:

                output_file = (
                    ARTWORK_PNG_FILE
                )

            # iTunes COMでは絶対パスを渡す必要がある
            artwork.SaveArtworkToFile(
                str(
                    output_file.resolve()
                )
            )

            print(
                "アートワーク保存:",
                output_file.name
            )

            self.display_artwork(
                output_file
            )

        except Exception as error:

            print(
                "アートワーク保存エラー:",
                error
            )

            self.delete_old_artwork()
            self.show_no_artwork()

    def display_artwork(
        self,
        artwork_path
    ):

        try:
            with Image.open(
                artwork_path
            ) as source_image:

                image = (
                    source_image
                    .convert("RGBA")
                )

                image = ImageOps.contain(
                    image,
                    (
                        ARTWORK_WIDTH,
                        ARTWORK_HEIGHT
                    ),
                    method=Image.Resampling.LANCZOS
                )

                # 正方形ではない画像でも96×96の中央に配置
                canvas = Image.new(
                    "RGBA",
                    (
                        ARTWORK_WIDTH,
                        ARTWORK_HEIGHT
                    ),
                    (
                        240,
                        240,
                        240,
                        255
                    )
                )

                paste_x = (
                    ARTWORK_WIDTH
                    - image.width
                ) // 2

                paste_y = (
                    ARTWORK_HEIGHT
                    - image.height
                ) // 2

                canvas.alpha_composite(
                    image,
                    (
                        paste_x,
                        paste_y
                    )
                )

                self.artwork_photo = (
                    ImageTk.PhotoImage(
                        canvas
                    )
                )

            self.artwork_label.configure(
                image=self.artwork_photo
            )

        except (
            OSError,
            ValueError
        ) as error:

            print(
                "アートワーク表示エラー:",
                error
            )

            self.show_no_artwork()

    def show_no_artwork(self):

        self.artwork_photo = (
            self.create_placeholder_image()
        )

        self.artwork_label.configure(
            image=self.artwork_photo
        )

    def refresh_artwork_display(self):

        artwork_path = (
            self.get_current_artwork_path()
        )

        if artwork_path is None:

            self.show_no_artwork()

        else:

            self.display_artwork(
                artwork_path
            )

    # --------------------------
    # iTunes情報更新
    # --------------------------

    def update_from_itunes(
        self,
        force=False
    ):

        track = (
            self.itunes.CurrentTrack
        )

        if not track:
            return

        current_track = (
            track.Name
        )

        current_album = (
            track.Album
        )

        current_artist = (
            track.Artist
        )

        current_track_id = (
            current_track,
            current_album,
            current_artist
        )

        if (
            not force
            and current_track_id
            == self.last_track_id
        ):
            return

        tweet_text = (
            self.save_tweet(
                track
            )
        )

        print()
        print(
            "=" * 60
        )
        print(
            tweet_text
        )

        album_changed = (
            current_album
            != self.last_album
        )

        artwork_missing = (
            self.get_current_artwork_path()
            is None
        )

        if (
            force
            or album_changed
            or artwork_missing
        ):

            if album_changed:
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

            self.refresh_artwork_display()

        self.last_track_id = (
            current_track_id
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

        self.state[
            "last_artist"
        ] = current_artist

        self.save_state()
        self.refresh_display()

    def manual_refresh(self):

        try:
            self.update_from_itunes(
                force=True
            )

        except Exception as error:

            print(
                "手動更新エラー:",
                error
            )

            messagebox.showerror(
                "更新エラー",
                "iTunesから再生情報を取得できませんでした。"
            )

    # --------------------------
    # iTunes監視
    # --------------------------

    def monitor(self):

        try:
            self.update_from_itunes()

        except Exception as error:

            print(
                "監視エラー:",
                error
            )

        self.root.after(
            MONITOR_INTERVAL_MS,
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
            "https://x.com/"
            "intent/post?text="
            + encoded_text
        )

        webbrowser.open(
            url
        )

    # --------------------------
    # 終了処理
    # --------------------------

    def close_app(self):

        result = messagebox.askyesno(
            "終了確認",
            f"{APP_NAME}を終了します。\n"
            "よろしいですか？"
        )

        if result:

            self.save_setting()

            self.root.destroy()


def main():

    root = tk.Tk()

    NowPlayingGUI(
        root
    )

    root.mainloop()


if __name__ == "__main__":
    main()
