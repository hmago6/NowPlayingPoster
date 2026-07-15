# NowPlayingPoster

![Release](https://img.shields.io/github/v/release/hmago6/NowPlayingPoster)
![License](https://img.shields.io/github/license/hmago6/NowPlayingPoster)

<img width="559" height="224" alt="image" src="https://github.com/user-attachments/assets/56def44d-0fdd-4653-a825-11ca46c05365" />

NowPlayingPoster は、Windows版 iTunes の現在再生中の曲情報を取得し、
アルバムアートワークと投稿文を表示しながら、
X (旧Twitter) への投稿を支援するデスクトップアプリケーションです。

投稿文は自動生成されますが、投稿前に自由に編集できます。


## Features

* Windows版 iTunes の現在再生中の曲情報を取得
* 曲名・アルバム名・アーティスト名を自動取得
* アルバムアートワークを表示(v1.2.0)
* X (旧Twitter) 投稿文を自動生成
* 投稿前に自由に編集可能
* X の投稿画面を既定ブラウザで自動オープン
* Apple公式版 iTunes の家庭内共有機能に対応
* API不要
* ローカル環境のみで動作
* アルバムアートワークも手動でアップロード可能
* ウィンドウ位置保存 (v1.1.0)
* 常に手前に表示(v1.1.0)

生成される投稿文の例：

```text
♪ 曲名 - アルバム名 by アーティスト名 #NowPlaying
```

---

## Requirements

### Supported OS

* Windows 11

> Windows 10 での動作は未検証です。

### Required Software

#### Release版 (.exe) を使用する場合

* Apple公式版 iTunes

> Microsoft Store版 iTunes は動作未確認です。

#### ソースコードから実行する場合
* Python 3.11 以降
* pywin32
* Pillow

### Required Python Packages

```bash
python -m pip install -r requirements.txt
```

## Usage

1. Windows版 iTunes を起動します。
2. 曲を再生します。
3. 投稿文が自動生成されます。
4. 必要に応じて投稿文を自由に編集します。
5. 「Xへ投稿」を押します。
   
```bash
python NowPlayingPoster.py
```

### Release

最新版は Releases からダウンロードできます。

アルバムアートワークは、実行モジュールと同一フォルダに`current_artwork.jpg` または `current_artwork.png` として保管されます。

必要に応じてブラウザの投稿画面で選択してから投稿してください。

### GUI Buttons

| Button | Description |
| ------ | ----------- |
| 更新     | 現在の曲情報とアートワークを再取得して更新 |
| Xへ投稿   | X の投稿画面を開く  |
| 終了     | アプリケーションを終了 |


## Known Limitations

* Apple Music アプリには対応していません。
* iTunes が起動していない場合は動作しません。
* X への投稿はユーザー自身が最終確認して行います。
* 本ソフトは X API を使用しません。

---

## License

This project is licensed under the MIT License.

See the LICENSE file for details.

---

## Author

HMago6

---

## Acknowledgements

This application uses the following open-source software:

* Python
* pywin32
* Tkinter
* Pillow
