# NowPlayingPoster

Windows版 iTunes で現在再生中の曲情報を取得し、X (旧Twitter) への投稿を支援するデスクトップアプリケーションです。

## Features

* Windows版 iTunes の現在再生中の曲情報を取得
* 曲名・アルバム名・アーティスト名を自動取得
* X (旧Twitter) 投稿文を自動生成
* X の投稿画面をブラウザで自動オープン
* Apple公式版 iTunes の家庭内共有機能に対応
* API不要
* ローカル環境のみで動作
* アルバムアートワークも手動でアップロード可能

生成される投稿文の例：

♪ 曲名 - アルバム名 by アーティスト名 #NowPlaying
```

---

## Requirements

### Supported OS

* Windows 11
* Windows 10（未検証）

### Required Software

* Apple公式版 iTunes
* Python 3.10 以降

> Microsoft Store版 iTunes は動作未確認です。

### Required Python Package

```bash
pip install pywin32
```

---

## Usage

iTunes を起動し、音楽を再生した状態で以下を実行します。

```bash
python NowPlayingPoster.py
```

または、配布されている EXE ファイルを実行してください。

### GUI Buttons

| Button | Description |
| ------ | ----------- |
| 更新     | 現在の投稿文を更新   |
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
