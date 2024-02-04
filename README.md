# name
overtime_auto

## Overview
このツールはexcelのシート「List」を元に、
Pythonのモジュール「selenium」でターゲットのサイトに自動で残業を転記をするツールです。

## Requirement
- Windows os
- Excel
- Python
- openpyxl
- bs4
- selenium(各種モジュール)

## Usage
excelを開き必要情報を入力し、Pythonツールを動作させる。

## Description
1.excelのシート「List」の「対象月」を対象の月に変更し、
「month_filter」をクリック。

2.「始業」、「就業」、「理由」を入力。

3.「file出力」をクリックし、ここまででexcelでの操作は終了。
※fileはローカルのドキュメントに出力されます。

4.「overtime_post.py」を開く。

5.開くとターミナルが立ち上がるので、「URL」、「ID」、「PW」を入力する。

6.ターミナル上で「処理が完了しました。」と表示されたら、Wチェックを行いサイト上の「登録」を押し作業は終了。
※動作終了後もChromeは立ち上がったままなので手動でウインドウを閉じてください。

## Reference

## Author

## Licence

