# DataExtractor

Google スプレッドシートからデータを抽出し、特定の形式で出力するためのGoogle Apps Scriptプロジェクトです。

## 概要

このツールは、スプレッドシート内の特定のパターンに一致するデータを検索し、別のシートに整形して出力します。主に「x」で始まるセル（男子データ）を抽出する機能が実装されています。

## 機能

- スプレッドシートに「データ取得」メニューを追加
- 「ソースコード作成」機能によるデータの抽出と整形
- 特定のパターン（「x」で始まるセル）に一致するデータの検索
- 抽出したデータを「スクリプト」シートに出力
- 「データ抽出」シートのクリア機能

## 使用方法

1. スプレッドシートを開く
2. 上部メニューから「データ取得」を選択
3. 「ソースコード作成」をクリックしてデータ抽出を実行

## シート構成

- **アクティブシート**：データの抽出元となるシート
- **スクリプト**：抽出したデータの出力先
- **データ抽出**：データのタイトル情報を含むシート

## 注意事項

- 現在は「x」で始まるセル（男子データ）の抽出のみが有効になっています
- 「y」で始まるセル（女子データ）の抽出機能はコメントアウトされています

## ライセンス

MIT License

Copyright (c) 2025 Noboru Ando @ Aoyama Gakuin University

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
