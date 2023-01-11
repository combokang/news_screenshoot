# News Screenshot Helper

新聞全頁截圖工具

## Table of Contents

- [Background](#background)
- [Usage](#usage)
- [Handling](#handling)
- [Known Issues](#known-issues)
- [License](#license)

## Background

公關公司經常有需要出為特定客戶出公關監測週報的需求。此程式配合其他監控軟體產出的新聞列表，可協助公關公司降低手動截圖的困難。

## Usage

- 前置工作
	1. 需將新聞整理成excel檔案(副檔名xlsx)，A欄需有編號，H欄需有新聞超連結
	2. 只支援第一個工作表
	3. xlsx檔需放置在與程式同個資料夾
- 開始使用
	1. 開啟excelhandler.py
	2. 輸入xlsx檔名(不須輸入.xlsx)
	3. 開始截圖，截圖將放置在與xlsx相同名稱之資料夾
	4. 截圖結束後，若有需要重新截圖者，將該圖檔移出資料夾，輸入1後將重新開始截圖
	5.若沒問題則結束程式

## Handling

- 有反爬蟲程式碼偵測連線速度，連線過快將被阻擋：台灣時報 => 全部等 2 秒
- 連線時會跳出訂閱通知：自由時報新聞、自由時報地產天下 => 透過 MouseEvent 按下取消
- 圖片為 lazy load 模式：蘋果地產、蘋果日報、新浪台灣、新浪新聞、UDN、自由時報 => 透過 ScrollBy 新聞區塊高度觸發 lazy load
- 擋版廣告：卡優新聞往 => 觸發 javascript 關閉

## Known Issues

- 在開啟很多個網頁時，可能會未完整載入，應是cache導致
- 中時的截圖會留下廣告圖片框
- New Talk的截圖會留下推薦文章框

## License

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)