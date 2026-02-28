---
applyTo: "**/*.go"
---

# project-context

## 概述

ar-tools 是一個 Windows 互動式 CLI 工具集，提供各種開發者實用功能。目前支援 Excel (.xlsx) 轉 Markdown (.md) 表格。

## 專案結構

```
ar-tools/
├── cmd/
│   └── ar-tools/
│       └── main.go              # 唯一入口點，呼叫 ui.Run()
├── internal/
│   ├── dialog/
│   │   └── dialog_windows.go   # Windows 原生檔案對話框 (comdlg32.dll GetOpenFileNameW)
│   ├── ui/
│   │   └── interactive.go      # 互動式選單與流程控制
│   └── xlsx2md/
│       ├── converter.go        # Excel → Markdown 核心轉換邏輯
│       └── converter_test.go   # 轉換邏輯測試
├── testdata/
│   └── sample.xlsx             # 測試用 Excel 範例檔
├── go.mod
└── go.sum
```

## 架構原則

- **入口點**: `cmd/<binary>/main.go`，遵循 Go 社群多 binary 慣例
- **業務邏輯**: 一律放在 `internal/` 底下，防止外部 import
- **平台相關**: 使用 Go build tag（如 `//go:build windows`）隔離平台特定程式碼
- **無框架依賴**: CLI 選單使用 `bufio.Scanner` + `fmt` 標準庫；檔案對話框直接透過 `syscall` 呼叫 Win32 API，不依賴第三方 TUI/GUI 框架

## 主要依賴

| 套件 | 用途 |
| --- | --- |
| `github.com/xuri/excelize/v2` | 讀取 .xlsx 檔案 |
| `github.com/stretchr/testify` | 測試斷言 (assert) |

## 使用方式

```bash
go run ./cmd/ar-tools
```

執行後進入互動式選單：
1. 選擇功能編號（如 `1` = Excel → Markdown）
2. 彈出 Windows 原生檔案選擇視窗，支援 Ctrl+Click 多選 .xlsx
3. 轉換完成後回到主選單，輸入 `0` 離開

## 新增功能指引

新增一個轉換功能時：

1. **核心邏輯**: 在 `internal/<feature>/` 建立新 package，包含轉換函式與測試
2. **選單項目**: 在 `internal/ui/interactive.go` 的 `selectAction` 新增選項編號，在 `Run()` 的 switch 新增 case
3. **檔案選擇**: 使用 `dialog.OpenMultipleFiles()` 取得使用者選擇的檔案路徑

## Windows 檔案對話框

`internal/dialog/dialog_windows.go` 直接透過 `syscall` 呼叫 `comdlg32.dll!GetOpenFileNameW`：

- 不使用 COM（避免 Go goroutine 線程切換問題）
- 不 spawn 子程序（避免 `CreateProcess` 被安全策略擋住）
- 使用 `runtime.LockOSThread()` 確保 syscall 在同一個 OS 線程執行
