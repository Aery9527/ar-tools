---
applyTo: "**/*.go"
---

# project-context

## 鐵律

此專案執行環境有嚴格的 Windows 安全策略：

- **禁止 `os/exec`、`exec.Command`** 等任何 spawn 子程序的方式
- **禁止使用 COM / OLE** 相關操作（`CoInitialize`、`IFileOpenDialog` 等）
- **禁止任何第三方 GUI / 對話框 / TUI 庫**（如 `zenity`、`go-common-file-dialog`、`bubbletea`、`huh`、`fyne` 等）
- Windows 原生對話框僅允許透過 `syscall` 直接呼叫 `comdlg32.dll`（如 `GetOpenFileNameW`），此方式不經 COM、不 spawn 子程序

## 概述

ar-tools 是一個互動式 CLI 工具集，提供各種開發者實用功能。目前支援 Excel (.xlsx) 轉 Markdown (.md) 表格。

## 專案結構

```
ar-tools/
├── main.go                      # 唯一入口點，呼叫 ui.Run()
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

- **入口點**: 單一 binary 專案，`main.go` 直接放專案根目錄，使用 `go run .` 執行
- **業務邏輯**: 一律放在 `internal/` 底下，防止外部 import
- **Windows 對話框**: 透過 `syscall` + `comdlg32.dll` 直接呼叫，不使用第三方庫

## 主要依賴

| 套件 | 用途 |
| --- | --- |
| `github.com/xuri/excelize/v2` | 讀取 .xlsx 檔案 |
| `github.com/stretchr/testify` | 測試斷言 (assert) |

## 使用方式

```bash
go run .
```

執行後進入互動式選單：
1. 選擇功能編號（如 `1` = Excel → Markdown）
2. 彈出 Windows 原生檔案選擇視窗，Ctrl+Click 多選 .xlsx 檔案
3. 轉換完成後回到主選單，輸入 `0` 離開

## 新增功能指引

新增一個轉換功能時：

1. **核心邏輯**: 在 `internal/<feature>/` 建立新 package，包含轉換函式與測試
2. **選單項目**: 在 `internal/ui/interactive.go` 的 `selectAction` 新增選項編號，在 `Run()` 的 switch 新增 case
3. **檔案選擇**: 使用 `dialog.OpenMultipleFiles()` 開啟原生多選視窗
