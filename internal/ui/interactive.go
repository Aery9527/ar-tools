package ui

import (
	"bufio"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"ar-tools/internal/dialog"
	"ar-tools/internal/xlsx2md"
)

// Run starts the interactive CLI flow.
func Run() error {
	scanner := bufio.NewScanner(os.Stdin)

	for {
		action, err := selectAction(scanner)
		if err != nil {
			return err
		}

		switch action {
		case 0:
			fmt.Println("Bye!")
			return nil
		case 1:
			if err := runXlsx2md(); err != nil {
				return err
			}
		default:
			fmt.Printf("無效的選項: %d\n", action)
		}
	}
}

func selectAction(scanner *bufio.Scanner) (int, error) {
	fmt.Println("\n請選擇功能:")
	fmt.Println("  1) Excel (.xlsx) → Markdown (.md)")
	fmt.Println("  0) 離開")
	fmt.Print("\n請輸入編號: ")

	scanner.Scan()
	input := strings.TrimSpace(scanner.Text())

	n, err := strconv.Atoi(input)
	if err != nil {
		return -1, nil
	}
	return n, nil
}

func runXlsx2md() error {
	files, err := dialog.OpenMultipleFiles(
		"選擇 Excel 檔案",
		"Excel files (*.xlsx)",
		"*.xlsx",
	)
	if err != nil {
		return fmt.Errorf("檔案選擇失敗: %w", err)
	}
	if len(files) == 0 {
		fmt.Println("已取消選擇")
		return nil
	}

	return convertFiles(files)
}

func convertFiles(files []string) error {
	opts := xlsx2md.ConvertOptions{}

	var succeeded, failed int
	for _, f := range files {
		outPath := strings.TrimSuffix(f, filepath.Ext(f)) + ".md"

		result, err := xlsx2md.Convert(f, opts)
		if err != nil {
			fmt.Printf("✗ %s: %v\n", filepath.Base(f), err)
			failed++
			continue
		}

		if err := os.WriteFile(outPath, []byte(result), 0644); err != nil {
			fmt.Printf("✗ %s: failed to write output: %v\n", filepath.Base(f), err)
			failed++
			continue
		}

		fmt.Printf("✓ %s → %s\n", filepath.Base(f), filepath.Base(outPath))
		succeeded++
	}

	fmt.Printf("\n完成: %d 成功", succeeded)
	if failed > 0 {
		fmt.Printf(", %d 失敗", failed)
	}
	fmt.Println()

	return nil
}
