package xlsx2md

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// ConvertOptions holds configuration for the conversion.
type ConvertOptions struct {
	// SheetNames specifies which sheets to convert. Empty means all sheets.
	SheetNames []string
}

// Convert reads an Excel file and returns its content as Markdown tables.
// Each sheet is rendered as a separate section with a heading.
func Convert(filePath string, opts ConvertOptions) (string, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return "", fmt.Errorf("failed to open excel file: %w", err)
	}
	defer f.Close()

	sheets := opts.SheetNames
	if len(sheets) == 0 {
		sheets = f.GetSheetList()
	}

	var sb strings.Builder
	for i, sheet := range sheets {
		rows, err := f.GetRows(sheet)
		if err != nil {
			return "", fmt.Errorf("failed to read sheet %q: %w", sheet, err)
		}
		if len(rows) == 0 {
			continue
		}

		if i > 0 {
			sb.WriteString("\n")
		}
		sb.WriteString(fmt.Sprintf("## %s\n\n", sheet))
		sb.WriteString(sheetToMarkdown(rows))
	}

	return sb.String(), nil
}

// ConvertSheet converts a single sheet's rows into a Markdown table string.
func ConvertSheet(rows [][]string) string {
	return sheetToMarkdown(rows)
}

func sheetToMarkdown(rows [][]string) string {
	if len(rows) == 0 {
		return ""
	}

	// Determine max column count across all rows
	maxCols := 0
	for _, row := range rows {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}
	if maxCols == 0 {
		return ""
	}

	var sb strings.Builder

	// Header row
	header := padRow(rows[0], maxCols)
	sb.WriteString("| " + strings.Join(escapeCells(header), " | ") + " |\n")

	// Separator row
	seps := make([]string, maxCols)
	for i := range seps {
		seps[i] = "---"
	}
	sb.WriteString("| " + strings.Join(seps, " | ") + " |\n")

	// Data rows
	for _, row := range rows[1:] {
		padded := padRow(row, maxCols)
		sb.WriteString("| " + strings.Join(escapeCells(padded), " | ") + " |\n")
	}

	return sb.String()
}

// padRow ensures the row has exactly n columns, padding with empty strings.
func padRow(row []string, n int) []string {
	if len(row) >= n {
		return row[:n]
	}
	padded := make([]string, n)
	copy(padded, row)
	return padded
}

// escapeCells escapes pipe characters in cell values to avoid breaking Markdown tables.
func escapeCells(cells []string) []string {
	out := make([]string, len(cells))
	for i, c := range cells {
		out[i] = strings.ReplaceAll(c, "|", "\\|")
	}
	return out
}
