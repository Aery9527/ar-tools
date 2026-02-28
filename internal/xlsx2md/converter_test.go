package xlsx2md

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

func TestConvertSheet(t *testing.T) {
	tests := []struct {
		name     string
		rows     [][]string
		expected string
	}{
		{
			name:     "empty rows",
			rows:     [][]string{},
			expected: "",
		},
		{
			name: "header only",
			rows: [][]string{{"Name", "Age"}},
			expected: "| Name | Age |\n" +
				"| --- | --- |\n",
		},
		{
			name: "normal table",
			rows: [][]string{
				{"Name", "Age", "City"},
				{"Alice", "30", "Taipei"},
				{"Bob", "25", "Tokyo"},
			},
			expected: "| Name | Age | City |\n" +
				"| --- | --- | --- |\n" +
				"| Alice | 30 | Taipei |\n" +
				"| Bob | 25 | Tokyo |\n",
		},
		{
			name: "ragged rows padded",
			rows: [][]string{
				{"A", "B", "C"},
				{"1"},
				{"x", "y"},
			},
			expected: "| A | B | C |\n" +
				"| --- | --- | --- |\n" +
				"| 1 |  |  |\n" +
				"| x | y |  |\n",
		},
		{
			name: "pipe character escaped",
			rows: [][]string{
				{"Col"},
				{"a|b"},
			},
			expected: "| Col |\n" +
				"| --- |\n" +
				"| a\\|b |\n",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := ConvertSheet(tt.rows)
			assert.Equal(t, tt.expected, result)
		})
	}
}

func TestConvert(t *testing.T) {
	// Create a temporary Excel file for testing
	f := excelize.NewFile()
	defer f.Close()

	// Sheet1 (default)
	f.SetCellValue("Sheet1", "A1", "Name")
	f.SetCellValue("Sheet1", "B1", "Score")
	f.SetCellValue("Sheet1", "A2", "Alice")
	f.SetCellValue("Sheet1", "B2", "95")

	// Sheet2
	f.NewSheet("Sheet2")
	f.SetCellValue("Sheet2", "A1", "Item")
	f.SetCellValue("Sheet2", "B1", "Price")
	f.SetCellValue("Sheet2", "A2", "Apple")
	f.SetCellValue("Sheet2", "B2", "10")

	tmpDir := t.TempDir()
	tmpFile := filepath.Join(tmpDir, "test.xlsx")
	err := f.SaveAs(tmpFile)
	assert.NoError(t, err)

	t.Run("convert all sheets", func(t *testing.T) {
		result, err := Convert(tmpFile, ConvertOptions{})
		assert.NoError(t, err)
		assert.Contains(t, result, "## Sheet1")
		assert.Contains(t, result, "## Sheet2")
		assert.Contains(t, result, "| Name | Score |")
		assert.Contains(t, result, "| Alice | 95 |")
		assert.Contains(t, result, "| Item | Price |")
		assert.Contains(t, result, "| Apple | 10 |")
	})

	t.Run("convert specific sheet", func(t *testing.T) {
		result, err := Convert(tmpFile, ConvertOptions{SheetNames: []string{"Sheet2"}})
		assert.NoError(t, err)
		assert.NotContains(t, result, "## Sheet1")
		assert.Contains(t, result, "## Sheet2")
		assert.Contains(t, result, "| Item | Price |")
	})

	t.Run("file not found", func(t *testing.T) {
		_, err := Convert("nonexistent.xlsx", ConvertOptions{})
		assert.Error(t, err)
	})
}

func TestConvertWithTestdata(t *testing.T) {
	samplePath := filepath.Join("..", "..", "testdata", "sample.xlsx")
	if _, err := os.Stat(samplePath); os.IsNotExist(err) {
		t.Skip("testdata/sample.xlsx not found, skipping")
	}

	result, err := Convert(samplePath, ConvertOptions{})
	assert.NoError(t, err)
	assert.NotEmpty(t, result)
}
