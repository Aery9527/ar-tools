package pptx2pdf

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestConvert_SamplePptx(t *testing.T) {
	src := filepath.Join("..", "..", "testdata", "sample.pptx")

	tmpDir := t.TempDir()
	data, err := os.ReadFile(src)
	assert.NoError(t, err)
	tmpFile := filepath.Join(tmpDir, "sample.pptx")
	assert.NoError(t, os.WriteFile(tmpFile, data, 0644))

	outPath, err := Convert(tmpFile, ConvertOptions{})
	assert.NoError(t, err)

	expected := filepath.Join(tmpDir, "sample.pdf")
	assert.Equal(t, expected, outPath)

	info, err := os.Stat(outPath)
	assert.NoError(t, err)
	assert.True(t, info.Size() > 0)

	// Verify PDF header
	pdfData, err := os.ReadFile(outPath)
	assert.NoError(t, err)
	assert.True(t, len(pdfData) >= 5)
	assert.Equal(t, "%PDF-", string(pdfData[:5]))
}

func TestConvert_FileNotFound(t *testing.T) {
	_, err := Convert("nonexistent.pptx", ConvertOptions{})
	assert.Error(t, err)
}
