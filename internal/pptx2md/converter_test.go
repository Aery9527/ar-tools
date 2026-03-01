package pptx2md

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestParse_SamplePptx(t *testing.T) {
	pres, err := Parse(filepath.Join("..", "..", "testdata", "sample.pptx"))
	assert.NoError(t, err)
	defer pres.Close()

	assert.Len(t, pres.Slides, 2)

	// Slide 1: title slide
	s1 := pres.Slides[0]
	assert.Equal(t, 1, s1.Index)
	assert.Equal(t, "Welcome to ar-tools", s1.Title)
	assert.Contains(t, s1.Bodies, "A developer toolkit")

	// Slide 2: content + image
	s2 := pres.Slides[1]
	assert.Equal(t, 2, s2.Index)
	assert.Equal(t, "Features Overview", s2.Title)
	assert.Contains(t, s2.Bodies, "Excel to Markdown conversion")
	assert.Contains(t, s2.Bodies, "PowerPoint to Markdown conversion")
	assert.Len(t, s2.Images, 1)
	assert.Equal(t, "ppt/media/image1.png", s2.Images[0].MediaPath)
}

func TestParse_FileNotFound(t *testing.T) {
	_, err := Parse("nonexistent.pptx")
	assert.Error(t, err)
}

func TestConvert_SamplePptx(t *testing.T) {
	src := filepath.Join("..", "..", "testdata", "sample.pptx")

	// Copy to temp dir to avoid polluting testdata
	tmpDir := t.TempDir()
	data, err := os.ReadFile(src)
	assert.NoError(t, err)
	tmpFile := filepath.Join(tmpDir, "sample.pptx")
	assert.NoError(t, os.WriteFile(tmpFile, data, 0644))

	result, err := Convert(tmpFile, ConvertOptions{})
	assert.NoError(t, err)

	md := result.Markdown

	// Check slide headings
	assert.Contains(t, md, "## Welcome to ar-tools")
	assert.Contains(t, md, "## Features Overview")

	// Check body text
	assert.Contains(t, md, "A developer toolkit")
	assert.Contains(t, md, "Excel to Markdown conversion")
	assert.Contains(t, md, "PowerPoint to Markdown conversion")

	// Check image link
	assert.Contains(t, md, "![image1.png](./sample_images/image1.png)")

	// Check image file was exported
	assert.NotEmpty(t, result.ImageDir)
	imgPath := filepath.Join(result.ImageDir, "image1.png")
	info, err := os.Stat(imgPath)
	assert.NoError(t, err)
	assert.True(t, info.Size() > 0)
}

func TestConvert_CustomImageDir(t *testing.T) {
	src := filepath.Join("..", "..", "testdata", "sample.pptx")

	tmpDir := t.TempDir()
	data, err := os.ReadFile(src)
	assert.NoError(t, err)
	tmpFile := filepath.Join(tmpDir, "test.pptx")
	assert.NoError(t, os.WriteFile(tmpFile, data, 0644))

	result, err := Convert(tmpFile, ConvertOptions{ImageDir: "my_pics"})
	assert.NoError(t, err)

	assert.Contains(t, result.Markdown, "![image1.png](./my_pics/image1.png)")
	assert.DirExists(t, filepath.Join(tmpDir, "my_pics"))
}

func TestReadMedia(t *testing.T) {
	pres, err := Parse(filepath.Join("..", "..", "testdata", "sample.pptx"))
	assert.NoError(t, err)
	defer pres.Close()

	data, err := pres.ReadMedia("ppt/media/image1.png")
	assert.NoError(t, err)
	assert.True(t, len(data) > 0)

	// PNG signature check
	assert.Equal(t, byte(0x89), data[0])
	assert.Equal(t, byte('P'), data[1])
	assert.Equal(t, byte('N'), data[2])
	assert.Equal(t, byte('G'), data[3])
}

func TestReadMedia_NotFound(t *testing.T) {
	pres, err := Parse(filepath.Join("..", "..", "testdata", "sample.pptx"))
	assert.NoError(t, err)
	defer pres.Close()

	_, err = pres.ReadMedia("ppt/media/nonexistent.png")
	assert.Error(t, err)
}
