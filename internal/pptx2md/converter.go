package pptx2md

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// ConvertOptions holds configuration for pptx to markdown conversion.
type ConvertOptions struct {
	// ImageDir overrides the output image directory name. Empty uses default "{basename}_images".
	ImageDir string
}

// ConvertResult holds the conversion output.
type ConvertResult struct {
	Markdown string
	ImageDir string // actual image directory path (empty if no images)
}

// Convert reads a .pptx file and returns Markdown with extracted images.
func Convert(filePath string, opts ConvertOptions) (*ConvertResult, error) {
	pres, err := Parse(filePath)
	if err != nil {
		return nil, err
	}
	defer pres.Close()

	baseName := strings.TrimSuffix(filepath.Base(filePath), filepath.Ext(filePath))
	outDir := filepath.Dir(filePath)

	imageDir := opts.ImageDir
	if imageDir == "" {
		imageDir = baseName + "_images"
	}
	imageDirFull := filepath.Join(outDir, imageDir)

	// Collect all images first to know if we need to create the directory
	type imageExport struct {
		mediaPath string
		fileName  string
	}
	var images []imageExport
	imageNames := make(map[string]int) // track duplicates

	for _, slide := range pres.Slides {
		for _, img := range slide.Images {
			if img.MediaPath == "" {
				continue
			}
			fileName := filepath.Base(img.MediaPath)
			// Handle duplicate filenames
			if count, ok := imageNames[fileName]; ok {
				ext := filepath.Ext(fileName)
				name := strings.TrimSuffix(fileName, ext)
				fileName = fmt.Sprintf("%s_%d%s", name, count+1, ext)
				imageNames[filepath.Base(img.MediaPath)] = count + 1
			} else {
				imageNames[fileName] = 1
			}
			images = append(images, imageExport{
				mediaPath: img.MediaPath,
				fileName:  fileName,
			})
		}
	}

	// Export images if any
	if len(images) > 0 {
		if err := os.MkdirAll(imageDirFull, 0755); err != nil {
			return nil, fmt.Errorf("failed to create image dir: %w", err)
		}
		for _, img := range images {
			data, err := pres.ReadMedia(img.mediaPath)
			if err != nil {
				return nil, fmt.Errorf("failed to read media %s: %w", img.mediaPath, err)
			}
			outPath := filepath.Join(imageDirFull, img.fileName)
			if err := os.WriteFile(outPath, data, 0644); err != nil {
				return nil, fmt.Errorf("failed to write image %s: %w", outPath, err)
			}
		}
	}

	// Build markdown
	md := buildMarkdown(pres, imageDir)

	result := &ConvertResult{Markdown: md}
	if len(images) > 0 {
		result.ImageDir = imageDirFull
	}
	return result, nil
}

// ConvertToString is a convenience function that returns only the Markdown string.
func ConvertToString(filePath string, opts ConvertOptions) (string, error) {
	result, err := Convert(filePath, opts)
	if err != nil {
		return "", err
	}
	return result.Markdown, nil
}

func buildMarkdown(pres *Presentation, imageDir string) string {
	var sb strings.Builder

	// Build a mapping from media path to exported filename
	imageFileMap := make(map[string]string)
	imageNames := make(map[string]int)
	for _, slide := range pres.Slides {
		for _, img := range slide.Images {
			if img.MediaPath == "" {
				continue
			}
			if _, exists := imageFileMap[img.MediaPath]; exists {
				continue
			}
			fileName := filepath.Base(img.MediaPath)
			if count, ok := imageNames[fileName]; ok {
				ext := filepath.Ext(fileName)
				name := strings.TrimSuffix(fileName, ext)
				fileName = fmt.Sprintf("%s_%d%s", name, count+1, ext)
				imageNames[filepath.Base(img.MediaPath)] = count + 1
			} else {
				imageNames[fileName] = 1
			}
			imageFileMap[img.MediaPath] = fileName
		}
	}

	for i, slide := range pres.Slides {
		if i > 0 {
			sb.WriteString("\n---\n\n")
		}

		// Slide heading
		if slide.Title != "" {
			sb.WriteString(fmt.Sprintf("## %s\n\n", slide.Title))
		} else {
			sb.WriteString(fmt.Sprintf("## Slide %d\n\n", slide.Index))
		}

		// Body text
		for _, body := range slide.Bodies {
			sb.WriteString(body + "\n\n")
		}

		// Images
		for _, img := range slide.Images {
			if img.MediaPath == "" {
				continue
			}
			fileName := imageFileMap[img.MediaPath]
			// Use forward slash for markdown compatibility
			imgPath := imageDir + "/" + fileName
			sb.WriteString(fmt.Sprintf("![%s](./%s)\n\n", fileName, imgPath))
		}
	}

	return strings.TrimRight(sb.String(), "\n") + "\n"
}
