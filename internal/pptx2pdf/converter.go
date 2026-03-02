package pptx2pdf

import (
	"bytes"
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"os"
	"path/filepath"
	"strings"

	"github.com/go-pdf/fpdf"

	"ar-tools/internal/pptx2md"
)

// ConvertOptions holds configuration for pptx to pdf conversion.
type ConvertOptions struct{}

const (
	pageW     = 297.0 // A4 landscape width (mm)
	pageH     = 210.0 // A4 landscape height (mm)
	margin    = 15.0
	contentW  = pageW - 2*margin
	titleSize = 20.0
	bodySize  = 12.0
	tableSize = 10.0
	titleLH   = 8.0 // title line height (mm)
	bodyLH    = 6.0 // body line height (mm)
	tableLH   = 5.0 // table cell line height (mm)
)

// Convert reads a .pptx file and produces a PDF in the same directory.
// Returns the output PDF file path.
func Convert(filePath string, opts ConvertOptions) (string, error) {
	pres, err := pptx2md.Parse(filePath)
	if err != nil {
		return "", err
	}
	defer pres.Close()

	fontPath, fontName := findSystemFont()

	pdf := fpdf.New("L", "mm", "A4", "")
	pdf.SetMargins(margin, margin, margin)
	pdf.SetAutoPageBreak(false, margin)

	if fontPath != "" {
		pdf.AddUTF8Font(fontName, "", fontPath)
	} else {
		fontName = "Helvetica"
	}

	totalSlides := len(pres.Slides)
	if totalSlides == 0 {
		pdf.AddPage()
	}

	for _, slide := range pres.Slides {
		pdf.AddPage()
		renderSlide(pdf, pres, slide, fontName, totalSlides)
	}

	if pdf.Err() {
		return "", fmt.Errorf("pdf generation error: %w", pdf.Error())
	}

	outPath := strings.TrimSuffix(filePath, filepath.Ext(filePath)) + ".pdf"
	if err := pdf.OutputFileAndClose(outPath); err != nil {
		return "", fmt.Errorf("failed to write PDF: %w", err)
	}

	return outPath, nil
}

// findSystemFont probes common Windows font paths and returns
// the first loadable TTF along with its family name.
func findSystemFont() (fontPath string, fontName string) {
	fontsDir := windowsFontsDir()
	candidates := []string{
		"simhei.ttf",
		"simkai.ttf",
		"KAIU.TTF",
		"arial.ttf",
		"calibri.ttf",
	}

	for _, fname := range candidates {
		fp := filepath.Join(fontsDir, fname)
		if _, err := os.Stat(fp); err != nil {
			continue
		}
		if !isTTF(fp) {
			continue
		}
		name := strings.TrimSuffix(strings.ToLower(fname), ".ttf")
		return fp, name
	}
	return "", ""
}

// isTTF checks that the file starts with a valid TrueType/OpenType signature
// (not a TTC collection, which most Go PDF libraries cannot load directly).
func isTTF(path string) bool {
	f, err := os.Open(path)
	if err != nil {
		return false
	}
	defer f.Close()

	header := make([]byte, 4)
	if _, err := f.Read(header); err != nil {
		return false
	}
	if header[0] == 0 && header[1] == 1 && header[2] == 0 && header[3] == 0 {
		return true
	}
	if string(header) == "OTTO" {
		return true
	}
	return false
}

func windowsFontsDir() string {
	if dir := os.Getenv("WINDIR"); dir != "" {
		return filepath.Join(dir, "Fonts")
	}
	return `C:\Windows\Fonts`
}

// ensureSpace checks if there's enough room on the current page;
// if not, adds a new page and resets y to margin.
func ensureSpace(pdf *fpdf.Fpdf, y *float64, needed float64) {
	if *y+needed > pageH-margin {
		pdf.AddPage()
		*y = margin
	}
}

func renderSlide(pdf *fpdf.Fpdf, pres *pptx2md.Presentation, slide *pptx2md.Slide, fontName string, totalSlides int) {
	y := margin

	// Slide number (top-right corner)
	pdf.SetFont(fontName, "", 9)
	pdf.SetTextColor(150, 150, 150)
	slideLabel := fmt.Sprintf("%d / %d", slide.Index, totalSlides)
	labelW := pdf.GetStringWidth(slideLabel)
	pdf.SetXY(pageW-margin-labelW, margin-3)
	pdf.CellFormat(labelW, 4, slideLabel, "", 0, "R", false, 0, "")
	pdf.SetTextColor(0, 0, 0)

	// Title
	title := slide.Title
	if title == "" {
		title = fmt.Sprintf("Slide %d", slide.Index)
	}
	pdf.SetFont(fontName, "", titleSize)
	pdf.SetXY(margin, y)
	pdf.MultiCell(contentW, titleLH, title, "", "L", false)
	y = pdf.GetY() + 3

	// Separator line
	pdf.SetDrawColor(200, 200, 200)
	pdf.Line(margin, y, margin+contentW, y)
	y += 4

	// Body text
	pdf.SetFont(fontName, "", bodySize)
	for _, body := range slide.Bodies {
		ensureSpace(pdf, &y, bodyLH*2)
		pdf.SetXY(margin, y)
		pdf.MultiCell(contentW, bodyLH, body, "", "L", false)
		y = pdf.GetY() + 2
	}

	// Tables
	for _, tbl := range slide.Tables {
		y = renderTable(pdf, tbl, fontName, y)
	}

	// Images
	for _, img := range slide.Images {
		if img.MediaPath == "" {
			continue
		}
		imgType := detectImageType(img.MediaPath)
		if imgType == "" {
			continue
		}
		data, err := pres.ReadMedia(img.MediaPath)
		if err != nil {
			continue
		}

		cfg, _, err := image.DecodeConfig(bytes.NewReader(data))
		if err != nil {
			continue
		}

		// Calculate scaled size before deciding on space
		w, h := scaleImage(float64(cfg.Width), float64(cfg.Height), contentW, pageH-2*margin)
		ensureSpace(pdf, &y, h+2)

		// Re-scale for remaining space on current page
		maxH := pageH - y - margin
		if h > maxH {
			w, h = scaleImage(float64(cfg.Width), float64(cfg.Height), contentW, maxH)
		}

		imgName := fmt.Sprintf("s%d_%s", slide.Index, filepath.Base(img.MediaPath))
		reader := bytes.NewReader(data)
		imgOpts := fpdf.ImageOptions{ImageType: imgType, ReadDpi: true}
		pdf.RegisterImageOptionsReader(imgName, imgOpts, reader)
		if pdf.Ok() {
			pdf.ImageOptions(imgName, margin, y, w, h, false, imgOpts, 0, "")
			y += h + 2
		}
	}
}

func renderTable(pdf *fpdf.Fpdf, tbl pptx2md.Table, fontName string, y float64) float64 {
	if len(tbl.Rows) == 0 {
		return y
	}

	ensureSpace(pdf, &y, tableLH*3)
	y += 2

	numCols := 0
	for _, row := range tbl.Rows {
		if len(row) > numCols {
			numCols = len(row)
		}
	}
	if numCols == 0 {
		return y
	}

	colW := contentW / float64(numCols)
	pdf.SetFont(fontName, "", tableSize)
	pdf.SetDrawColor(180, 180, 180)

	for rowIdx, row := range tbl.Rows {
		// Calculate row height based on tallest cell
		rowH := tableLH
		for _, cell := range row {
			lines := pdf.SplitText(cell, colW-2)
			cellH := float64(len(lines)) * tableLH
			if cellH > rowH {
				rowH = cellH
			}
		}

		ensureSpace(pdf, &y, rowH+1)

		// Header row: light gray background
		if rowIdx == 0 {
			pdf.SetFillColor(240, 240, 240)
		}

		for colIdx := 0; colIdx < numCols; colIdx++ {
			x := margin + float64(colIdx)*colW
			cellText := ""
			if colIdx < len(row) {
				cellText = row[colIdx]
			}

			// Draw cell border
			pdf.Rect(x, y, colW, rowH, "D")
			if rowIdx == 0 {
				pdf.Rect(x, y, colW, rowH, "F")
				pdf.Rect(x, y, colW, rowH, "D")
			}

			// Draw text inside cell
			lines := pdf.SplitText(cellText, colW-2)
			for lineIdx, line := range lines {
				pdf.SetXY(x+1, y+float64(lineIdx)*tableLH)
				pdf.CellFormat(colW-2, tableLH, line, "", 0, "L", false, 0, "")
			}
		}
		y += rowH
	}

	return y + 3
}

func detectImageType(mediaPath string) string {
	switch strings.ToLower(filepath.Ext(mediaPath)) {
	case ".png":
		return "PNG"
	case ".jpg", ".jpeg":
		return "JPG"
	case ".gif":
		return "GIF"
	default:
		return ""
	}
}

func scaleImage(origW, origH, maxW, maxH float64) (float64, float64) {
	w := origW * 25.4 / 96.0
	h := origH * 25.4 / 96.0

	if w > maxW {
		r := maxW / w
		w, h = maxW, h*r
	}
	if h > maxH {
		r := maxH / h
		w, h = w*r, maxH
	}
	return w, h
}
