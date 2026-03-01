package pptx2md

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"path"
	"sort"
	"strings"
)

// Slide represents a single parsed slide.
type Slide struct {
	Index  int
	Title  string
	Bodies []string   // text paragraphs (non-title)
	Images []ImageRef // image references
}

// ImageRef links an image to its media path inside the ZIP.
type ImageRef struct {
	RelID     string
	MediaPath string // e.g. "ppt/media/image1.png"
}

// Presentation holds all parsed slides and a handle to the ZIP for media extraction.
type Presentation struct {
	Slides []*Slide
	zip    *zip.ReadCloser
}

// Close releases the underlying ZIP reader.
func (p *Presentation) Close() error {
	if p.zip != nil {
		return p.zip.Close()
	}
	return nil
}

// ReadMedia reads a media file from the PPTX archive.
func (p *Presentation) ReadMedia(mediaPath string) ([]byte, error) {
	for _, f := range p.zip.File {
		if f.Name == mediaPath {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, fmt.Errorf("media not found: %s", mediaPath)
}

// Parse opens a .pptx file and extracts slide content.
func Parse(filePath string) (*Presentation, error) {
	zr, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open pptx: %w", err)
	}

	pres := &Presentation{zip: zr}

	slideOrder, err := getSlideOrder(zr)
	if err != nil {
		zr.Close()
		return nil, err
	}

	for i, slidePath := range slideOrder {
		slide, err := parseSlide(zr, slidePath, i+1)
		if err != nil {
			zr.Close()
			return nil, fmt.Errorf("failed to parse %s: %w", slidePath, err)
		}
		pres.Slides = append(pres.Slides, slide)
	}

	return pres, nil
}

// getSlideOrder determines slide ordering from presentation.xml and its rels.
func getSlideOrder(zr *zip.ReadCloser) ([]string, error) {
	presRels, err := parseRels(zr, "ppt/_rels/presentation.xml.rels")
	if err != nil {
		return nil, fmt.Errorf("failed to read presentation rels: %w", err)
	}

	presXML, err := readZipFile(zr, "ppt/presentation.xml")
	if err != nil {
		return nil, fmt.Errorf("failed to read presentation.xml: %w", err)
	}

	var pres xmlPresentation
	if err := xml.Unmarshal(presXML, &pres); err != nil {
		return nil, fmt.Errorf("failed to parse presentation.xml: %w", err)
	}

	var slides []string
	for _, sid := range pres.SlideIdList.SlideIds {
		relTarget, ok := presRels[sid.RID]
		if !ok {
			continue
		}
		// Resolve relative path: targets are relative to ppt/
		slidePath := resolveRelPath("ppt", relTarget)
		slides = append(slides, slidePath)
	}

	if len(slides) == 0 {
		// Fallback: scan for slide files directly
		slides = scanSlideFiles(zr)
	}

	return slides, nil
}

// scanSlideFiles finds slide XML files by scanning the ZIP.
func scanSlideFiles(zr *zip.ReadCloser) []string {
	var slides []string
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			if !strings.Contains(f.Name, "_rels") {
				slides = append(slides, f.Name)
			}
		}
	}
	sort.Strings(slides)
	return slides
}

func parseSlide(zr *zip.ReadCloser, slidePath string, index int) (*Slide, error) {
	data, err := readZipFile(zr, slidePath)
	if err != nil {
		return nil, err
	}

	var sld xmlSlide
	if err := xml.Unmarshal(data, &sld); err != nil {
		return nil, fmt.Errorf("failed to parse slide XML: %w", err)
	}

	// Parse rels for this slide
	relsPath := slideRelsPath(slidePath)
	slideRels, _ := parseRels(zr, relsPath)

	slide := &Slide{Index: index}

	// Extract shapes (text + images)
	for _, sp := range sld.CSld.SpTree.Shapes {
		extractShapeText(sp, slide)
	}

	// Extract grouped shapes
	for _, grp := range sld.CSld.SpTree.GroupShapes {
		for _, sp := range grp.Shapes {
			extractShapeText(sp, slide)
		}
		for _, pic := range grp.Pictures {
			extractPicture(pic, slide, slideRels)
		}
	}

	// Extract pictures
	for _, pic := range sld.CSld.SpTree.Pictures {
		extractPicture(pic, slide, slideRels)
	}

	return slide, nil
}

func extractShapeText(sp xmlShape, slide *Slide) {
	if sp.TxBody == nil {
		return
	}

	isTitle := isPlaceholderTitle(sp)
	for _, para := range sp.TxBody.Paragraphs {
		text := paragraphText(para)
		if text == "" {
			continue
		}
		if isTitle && slide.Title == "" {
			slide.Title = text
		} else {
			slide.Bodies = append(slide.Bodies, text)
		}
	}
}

func extractPicture(pic xmlPicture, slide *Slide, rels map[string]string) {
	if pic.BlipFill == nil || pic.BlipFill.Blip == nil {
		return
	}
	rID := pic.BlipFill.Blip.Embed
	if rID == "" {
		return
	}
	ref := ImageRef{RelID: rID}
	if target, ok := rels[rID]; ok {
		ref.MediaPath = resolveRelPath("ppt/slides", target)
	}
	slide.Images = append(slide.Images, ref)
}

func isPlaceholderTitle(sp xmlShape) bool {
	if sp.NvSpPr == nil || sp.NvSpPr.NvPr == nil || sp.NvSpPr.NvPr.Ph == nil {
		return false
	}
	phType := sp.NvSpPr.NvPr.Ph.Type
	return phType == "title" || phType == "ctrTitle"
}

func paragraphText(para xmlParagraph) string {
	var parts []string
	for _, run := range para.Runs {
		if run.Text != "" {
			parts = append(parts, run.Text)
		}
	}
	return strings.Join(parts, "")
}

// --- Rels parsing ---

func parseRels(zr *zip.ReadCloser, relsPath string) (map[string]string, error) {
	data, err := readZipFile(zr, relsPath)
	if err != nil {
		return nil, err
	}

	var rels xmlRelationships
	if err := xml.Unmarshal(data, &rels); err != nil {
		return nil, err
	}

	m := make(map[string]string, len(rels.Relationships))
	for _, r := range rels.Relationships {
		m[r.ID] = r.Target
	}
	return m, nil
}

func slideRelsPath(slidePath string) string {
	dir := path.Dir(slidePath)
	base := path.Base(slidePath)
	return dir + "/_rels/" + base + ".rels"
}

func resolveRelPath(base, target string) string {
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(target, "/")
	}
	// Handle "../media/image1.png" style paths
	combined := base + "/" + target
	return path.Clean(combined)
}

// --- ZIP helpers ---

func readZipFile(zr *zip.ReadCloser, name string) ([]byte, error) {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, fmt.Errorf("file not found in zip: %s", name)
}

// --- XML structures ---

type xmlPresentation struct {
	XMLName     xml.Name       `xml:"presentation"`
	SlideIdList xmlSlideIdList `xml:"sldIdLst"`
}

type xmlSlideIdList struct {
	SlideIds []xmlSlideId `xml:"sldId"`
}

type xmlSlideId struct {
	RID string `xml:"id,attr"`
}

type xmlSlide struct {
	XMLName xml.Name `xml:"sld"`
	CSld    xmlCSld  `xml:"cSld"`
}

type xmlCSld struct {
	SpTree xmlSpTree `xml:"spTree"`
}

type xmlSpTree struct {
	Shapes      []xmlShape      `xml:"sp"`
	Pictures    []xmlPicture    `xml:"pic"`
	GroupShapes []xmlGroupShape `xml:"grpSp"`
}

type xmlGroupShape struct {
	Shapes   []xmlShape   `xml:"sp"`
	Pictures []xmlPicture `xml:"pic"`
}

type xmlShape struct {
	NvSpPr *xmlNvSpPr `xml:"nvSpPr"`
	TxBody *xmlTxBody `xml:"txBody"`
}

type xmlNvSpPr struct {
	NvPr *xmlNvPr `xml:"nvPr"`
}

type xmlNvPr struct {
	Ph *xmlPh `xml:"ph"`
}

type xmlPh struct {
	Type string `xml:"type,attr"`
}

type xmlTxBody struct {
	Paragraphs []xmlParagraph `xml:"p"`
}

type xmlParagraph struct {
	Runs []xmlRun `xml:"r"`
}

type xmlRun struct {
	Text string `xml:"t"`
}

type xmlPicture struct {
	BlipFill *xmlBlipFill `xml:"blipFill"`
}

type xmlBlipFill struct {
	Blip *xmlBlip `xml:"blip"`
}

type xmlBlip struct {
	Embed string `xml:"embed,attr"`
}

type xmlRelationships struct {
	XMLName       xml.Name          `xml:"Relationships"`
	Relationships []xmlRelationship `xml:"Relationship"`
}

type xmlRelationship struct {
	ID     string `xml:"Id,attr"`
	Target string `xml:"Target,attr"`
	Type   string `xml:"Type,attr"`
}
