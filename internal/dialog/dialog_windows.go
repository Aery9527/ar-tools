//go:build windows

package dialog

import (
	"path/filepath"
	"runtime"
	"syscall"
	"unicode/utf16"
	"unsafe"
)

var (
	comdlg32             = syscall.NewLazyDLL("comdlg32.dll")
	procGetOpenFileNameW = comdlg32.NewProc("GetOpenFileNameW")
)

const (
	ofnAllowMultiSelect = 0x00000200
	ofnExplorer         = 0x00080000
	ofnFileMustExist    = 0x00001000
	ofnNoChangeDir      = 0x00000008
	maxFileBuf          = 65536
)

type openFileNameW struct {
	lStructSize       uint32
	hwndOwner         uintptr
	hInstance         uintptr
	lpstrFilter       *uint16
	lpstrCustomFilter *uint16
	nMaxCustFilter    uint32
	nFilterIndex      uint32
	lpstrFile         *uint16
	nMaxFile          uint32
	lpstrFileTitle    *uint16
	nMaxFileTitle     uint32
	lpstrInitialDir   *uint16
	lpstrTitle        *uint16
	flags             uint32
	nFileOffset       uint16
	nFileExtension    uint16
	lpstrDefExt       *uint16
	lCustData         uintptr
	lpfnHook          uintptr
	lpTemplateName    *uint16
	pvReserved        uintptr
	dwReserved        uint32
	flagsEx           uint32
}

// OpenMultipleFiles shows a native Windows file open dialog with multi-select.
// Returns selected file paths, or nil if the user cancelled.
func OpenMultipleFiles(title string, filterName string, filterPattern string) ([]string, error) {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	fileBuf := make([]uint16, maxFileBuf)
	titlePtr, _ := syscall.UTF16PtrFromString(title)
	filter := makeFilter(filterName, filterPattern)

	ofn := openFileNameW{
		lpstrFile:   &fileBuf[0],
		nMaxFile:    uint32(len(fileBuf)),
		lpstrTitle:  titlePtr,
		lpstrFilter: &filter[0],
		flags:       ofnAllowMultiSelect | ofnExplorer | ofnFileMustExist | ofnNoChangeDir,
	}
	ofn.lStructSize = uint32(unsafe.Sizeof(ofn))

	ret, _, _ := procGetOpenFileNameW.Call(uintptr(unsafe.Pointer(&ofn)))
	if ret == 0 {
		// User cancelled or error; treat both as cancel
		return nil, nil
	}

	return parseMultiSelect(fileBuf), nil
}

// makeFilter builds a Windows file filter string (double-null terminated).
func makeFilter(displayName, pattern string) []uint16 {
	// Format: "Display Name\0Pattern\0\0"
	raw := displayName + "\x00" + pattern + "\x00"
	result := utf16.Encode([]rune(raw))
	result = append(result, 0) // trailing null
	return result
}

// parseMultiSelect parses the file buffer returned by GetOpenFileNameW.
// Single selection: "C:\path\file.xlsx\0\0"
// Multi selection:  "C:\directory\0file1.xlsx\0file2.xlsx\0\0"
func parseMultiSelect(buf []uint16) []string {
	var parts []string
	start := 0
	for i, c := range buf {
		if c == 0 {
			if i == start {
				break
			}
			parts = append(parts, syscall.UTF16ToString(buf[start:i]))
			start = i + 1
		}
	}

	if len(parts) == 0 {
		return nil
	}
	if len(parts) == 1 {
		return parts
	}

	// Multiple files: first part is directory, rest are file names
	dir := parts[0]
	files := make([]string, len(parts)-1)
	for i, f := range parts[1:] {
		files[i] = filepath.Join(dir, f)
	}
	return files
}
