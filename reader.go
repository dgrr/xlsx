package xlsx

import (
	"archive/zip"
	"errors"
	"fmt"
	"io"
	"strings"

	"github.com/dgrr/xml"
)

// XLSX ...
type XLSX struct {
	sharedStrings []string
	zr            *zip.ReadCloser
	Sheets        []*Sheet
}

// sheetData
//   row: r="1"
//     c: r="A1" t="inlineStr"|"n" s="1"
//       is:
//         t:
//       v:

// xlsxIndex represents an index of the sharedStrings file and
// the spreadsheets files.
type xlsxIndex struct {
	sharedStr string
	files     []string
}

// Close closes all the buffers and readers.
func (xlsx *XLSX) Close() error {
	if xlsx.zr == nil {
		return nil
	}
	return xlsx.zr.Close()
}

// Open just opens the file for reading.
func Open(filename string) (*XLSX, error) {
	zr, err := zip.OpenReader(filename)
	if err == nil {
		for _, zFile := range zr.File {
			// read where the worksheets are
			if zFile.Name == "[Content_Types].xml" {
				index, err := parseContentType(zFile)
				if err != nil {
					return nil, err
				}

				// read the worksheets
				return extractWorksheets(zr, &index)
			}
		}
	}

	return nil, errors.New("rels file not found")
}

func getPartName(e *xml.StartElement) (partName string, err error) {
	for _, kv := range e.Attrs {
		if kv.K == "PartName" {
			partName = kv.V
			break
		}
	}
	if partName == "" {
		err = errors.New("PartName parameter not found")
	}
	return
}

func parseContentType(zFile *zip.File) (index xlsxIndex, err error) {
	var (
		zfr io.ReadCloser
	)

	zfr, err = zFile.Open()
	if err != nil {
		return
	}
	defer zfr.Close()

	r := xml.NewReader(zfr)
	for r.Next() {
		switch e := r.Element().(type) {
		case *xml.StartElement:
			if e.Name != "Override" {
				continue
			}
			var partName string
			for _, kv := range e.Attrs {
				if kv.K == "ContentType" {
					switch kv.V {
					case "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml":
						partName, err = getPartName(e)
						if err == nil {
							index.files = append(index.files, partName)
						}
					case "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml":
						partName, err = getPartName(e)
						if err == nil {
							index.sharedStr = partName
						}
					}
					break
				}
			}
		}
	}
	if err == nil && len(index.files) == 0 {
		err = errors.New("no data files found")
	}

	return
}

func extractWorksheets(zr *zip.ReadCloser, index *xlsxIndex) (*XLSX, error) {
	var (
		err    error
		shared []string
	)
	sharedFile := index.sharedStr

	if len(sharedFile) > 0 {
		shared, err = readShared(zr, sharedFile)
		if err != nil {
			return nil, err
		}
	}

	xs := new(XLSX)
	xs.sharedStrings = shared

	for _, filename := range index.files {
		zFile, err := getZipFile(zr, filename)
		if err != nil {
			xs = nil
			return nil, err
		}

		xs.Sheets = append(xs.Sheets, &Sheet{
			parent: xs,
			zFile:  zFile,
		})
	}

	return xs, err
}

func findNameIn(name, where string) bool {
	if name[0] == '/' {
		return name[1:] == where
	}
	return strings.Contains(where, name)
}

func getZipFile(zr *zip.ReadCloser, filename string) (zFile *zip.File, err error) {
	var found = false
	for _, zFile = range zr.File {
		found = findNameIn(filename, zFile.Name)
		if found {
			break
		}
	}
	if !found {
		err = fmt.Errorf("%s not found", filename)
	}

	return zFile, err
}

func readShared(zr *zip.ReadCloser, filename string) ([]string, error) {
	var (
		rc    io.ReadCloser
		found bool
		err   error
	)
	for _, zFile := range zr.File {
		found = findNameIn(filename, zFile.Name)
		if found {
			rc, err = zFile.Open()
			break
		}
	}
	if !found {
		err = fmt.Errorf("%s not found", filename)
	}
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	ss := make([]string, 0)
	r := xml.NewReader(rc)
	T := false
loop:
	for r.Next() {
		switch e := r.Element().(type) {
		case *xml.StartElement:
			if e.Name == "t" {
				T = true
			} else {
				T = false
			}
		case *xml.TextElement:
			if T {
				ss = append(ss, string(*e))
			}
		case *xml.EndElement:
			if e.Name == "sst" {
				break loop
			}
		}
	}

	return ss, err
}
