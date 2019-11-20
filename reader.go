package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"net/url"
	"path/filepath"
	"strings"
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

type xlsxSharedStrings struct {
	XMLName xml.Name           `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sst"`
	Strings []xlsxSharedString `xml:"si"`
	//Count   int                `xml:"count,attr"`
}

type xlsxSharedString struct {
	T string `xml:"t"`
}

type xlsxWorkbookRels struct {
	XMLName       xml.Name               `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []xlsxWorkbookRelation `xml:"Relationship"`
}

type xlsxWorkbookRelation struct {
	Target string `xml:"Target,attr"`
	Type   string `xml:"Type,attr"`
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
			if strings.Contains(zFile.Name, "workbook.xml.rels") {
				index, err := parseRels(zFile)
				if err != nil {
					return nil, err
				}

				// read the worksheets
				return extractWorksheets(zr, index)
			}
		}
	}

	return nil, errors.New("rels file not found")
}

func parseRels(zFile *zip.File) (index xlsxIndex, err error) {
	var (
		zfr io.ReadCloser
		uri *url.URL
	)

	zfr, err = zFile.Open()
	if err != nil {
		return
	}
	defer zfr.Close()

	rels := xlsxWorkbookRels{}

	dec := xml.NewDecoder(zfr)
	err = dec.Decode(&rels)
	if err == nil {
		for _, r := range rels.Relationships {
			uri, err = url.Parse(r.Type)
			if err != nil {
				return // TODO: Do not return yet, try other
			}
			switch filepath.Base(uri.Path) {
			case "worksheet":
				index.files = append(index.files, r.Target)
			case "sharedStrings":
				index.sharedStr = r.Target
			}
		}
	}
	if len(index.files) == 0 {
		err = errors.New("no data files found")
	}

	return
}

func extractWorksheets(zr *zip.ReadCloser, index xlsxIndex) (*XLSX, error) {
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
		return xs, nil
	}

	return nil, err
}

func getZipFile(zr *zip.ReadCloser, filename string) (zFile *zip.File, err error) {
	var found = false
	for _, zFile = range zr.File {
		if filename[0] == '/' {
			found = zFile.Name == filename[1:]
		} else {
			found = strings.Contains(zFile.Name, filename)
		}
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
		if found = strings.Contains(zFile.Name, filename); found {
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

	ss := xlsxSharedStrings{}
	ec := xml.NewDecoder(rc)

	err = ec.Decode(&ss)
	if err == nil {
		sstr := make([]string, len(ss.Strings))
		for i := range ss.Strings {
			sstr[i] = ss.Strings[i].T
		}
		return sstr, nil
	}

	return nil, err
}
