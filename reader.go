package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
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

type xlsxContentType struct {
	XMLName xml.Name   `xml:"http://schemas.openxmlformats.org/package/2006/content-types Types"`
	Types   []xlsxType `xml:"Override"`
}

type xlsxType struct {
	PartName    string `xml:"PartName,attr"`    // file
	ContentType string `xml:"ContentType,attr"` // url
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
				return extractWorksheets(zr, index)
			}
		}
	}

	return nil, errors.New("rels file not found")
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

	ct := xlsxContentType{}

	dec := xml.NewDecoder(zfr)
	err = dec.Decode(&ct)
	if err == nil {
		for _, ts := range ct.Types {
			switch ts.ContentType {
			case "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml":
				index.files = append(index.files, ts.PartName)
			case "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml":
				index.sharedStr = ts.PartName
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
