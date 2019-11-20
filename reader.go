package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"net/url"
	"path/filepath"
	"strconv"
	"strings"
)

// XLSX ...
type XLSX struct {
	Cells [][]string
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
	Count   int                `xml:"count,attr"`
	Strings []xlsxSharedString `xml:"si"`
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

type xlsxWorksheet struct {
	sharedStrings []string
	XMLName       xml.Name      `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	SheetData     xlsxSheetData `xml:"sheetData"`
}

type xlsxSheetData struct {
	XMLName xml.Name  `xml:"sheetData"`
	Row     []xlsxRow `xml:"row"`
}

type xlsxRow struct {
	R int     `xml:"r,attr"`
	C []xlsxC `xml:"c"`
}

type xlsxC struct {
	XMLName xml.Name
	T       string  `xml:"t,attr,omitempty"` // can be `inlineStr`, `n`, `s`
	V       string  `xml:"v,omitempty"`      // Value
	Is      *xlsxIS `xml:"is,omitempty"`     // inline string
}

type xlsxIS struct {
	XMLName xml.Name
	T       string `xml:"t"` // value of the inline string
}

// ReadFrom ...
func ReadFrom(filename string) (*XLSX, error) {
	zr, err := zip.OpenReader(filename)
	if err == nil {
		defer zr.Close()
		for _, zFile := range zr.File {
			if strings.Contains(zFile.Name, "workbook.xml.rels") {
				index, err := parseRels(zFile)
				if err != nil {
					return nil, err
				}
				return readFrom(zr, index)
			}
		}
	}

	return nil, errors.New("not found")
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
	//sharedStr :=

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
		err = errors.New("data not found")
	}

	return
}

func readFrom(zr *zip.ReadCloser, index xlsxIndex) (*XLSX, error) {
	var (
		err    error
		shared []string
	)
	sharedFile := index.sharedStr
	filename := index.files[0]

	if len(sharedFile) > 0 {
		shared, err = readShared(zr, sharedFile)
	}
	if err != nil {
		return nil, err
	}

	ws, err := readWorksheet(zr, filename)
	if err == nil {
		return newXLSX(ws, shared)
	}

	return nil, err
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

func readWorksheet(zr *zip.ReadCloser, filename string) (*xlsxWorksheet, error) {
	var (
		rc    io.ReadCloser
		found bool
		err   error
	)
	for _, zfr := range zr.File {
		if filename[0] == '/' {
			found = zfr.Name == filename[1:]
		} else {
			found = strings.Contains(zfr.Name, filename)
		}
		if found {
			rc, err = zfr.Open()
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

	ws := new(xlsxWorksheet)
	ec := xml.NewDecoder(rc)

	return ws, ec.Decode(ws)
}

func newXLSX(ws *xlsxWorksheet, shared []string) (xs *XLSX, err error) {
	if ws == nil {
		panic("ws cannot be nil")
	}
	xs = new(XLSX)
	xs.Cells = make([][]string, len(ws.SheetData.Row))
	for i, row := range ws.SheetData.Row {
		xs.Cells[i] = make([]string, len(row.C))
		for j, c := range row.C {
			switch c.T {
			case "inlineStr": // inline string
				xs.Cells[i][j] = c.Is.T
			case "s": // shared string
				idx, err := strconv.Atoi(c.V)
				if err == nil && idx < len(shared) {
					xs.Cells[i][j] = shared[idx]
				}
			default: // "n"
				xs.Cells[i][j] = c.V
			}
		}
	}

	return
}
