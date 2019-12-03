package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"strconv"
)

// Sheet represents an spreadsheet.
type Sheet struct {
	parent *XLSX
	zFile  *zip.File
}

type xlsxRow struct {
	R int     `xml:"r,attr,omitempty"` // Row number
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

// Open opens a sheet to read it.
func (s *Sheet) Open() (*SheetReader, error) {
	rc, err := s.zFile.Open()
	if err == nil {
		sr := &SheetReader{
			s:  s,
			rc: rc,
			ec: xml.NewDecoder(rc),
		}
		return sr, sr.skip()
	}
	return nil, err
}

// SheetReader creates an structure able to read row by row
// the spreadsheet data.
type SheetReader struct {
	s   *Sheet
	rc  io.ReadCloser
	ec  *xml.Decoder
	row []string
	err error
}

// Error returns the error occurred during Next().
//
// If no error is returned here but Next() returned false it can
// be caused because the EOF was reach.
func (sr *SheetReader) Error() error {
	if sr.err == io.EOF {
		return nil
	}
	return sr.err
}

func (sr *SheetReader) skip() error {
loop:
	for {
		tk, err := sr.ec.Token()
		if err != nil {
			if err != io.EOF {
				sr.err = err
			}
			return sr.err
		}
		switch stk := tk.(type) {
		case xml.StartElement:
			if stk.Name.Local == "sheetData" {
				break loop
			}
		}
	}

	return nil
}

// Next returns true if the row has been successfully readed.
//
// if false is returned check the Error() function.
func (sr *SheetReader) Next() bool {
	var (
		err error
		tk  xml.Token
	)
	sr.row = sr.row[:0]
loop:
	for sr.err == nil {
		tk, err = sr.ec.Token()
		if err != nil {
			sr.err = err
			break
		}

		switch stk := tk.(type) {
		case xml.StartElement:
			if stk.Name.Local != "row" {
				continue
			}

			row := xlsxRow{}
			shared := sr.s.parent.sharedStrings
			sr.err = sr.ec.DecodeElement(&row, &stk)
			if sr.err == nil {
				// TODO: Check the `r` parameter in rows.
				for _, c := range row.C {
					switch c.T {
					case "inlineStr": // inline string
						sr.row = append(sr.row, c.Is.T)
					case "s": // shared string
						idx, err := strconv.Atoi(c.V)
						if err == nil && idx < len(shared) {
							sr.row = append(sr.row, shared[idx])
						}
					default: // "n"
						f, err := strconv.ParseFloat(c.V, 64)
						if err == nil {
							sr.row = append(sr.row, strconv.FormatFloat(f, 'f', -1, 64))
						} else {
							sr.row = append(sr.row, c.V)
						}
					}
				}
			}
			break loop
		case xml.EndElement:
			if stk.Name.Local == "sheetData" {
				sr.err = io.EOF
			}
		}
	}

	return sr.err == nil
}

// Row returns the last readed row.
func (sr *SheetReader) Row() []string {
	return sr.row
}

// Close closes the sheet file reader.
func (sr *SheetReader) Close() error {
	return sr.rc.Close()
}
