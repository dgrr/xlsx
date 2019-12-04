package xlsx

import (
	"archive/zip"
	"fmt"
	"io"
	"strconv"

	"github.com/dgrr/xml"
)

// Sheet represents an spreadsheet.
type Sheet struct {
	parent *XLSX
	zFile  *zip.File
}

type xlsxRow struct {
	R int // Row number
	C []xlsxC
}

type xlsxC struct {
	T  string  // can be `inlineStr`, `n`, `s`
	V  string  // Value
	Is *xlsxIS // inline string
}

type xlsxIS struct {
	T string // value of the inline string
}

// Open opens a sheet to read it.
func (s *Sheet) Open() (*SheetReader, error) {
	rc, err := s.zFile.Open()
	if err == nil {
		sr := &SheetReader{
			s:  s,
			rc: rc,
			r:  xml.NewReader(rc),
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
	r   *xml.Reader
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

// skip will skips all the irrelevant fields
func (sr *SheetReader) skip() error {
loop:
	for sr.r.Next() {
		switch e := sr.r.Element().(type) {
		case *xml.StartElement:
			if e.Name == "sheetData" {
				break loop
			}
		}
	}

	return sr.r.Err()
}

func parseIntOrZero(s string) int {
	n, err := strconv.Atoi(s)
	if err != nil {
		n = 0
	}
	return n
}

// Next returns true if the row has been successfully readed.
//
// if false is returned check the Error() function.
func (sr *SheetReader) Next() bool {
	sr.row = sr.row[:0]
loop:
	for sr.r.Next() {
		switch e := sr.r.Element().(type) {
		case *xml.StartElement:
			if e.Name != "row" {
				continue
			}

			row := xlsxRow{}
			for _, kv := range e.Attrs {
				if kv.K == "r" {
					row.R = parseIntOrZero(kv.V)
				}
			}

			shared := sr.s.parent.sharedStrings
			sr.err = sr.decodeRow(&row)
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
		case *xml.EndElement:
			if e.Name == "sheetData" {
				sr.err = io.EOF
			}
		}
		if sr.err != nil {
			break
		}
	}
	if sr.err == nil && sr.r.Err() != nil {
		sr.err = sr.r.Err()
	}

	return sr.err == nil
}

func (sr *SheetReader) decodeRow(row *xlsxRow) error {
	c := xlsxC{}
loop:
	for sr.r.Next() {
		switch e := sr.r.Element().(type) {
		case *xml.StartElement:
			switch e.Name {
			case "c":
				for _, kv := range e.Attrs {
					if kv.K == "t" {
						c.T = kv.V
						break
					}
				}
			case "is":
				c.Is = new(xlsxIS)
			case "t", "v":
			default:
				return fmt.Errorf("unexpected element: `%s` when looking for a `c`", e.Name)
			}
		case *xml.TextElement:
			if c.Is != nil {
				c.Is.T = string(*e)
			} else {
				c.V = string(*e)
			}

			row.C = append(row.C, c)
			c = xlsxC{}
		case *xml.EndElement:
			if e.Name == "row" {
				break loop
			}
		}
	}
	return nil
}

// Row returns the last readed row.
func (sr *SheetReader) Row() []string {
	return sr.row
}

// Close closes the sheet file reader.
func (sr *SheetReader) Close() error {
	return sr.rc.Close()
}
