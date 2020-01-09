package xlsx

import (
	"archive/zip"
	"bytes"
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
	ReuseRecords bool
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
			if e.NameUnsafe() == "sheetData" {
				xml.ReleaseStart(e)
				break loop
			}
			xml.ReleaseStart(e)
		}
	}

	return sr.r.Error()
}

func parseIntOrZero(s string) int {
	n, err := strconv.Atoi(s)
	if err != nil {
		n = 0
	}
	return n
}

func (sr *SheetReader) nextString(idx int) *string {
	if idx < cap(sr.row) {
		sr.row = sr.row[:idx+1]
	} else {
		sr.row = append(sr.row, make([]string, idx+1-cap(sr.row))...)
	}

	return &sr.row[idx]
}

// Next returns true if the row has been successfully readed.
//
// if false is returned check the Error() function.
func (sr *SheetReader) Next() bool {
	if sr.ReuseRecords {
		sr.row = sr.row[:0]
	} else {
		sr.row = nil
	}
	shared := sr.s.parent.sharedStrings
loop:
	for sr.r.Next() {
		switch e := sr.r.Element().(type) {
		case *xml.StartElement:
			if !bytes.Equal(e.NameBytes(), rowString) {
				xml.ReleaseStart(e)
				continue
			}

			xml.ReleaseStart(e)
			sr.err = sr.decodeRow(shared)
			break loop
		case *xml.EndElement:
			if e.NameUnsafe() == "sheetData" {
				sr.err = io.EOF
			}
			xml.ReleaseEnd(e)
		}
		if sr.err != nil {
			break
		}
	}
	if sr.err == nil && sr.r.Error() != nil {
		sr.err = sr.r.Error()
	}

	return sr.err == nil
}

func (sr *SheetReader) decodeRow(shared []string) error {
	var (
		T   []byte
		Is  bool
		idx int
		s   *string
	)
loop:
	for sr.r.Next() {
		switch e := sr.r.Element().(type) {
		case *xml.StartElement:
			switch e.NameUnsafe() {
			case "c":
				attr := e.Attrs().GetBytes(tString)
				if attr != nil {
					T = attr.ValueBytes()
				}
				s = sr.nextString(idx)
				sr.r.AssignNext(s)
				idx++
			case "is":
				Is = true
			case "t", "v":
			default:
				return fmt.Errorf("unexpected element: `%s` when looking for a `c`", e.Name())
			}
			xml.ReleaseStart(e)
		case *xml.EndElement:
			switch {
			case bytes.Equal(e.NameBytes(), cString):
				if s == nil {
					return fmt.Errorf("XML `%s` end element reached before `c` start element", e.Name())
				}
				switch {
				case Is, bytes.Equal(T, inlineString): // already assigned
				case bytes.Equal(T, sString):
					idx, err := strconv.Atoi(*s)
					if err != nil {
						return err
					}
					if idx < len(shared) && idx >= 0 {
						*s = shared[idx]
					} else {
						return fmt.Errorf("Got index %d. But overflows shared strings (%d)", idx, len(shared))
					}
				default:
					f, err := strconv.ParseFloat(*s, 64)
					if err == nil {
						*s = strconv.FormatFloat(f, 'f', -1, 64)
					}
				}
				Is = false
				T = nil
			case bytes.Equal(e.NameBytes(), rowString):
				xml.ReleaseEnd(e)
				break loop
			}
			xml.ReleaseEnd(e)
		}
	}
	return nil
}

// Row returns the last readed row.
func (sr *SheetReader) Row() []string {
	return sr.row
}

// Read returns the row or error
func (sr *SheetReader) Read() (record []string, err error) {
	if sr.Next() {
		record = sr.Row()
	} else {
		err = sr.Error()
		if err == nil {
			err = io.EOF
		}
	}
	return
}

// Close closes the sheet file reader.
func (sr *SheetReader) Close() error {
	return sr.rc.Close()
}
