package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"strconv"
)

// Sheet ...
type Sheet struct {
	parent *XLSX
	zFile  *zip.File
}

// Open ...
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

// SheetReader ...
type SheetReader struct {
	s   *Sheet
	rc  io.ReadCloser
	ec  *xml.Decoder
	row []string
	err error
}

func (sr *SheetReader) Error() error {
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

// Next ...
func (sr *SheetReader) Next() bool {
	var (
		err error
		tk  xml.Token
	)
loop:
	for err == nil {
		tk, err = sr.ec.Token()
		if err != nil {
			if err != io.EOF {
				sr.err = err
			}
			break
		}

		switch stk := tk.(type) {
		case xml.StartElement:
			if stk.Name.Local != "row" {
				continue
			}
			sr.row = sr.row[:0]

			row := xlsxRow{}
			shared := sr.s.parent.sharedStrings
			sr.err = sr.ec.DecodeElement(&row, &stk)
			if sr.err == nil {
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
						sr.row = append(sr.row, c.V)
					}
				}
			}
			break loop
		case xml.EndElement:
			if stk.Name.Local == "sheetData" {
				break loop
			}
		}
	}

	return err == nil && sr.err == nil
}

// Row ...
func (sr *SheetReader) Row() []string {
	return sr.row
}

// Close ...
func (sr *SheetReader) Close() error {
	return sr.rc.Close()
}
