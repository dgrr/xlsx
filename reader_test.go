package xlsx

import (
	"archive/zip"
	"os"
	"strings"
	"testing"
)

func TestParseShared(t *testing.T) {
	sts := []string{
		"A", "B", "C", "D",
	}
	const sharedStringsText = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>A</t></si><si><t>B</t></si><si><t>C</t></si><si><t>D</t></si></sst>`
	r := strings.NewReader(sharedStringsText)
	ss, err := parseShared(r)
	if err != nil {
		t.Fatalf("Unexpected err: %q", err)
	}

	for i := range ss {
		if ss[i] != sts[i] {
			t.Fatalf("Unexpected: %s<>%s", ss[i], sts[i])
		}
	}
}

func TestParseSharedWithEmpty(t *testing.T) {
	sts := []string{
		"A", "B", "", "C", "", "D",
	}
	const sharedStringsText = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>A</t></si><si><t>B</t></si><si><t/></si><si><t>C</t></si><si><t/></si><si><t>D</t></si></sst>`
	r := strings.NewReader(sharedStringsText)
	ss, err := parseShared(r)
	if err != nil {
		t.Fatalf("Unexpected err: %q", err)
	}

	for i := range ss {
		if ss[i] != sts[i] {
			t.Fatalf("Unexpected: %s<>%s", ss[i], sts[i])
		}
	}
}

const xlsxFile = "test/spreadsheet.xlsx"

func TestParseContentType(t *testing.T) {
	file, err := os.Open(xlsxFile)
	if err != nil {
		t.Fatal(err)
	}
	defer file.Close()

	st, err := file.Stat()
	if err != nil {
		t.Fatal(err)
	}

	zr, err := zip.NewReader(file, st.Size())
	if err != nil {
		t.Fatal(err)
	}

	for _, zFile := range zr.File {
		// read where the worksheets are
		if zFile.Name == "[Content_Types].xml" {
			index, err := parseContentType(zFile)
			if err != nil {
				t.Fatal(err)
			}

			if len(index.files) != 1 {
				t.Fatalf("Unexpected len: %d. Expected 1", len(index.files))
			}
			if index.files[0] != "/xl/worksheets/sheet1.xml" {
				t.Fatalf("Unexpected spreadsheet file: %s", index.files[0])
			}

			if index.sharedStr != "/xl/sharedStrings.xml" {
				t.Fatalf("Unexpected sharedStrings file: %s", index.sharedStr)
			}

			break
		}
	}
}

func TestReadShared(t *testing.T) {
	file, err := os.Open(xlsxFile)
	if err != nil {
		t.Fatal(err)
	}
	defer file.Close()

	st, err := file.Stat()
	if err != nil {
		t.Fatal(err)
	}

	zr, err := zip.NewReader(file, st.Size())
	if err != nil {
		t.Fatal(err)
	}

	for _, zFile := range zr.File {
		// read where the worksheets are
		if zFile.Name == "[Content_Types].xml" {
			index, err := parseContentType(zFile)
			if err != nil {
				t.Fatal(err)
			}

			if len(index.files) != 1 {
				t.Fatalf("Unexpected len: %d. Expected 1", len(index.files))
			}
			if index.files[0] != "/xl/worksheets/sheet1.xml" {
				t.Fatalf("Unexpected spreadsheet file: %s", index.files[0])
			}

			if index.sharedStr != "/xl/sharedStrings.xml" {
				t.Fatalf("Unexpected sharedStrings file: %s", index.sharedStr)
			}

			shared, err := readShared(zr, index.sharedStr)
			if err != nil {
				t.Fatal(err)
			}

			expectedShared := []string{
				"Date", "A", "B", "C", "D",
			}
			for i := range shared {
				if shared[i] != expectedShared[i] {
					t.Fatalf("%s <> %s", shared[i], expectedShared[i])
				}
			}

			break
		}
	}
}
