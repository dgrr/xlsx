package xlsx

import (
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
