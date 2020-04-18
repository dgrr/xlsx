package xlsx

import (
	"testing"
)

func TestFloatToDate(t *testing.T) {
	f := 43889.0
	d := ToDate(f)
	if s := d.Format("2006-01-02"); s != "2020-02-28" {
		t.Fatalf("Expected %s. Got %s", "2020-02-28", s)
	}
}
