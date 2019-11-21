package xlsx

import (
	"strconv"
	"time"
)

// ToDate converts the date from excel format to time.Time
func ToDate(n float64) time.Time {
	return time.Unix(int64((n-25569)*86400), 0)
}

// StringToDate converts the string from string to float and to time.Time
func StringToDate(s string) (t time.Time, err error) {
	n, err := strconv.ParseFloat(s, 64)
	if err == nil {
		t = ToDate(n)
	}
	return t, err
}

// StringToInt ...
func StringToInt(s string) (int64, error) {
	return strconv.ParseInt(s, 10, 64)
}

// StringToUint ...
func StringToUint(s string) (uint64, error) {
	return strconv.ParseUint(s, 10, 64)
}

// MustStringToInt ...
func MustStringToInt(s string) int64 {
	n, _ := strconv.ParseInt(s, 10, 64)
	return n
}

// MustStringToUint ...
func MustStringToUint(s string) uint64 {
	n, _ := strconv.ParseUint(s, 10, 64)
	return n
}
