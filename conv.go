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

// StringToInt converts the string to an int64.
func StringToInt(s string) (int64, error) {
	return strconv.ParseInt(s, 10, 64)
}

// StringToUint converts the string to an uint64.
func StringToUint(s string) (uint64, error) {
	return strconv.ParseUint(s, 10, 64)
}

// MustStringToInt returns a int64 ignoring any error.
func MustStringToInt(s string) int64 {
	n, _ := strconv.ParseInt(s, 10, 64)
	return n
}

// MustStringToUint returns a uint64 ignoring any error.
func MustStringToUint(s string) uint64 {
	n, _ := strconv.ParseUint(s, 10, 64)
	return n
}
