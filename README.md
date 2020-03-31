# xlsx

[![Go Report Card](https://goreportcard.com/badge/github.com/dgrr/xlsx)](https://goreportcard.com/report/github.com/dgrr/xlsx)
[![Build Status](https://travis-ci.com/dgrr/xlsx.svg?branch=master)](https://travis-ci.com/dgrr/xlsx)
[![Coverage Status](https://coveralls.io/repos/github/dgrr/xlsx/badge.svg?branch=master)](https://coveralls.io/github/dgrr/xlsx?branch=master)

Working with XLSX is most of the times a pain (is built with XML). This package aims to work with XLSX files to extract only the data inside. It doesn't manage styles or any other fancy feature. It supports shared strings (because it's not a fancy feature) and it is fast and easy to use.

```go
package main

import (
	"fmt"
	"log"
	"os"

	"github.com/dgrr/xlsx"
)

func main() {
	// open the XLSX file.
	ws, err := xlsx.Open(os.Args[1])
	if err != nil {
		log.Fatalln(err)
	}
	defer ws.Close() // do not forget to close

	// iterate over the sheets
	for _, wb := range ws.Sheets {
		r, err := wb.Open()
		if err != nil {
			log.Fatalln(err)
		}

		for r.Next() { // get next row
			fmt.Println(r.Row())
		}
		if r.Error() != nil { // error checking
			log.Fatalln(r.Error())
		}
		// don't forget to close the sheet!!
		r.Close()
	}
}
```
