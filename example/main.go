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
