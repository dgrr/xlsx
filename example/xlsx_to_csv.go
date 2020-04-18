// +build ignore
package main

import (
	"encoding/csv"
	"log"
	"os"
	"strings"

	"github.com/dgrr/xlsx"
)

func main() {
	if len(os.Args) < 2 {
		log.Fatalf("%s <xlsx file>\n", os.Args[0])
	}

	ws, err := xlsx.Open(os.Args[1])
	if err != nil {
		log.Fatalln(err)
	}
	defer ws.Close()

	dest := strings.Replace(os.Args[1], ".xlsx", ".csv", -1)

	wr, err := os.Create(dest)
	if err != nil {
		log.Fatalln(err)
	}
	defer wr.Close()

	cw := csv.NewWriter(wr)
	defer cw.Flush()

	wb := ws.Sheets[0]
	r, err := wb.Open()
	if err != nil {
		log.Fatalln(err)
	}

	for r.Next() {
		cw.Write(r.Row())
	}
	if r.Error() != nil {
		log.Fatalln(r.Error())
	}
	r.Close()
}
