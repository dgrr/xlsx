package main

import (
	"fmt"
	"log"
	"os"

	"github.com/digilant/xlsx"
)

func main() {
	ws, err := xlsx.Open(os.Args[1])
	if err != nil {
		log.Fatalln(err)
	}
	defer ws.Close()

	for _, wb := range ws.Sheets {
		r, err := wb.Open()
		if err != nil {
			log.Fatalln(err)
		}

		for r.Next() {
			fmt.Println(r.Row())
		}
		if r.Error() != nil {
			log.Fatalln(r.Error())
		}
		r.Close()
	}
}
