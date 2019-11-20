package main

import (
	"fmt"
	"os"

	"github.com/digilant/xlsx2"
)

func main() {
	ws, err := xlsx.ReadFrom(os.Args[1])
	if err != nil {
		panic(err)
	}
	for _, row := range ws.Cells {
		for i, c := range row {
			if i == 0 {
				s, err := xlsx.StringToDate(c)
				if err != nil {
					fmt.Printf("%s", c)
					continue
				}

				fmt.Printf("%s", s.Format("2006-01-02"))
			} else {
				fmt.Printf(",%s", c)
			}
		}
		fmt.Println()
	}
}
