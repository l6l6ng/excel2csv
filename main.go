package main

import (
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"strings"
)

func main() {
	var filePath string
	flag.StringVar(&filePath, "f", "", "excel path")
	flag.Parse()

	//var filePath = "测试文件.xlsx"

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	sheetList := f.GetSheetList()

	for _, sheetName := range sheetList {
		rows, err := f.GetRows(sheetName)
		if err != nil {
			fmt.Println(err)
			return
		}

		if len(rows) > 0 {
			saveCsv(sheetName, rows)
		}
	}
}

func saveCsv(fileName string, data [][]string) {
	handler, _ := os.Create(fileName + ".csv")

	defer func(handler *os.File) {
		err := handler.Close()
		if err != nil {
			fmt.Println(err)
			return
		}
	}(handler)

	var list []string
	for _, v := range data {
		list = append(list, strings.Join(v, ","))
	}
	b := []byte(strings.Join(list, "\n"))
	n, err := handler.Write(b)
	if n != len(b) {
		log.Fatalln(err)
	}
}
