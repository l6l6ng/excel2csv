package main

import (
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"path"
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

	dirName := dir(filePath)
	if err := mkdir(dirName); err != nil {
		fmt.Println(err)
		return
	}

	sheetList := f.GetSheetList()
	for _, sheetName := range sheetList {
		rows, err := f.GetRows(sheetName)
		if err != nil {
			fmt.Println(err)
			return
		}

		if len(rows) > 0 {
			saveCsv(dirName, sheetName, rows)
		}
	}
}

func saveCsv(dirName, fileName string, data [][]string) {
	handler, _ := os.Create("./" + dirName + "/" + fileName + ".csv")

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

func dir(filePath string) string {
	filename := path.Base(filePath)
	ext := path.Ext(filename)
	return strings.Trim(filename, ext)
}

func mkdir(dirName string) error {
	dirPath := "./" + dirName
	err := os.Mkdir(dirPath, 0777)
	if err != nil && !os.IsExist(err) {
		return err
	}
	return nil
}
