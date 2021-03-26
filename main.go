package main

// フォルダ内の画像を縦に並べたExcelを作成する

import (
	"fmt"
	"io/ioutil"
	"path/filepath"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"

	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
)

func main() {
	rootDir := "evidence" // 探索するフォルダ
	excelFile := excelize.NewFile()

	files, err := ioutil.ReadDir(rootDir)
	if err != nil {
		panic(err)
	}

	for _, file := range files {
		filePath := filepath.Join(rootDir, file.Name())
		if file.IsDir() {
			procDir(excelFile, filePath, "")
		} else {
			fmt.Println("Ignored:", filePath)
		}
	}

	excelFile.DeleteSheet("Sheet1")
	if err := excelFile.SaveAs("evidence.xlsx"); err != nil {
		panic(err)
	}
}

func procDir(excelFile *excelize.File, dirPath string, parentSheetName string) {
	fmt.Println("procDir. dirPath:", dirPath, ", parentSheetName:", parentSheetName)
	dirName := filepath.Base(dirPath)
	var sheetName string
	if len(parentSheetName) > 0 {
		sheetName = parentSheetName + "_" + dirName
	} else {
		sheetName = dirName
	}
	files, err := ioutil.ReadDir(dirPath)
	if err != nil {
		panic(err)
	}
	for index, file := range files {
		filePath := filepath.Join(dirPath, file.Name())
		fmt.Println("filePath:", filePath, "index:", index)
		if file.IsDir() {
			procDir(excelFile, filePath, sheetName)
		} else if strings.HasSuffix(file.Name(), ".png") {
			isSheetExists := excelFile.Sheet[sheetName] != nil
			if !isSheetExists {
				excelFile.NewSheet(sheetName)
			}
			row := index + 1
			col := "A"
			cell := fmt.Sprintf("%s%d", col, row)
			fmt.Println("cell:", cell)
			excelFile.SetRowHeight(sheetName, row, 400)
			excelFile.SetColWidth(sheetName, col, col, 40)
			if err := excelFile.AddPicture(sheetName, cell, filePath, `{
				"x_offset": 10,
				"y_offset": 10,
				"x_scale": 0.35,
				"y_scale": 0.35,
				"print_obj": true,
				"lock_aspect_ratio": true
				}`); err != nil {
				fmt.Println(err)
			}
		} else {
			fmt.Println("Ignored:", filePath)
		}
	}
}
