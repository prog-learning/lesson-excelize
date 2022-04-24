package main

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

/* 1.ファイルを作成 */

// func main() {
// 	for i := 0; i < 9; i++ {
// 		/* ファイルの作成と入力 */
// 		f := excelize.NewFile()
// 		id := i + 1
// 		f.SetCellValue("Sheet1", "A1", "ID")
// 		f.SetCellValue("Sheet1", "B1", "NAME")
// 		f.SetCellValue("Sheet1", "C1", "VALUE")
// 		f.SetCellValue("Sheet1", "A2", id)
// 		f.SetCellValue("Sheet1", "B2", fmt.Sprintf("佐藤さん%v", i))
// 		f.SetCellValue("Sheet1", "C2", id*100)
// 		// 指定されたパスに従ってファイルを保存します
// 		if err := f.SaveAs(fmt.Sprintf("data/book%v.xlsx", id)); err != nil {
// 			fmt.Println(err)
// 		}
// 	}
// }

/* 2.作成したファイルの書き換え */
func main() {
	/* 新規ファイルの作成 */
	newFile := excelize.NewFile()
	newFile.SetCellValue("Sheet1", "A1", "ID")
	newFile.SetCellValue("Sheet1", "B1", "NAME")
	newFile.SetCellValue("Sheet1", "C1", "VALUE")

	/* 参照ファイルの読み込み */
	root := "data"
	files, err := listFiles(root)
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	fmt.Println(files)

	for i, file := range files {
		fmt.Println(i)
		f, err := excelize.OpenFile(file)
		if err != nil {
			fmt.Println(err)
			return
		}
		defer func() {
			if err := f.Close(); err != nil {
				fmt.Println(err)
			}
		}()

		rows, err := f.GetRows("Sheet1")
		if err != nil {
			fmt.Println(err)
			return
		}
		fmt.Println(rows[1])
		fmt.Println(fmt.Sprintf("A%v", i+2))

		err = newFile.SetSheetRow("Sheet1", fmt.Sprintf("A%v", i+2), &rows[1])
		if err != nil {
			fmt.Println("newFile.SetSheetRow:", err)
			return
		}
	}
	if err := newFile.SaveAs("newdata/newbook.xlsx"); err != nil {
		fmt.Println(err)
		return
	}
}

func listFiles(root string) ([]string, error) {
	var files []string
	// root以下を走査
	err := filepath.Walk(root, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		// ディレクトリは除く
		if !info.IsDir() {
			files = append(files, path)
		}
		return nil
	})
	if err != nil {
		return nil, err
	}
	return files, nil
}

/* 書き換えテスト */
// func main() {
// 	/* ファイルを開く */
// 	f, err := excelize.OpenFile("test-data.xlsx")
// 	if err != nil {
// 		fmt.Println(err)
// 		return
// 	}
// 	defer func() {
// 		if err := f.Close(); err != nil {
// 			fmt.Println(err)
// 		}
// 	}()

// 	/* ファイルを書き換える */
// 	f.SetCellValue("Sheet1", "A1", "Hello, world!")
// 	f.SetSheetRow("Sheet1", "B2", &[]interface{}{"配列", "で", "横に", "書き込んでくれる", nil, 123})

// 	/* ファイルを保存する */
// 	if err := f.Save(); err != nil {
// 		fmt.Println(err)
// 		return
// 	}

// 	// ワークシート内の指定されたセルの値を取得します
// 	cell, err := f.GetCellValue("Sheet1", "A1")
// 	if err != nil {
// 		fmt.Println(err)
// 		return
// 	}
// 	fmt.Println("B2:", cell)

// 	// Sheet1 のすべてのセルを取得
// 	rows, err := f.GetRows("Sheet1")
// 	if err != nil {
// 		fmt.Println(err)
// 		return
// 	}
// 	fmt.Println(rows)
// 	for _, row := range rows {
// 		for _, colCell := range row {
// 			fmt.Print(colCell, "\t")
// 		}
// 		fmt.Println()
// 	}
// }
