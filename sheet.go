package excelizeHelper

import (
	"github.com/xuri/excelize/v2"
	"strconv"
)

// NewSheet 生成新表
func NewSheet(excelizeFile *excelize.File, sheetName string, rows [][]any) (*excelize.File, error) {
	var (
		err           error
		newSheetIndex int
	)
	if newSheetIndex, err = excelizeFile.NewSheet(sheetName); err != nil {
		return excelizeFile, err
	} else {
		excelizeFile.SetActiveSheet(newSheetIndex)
		for rowKey, row := range rows {
			for columnKey, cellValue := range row {
				if err = excelizeFile.SetCellValue(sheetName, GetColumnName(columnKey+1)+strconv.Itoa(rowKey+1), cellValue); err != nil {
					return excelizeFile, err
				}
			}
		}
		return excelizeFile, nil
	}
}
