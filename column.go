package excelizeHelper

// GetColumnName 获取列号
func GetColumnName(columnKey int) string {
	var mod int
	columnName := ""
	for {
		if columnKey > 0 {
			mod = columnKey % 26
			if mod == 0 {
				mod = 26
			}
			columnName = string(rune(mod+64)) + columnName
			columnKey = (columnKey - mod) / 26
		} else {
			break
		}
	}
	return columnName
}
