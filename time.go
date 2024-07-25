package excelizeHelper

import (
	"time"
)

func Time2ExcelDate(value time.Time) float64 {
	// Excel 基准日期为 1900 年 1 月 1 日
	date := time.Date(1900, time.January, 1, 0, 0, 0, 0, time.UTC)
	// 计算时间差（天数）
	days := value.Sub(date) / (24 * time.Hour)
	// Excel 日期从 1 开始计数，所以需要加 2
	return float64(days) + 2
}

func ExcelDate2Time(value float64) time.Time {
	//Excel日期是从1900年1月1日开始计数的，但是由于1900年的Excel日期系统有一个错误（即1900年被错误地认为是闰年），因此实际上Excel日期是从1900年1月0日开始的，这意味着1900年1月1日在Excel中表示为1.0
	//Excel基准日期为1900年1月1日，但Excel将其视为1.0
	baseDate := time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)
	// 将Excel日期减去1以修正基准日期，然后乘以一天的纳秒数
	ns := int64((value - 1) * 24 * 60 * 60 * 1e9)
	// 返回基于基准日期和时间差的结果
	return baseDate.Add(time.Duration(ns))
}
