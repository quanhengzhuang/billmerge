package main

import (
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"
	"log/slog"
	"flag"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

 type BillItem struct {
	RowNum []int
	Date string
	Amount string
	Remark string
	IsDuplicate bool
 }

func main() {
	var mainFilePath string

	// 只设置主文件的标志
	flag.StringVar(&mainFilePath, "main", "", "主账单文件路径")
	flag.Parse()

	// 验证主文件参数
	if mainFilePath == "" {
		slog.Error("主账单文件路径不能为空")
		flag.Usage()
		os.Exit(1)
	}

	// 获取剩余的所有参数作为子文件路径
	filePaths := flag.Args()
	if len(filePaths) == 0 {
		slog.Error("至少需要指定一个子账单文件")
		flag.Usage()
		os.Exit(1)
	}

	slog.Info("解析参数成功",
		"main", mainFilePath,
		"files", filePaths,
	)

	// 配置 slog 输出到标准输出，并设置格式
	logger := slog.New(slog.NewTextHandler(os.Stdout, &slog.HandlerOptions{
		Level: slog.LevelDebug,
	}))
	slog.SetDefault(logger)
	slog.Info("start")

	bills, err := GetBillsMap(filePaths, 0, 1, 2)
	if err != nil {
		panic(fmt.Sprintf("GetBillsMap error. err: %v", err))
	}

	slog.Info("bills count", "count", len(bills))
	for k, v := range bills {
		slog.Info("bill", "key", k, "value", v.Remark)
	}

	// main
	dateLine := 1
	remarkLine := 3
	amountLine := 7
	resultLineName := "E"
	mainBillMap, err := GetBillsMap([]string{mainFilePath}, dateLine, remarkLine, amountLine)
	if err != nil {
		panic(fmt.Sprintf("GetBillsMap error. err: %v", err))
	}

	slog.Info("main bill count", "count", len(mainBillMap))
	for k, v := range mainBillMap {
		slog.Info("main bill", "key", k, "value", v.Remark)
	}

	xl, err := excelize.OpenFile(mainFilePath)
	if err != nil {
		panic(fmt.Sprintf("open main bill file failed. err: %v", err))
	}
	sheetName := xl.GetSheetName(0)

	matchedCount := 0
	for k, v := range mainBillMap {
		if _, ok := bills[k]; ok {
			remark := bills[k].Remark
			if v.IsDuplicate {
				remark = "MAIN_DUPLICATE_TODO:" + remark
			}

			for _, rowNum := range v.RowNum {
				xl.SetCellValue(sheetName, fmt.Sprintf("%s%d", resultLineName, rowNum+1), remark)
			}

			matchedCount++
			slog.Info("match success", "key", k, "value", remark)
		}
	}

	slog.Info("matched count", "count", matchedCount)

	newFilePath := fmt.Sprintf("result_%s.xlsx", time.Now().Format("20060102_150405"))
	err = xl.SaveAs(newFilePath)
	if err != nil {
		panic(fmt.Sprintf("save as failed. err: %v", err))
	}

	slog.Info("save as success", "new file path", newFilePath)
}

// GetBillItems 返回 BillItem 列表
func GetBillItems(filePaths []string, dateLine int, remarkLine int, amountLine int) ([]BillItem, error) {
	bills := make([]BillItem, 0)

	for _, filePath := range filePaths {
		rows, err := readExcel(filePath)
		if err != nil {
			return nil, fmt.Errorf("readExcel error. err: %v", err)
		}

		for i, row := range rows {
			if i == 0 {
				continue
			}

			if len(row) < dateLine + 1 || len(row) < amountLine + 1 {
				return nil, fmt.Errorf("row count is less than dateLine + 1 or amountLine + 1. dateLine: %d, amountLine: %d, filePath: %s, row: %v", dateLine, amountLine, filePath, row)
			}

			if _, err := time.Parse("2006-01-02", row[dateLine]); err != nil {
				return nil, fmt.Errorf("date format error. err: %v, filePath: %s, row: %v", err, filePath, row)
			}

			amount, err := strconv.ParseFloat(row[amountLine], 64)
			if err != nil {
				return nil, fmt.Errorf("amount format error. err: %v, filePath: %s, row: %v", err, filePath, row)
			}

			bill := BillItem{
				RowNum: []int{i},
				Date: row[dateLine],
				Amount: fmt.Sprintf("%.2f", amount),
				Remark: strings.TrimSpace(row[remarkLine]),
			}
			bills = append(bills, bill)
		}
	}

	return bills, nil
}

// GetBillsMap 返回 map[string]string，key 为日期和金额，value 为备注
func GetBillsMap(filePaths []string, dateLine int, remarkLine int, amountLine int) (map[string]BillItem, error) {
	items, err := GetBillItems(filePaths, dateLine, remarkLine, amountLine)
	if err != nil {
		return nil, fmt.Errorf("GetBillItems error. err: %v", err)
	}	

	bills := make(map[string]BillItem)
	for _, item := range items {
		key := fmt.Sprintf("%v:%v", item.Date, item.Amount)

		if _, ok := bills[key]; ok {
			item.IsDuplicate = true
			item.Remark = "DUPLICATE_TODO:" + bills[key].Remark + ":::" + item.Remark
			item.RowNum = append(bills[key].RowNum, item.RowNum...)
		} 

		bills[key] = item
	}

	return bills, nil
}

// 打开一个 excel 文件，并将第一个 sheet 中的数据，存入一个二维数组中
func readExcel(filePath string) ([][]string, error) {
	xl, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("open file failed. err: %v, filePath: %s", err, filePath)
	}

	sheetName := xl.GetSheetName(0)
	rows, err := xl.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("get rows failed. err: %v, filePath: %s", err, filePath)
	}	
	
	return rows, nil
}

