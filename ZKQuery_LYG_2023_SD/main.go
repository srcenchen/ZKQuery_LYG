package main

import (
	"fmt"
	"github.com/go-resty/resty/v2"
	"github.com/xuri/excelize/v2"
	"net/url"
	"os"
	"strconv"
)

type Student struct {
	name  string
	class string
	jdNo  string
}

type Response struct {
	Code int    `json:"code"`
	Msg  string `json:"msg"`
	Data Data   `json:"data"`
}
type Data struct {
	Ksh  string `json:"KSH"`
	Sw   string `json:"SW"`
	Dl   string `json:"DL"`
	Jdh  string `json:"JDH"`
	Xm   string `json:"XM"`
	Xczf string `json:"XCZF"`
}

var done = make(chan bool)

func main() {
	r := resty.New()
	go func() {
		for true {
			continue
		}
	}() // 防止死锁
	// 环境检查
	if !fileIsExists("infoSD.xlsx") {
		fmt.Println("infoSD.xlsx文件不存在，此文件包含了查询的必要信息")
		<-done
	}

	if fileIsExists("./result.xlsx") {
		fmt.Println("result.xlsx 文件已经存在，说明已经运行过此程序，请及时将此文件移动到其他地方，以免数据被覆盖！")
		// <-done
	}
	initResult()

	studentList := readStudentInfo()
	for _, item := range studentList {
		item.getGrade(r)
	}
	println("Done!")
	<-done
}

func initResult() {
	// 创建相关文件
	f := excelize.NewFile()
	// create sheet
	index, _ := f.NewSheet("Sheet1")
	// 设置表头
	title := []string{
		"姓名", "建档号", "班级", "生地总分", "生物", "地理",
	}
	for index, item := range title {
		as := index + 65
		r := rune(as)
		_ = f.SetCellValue("Sheet1", string(r)+"1", item)
	}
	// 设置默认
	f.SetActiveSheet(index)
	_ = f.SaveAs("./result.xlsx")
	_ = f.Close()

	// 错误信息
	f = excelize.NewFile()
	// create sheet
	index, _ = f.NewSheet("Sheet1")
	f.SetActiveSheet(index)
	_ = f.SaveAs("./error.xlsx")
	_ = f.Close()
}

// checkFileIsExists
func fileIsExists(filename string) bool {
	_, err := os.Stat(filename)
	if os.IsNotExist(err) {
		return false
	}
	return true
}

// 读取xlsx
func readStudentInfo() (studentList []Student) {
	f, _ := excelize.OpenFile("./infoSD.xlsx")
	sheetName := f.GetSheetName(0)
	rows, _ := f.GetRows(sheetName)
	for index := range rows {
		if index == 0 {
			continue
		}
		class, _ := f.GetCellValue(sheetName, "A"+strconv.Itoa(index+1))
		name, _ := f.GetCellValue(sheetName, "B"+strconv.Itoa(index+1))
		jdNo, _ := f.GetCellValue(sheetName, "C"+strconv.Itoa(index+1))
		studentList = append(studentList, Student{
			name:  name,
			class: class,
			jdNo:  jdNo,
		})
	}
	return
}

// getGrade 获取成绩
func (s Student) getGrade(r *resty.Client) {
	print("正在查询：" + s.name + " " + s.jdNo + " " + s.class + " ")
	resp := &Response{}
	queryURL := "http://www.lygzsks.cn/grade/getGrade?param=" + s.jdNo + "__" + url.PathEscape(s.name) + "&password=0&captcha=6785"
	_, _ = r.R().SetResult(resp).Get(queryURL)
	// 判断Code
	switch resp.Code {
	case 0: // 获取成绩成功 保存
		resp.saveGrade(s)
		println("总分：" + resp.Data.Xczf + " 生物" + resp.Data.Sw + " 地理" + resp.Data.Dl)
	default: // 其他错误 要做登记
		s.error(resp)
		println("错误响应:" + resp.Msg)
	}
}

func (resp Response) saveGrade(stu Student) {
	resultFile, _ := excelize.OpenFile("./result.xlsx")
	defer func() {
		_ = resultFile.Close()
	}()
	// 添加数据
	// 获取row数据
	rows, _ := resultFile.GetRows("Sheet1")
	i := strconv.Itoa(len(rows) + 1)
	_ = resultFile.SetCellValue("Sheet1", "A"+i, stu.name)       // 姓名
	_ = resultFile.SetCellValue("Sheet1", "B"+i, stu.jdNo)       // 建档号
	_ = resultFile.SetCellValue("Sheet1", "C"+i, stu.class)      // 班级
	_ = resultFile.SetCellValue("Sheet1", "D"+i, resp.Data.Xczf) // 总分
	_ = resultFile.SetCellValue("Sheet1", "E"+i, resp.Data.Sw)   // 生物
	_ = resultFile.SetCellValue("Sheet1", "F"+i, resp.Data.Dl)   // 地理
	_ = resultFile.Save()
}

func (s Student) error(resp *Response) {
	errorFile, _ := excelize.OpenFile("./error.xlsx")
	defer func() {
		_ = errorFile.Close()
	}()
	rows, _ := errorFile.GetRows("Sheet1")
	i := strconv.Itoa(len(rows) + 1)
	_ = errorFile.SetCellValue("Sheet1", "A"+i, s.name)   // 姓名
	_ = errorFile.SetCellValue("Sheet1", "B"+i, resp.Msg) // 错误原因
	_ = errorFile.SetCellValue("Sheet1", "C"+i, s.class)  // 班级
	_ = errorFile.SetCellValue("Sheet1", "D"+i, s.jdNo)   // 建档号
	_ = errorFile.Save()
}
