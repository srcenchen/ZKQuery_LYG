package main

import (
	"fmt"
	"github.com/go-resty/resty/v2"
	"github.com/xuri/excelize/v2"
	"os"
	"strconv"
)

type Student struct {
	name     string
	class    string
	bmNo     string
	jdNo     string
	password string
}

type Response struct {
	Code int    `json:"code"`
	Msg  string `json:"msg"`
	Data Data   `json:"data"`
}
type Data struct {
	Zz         string `json:"zz"`
	Ty         string `json:"ty"`
	Ls         string `json:"ls"`
	Dl         string `json:"dl"`
	Yybs       string `json:"yybs"`
	Ds         string `json:"ds"`
	Hx         string `json:"hx"`
	Wh         string `json:"wh"`
	Kytl       string `json:"kytl"`
	Wl         string `json:"wl"`
	Jfzf       string `json:"jfzf"`
	ID         string `json:"id"`
	Yw         string `json:"yw"`
	Yy         string `json:"yy"`
	Sw         string `json:"sw"`
	Sx         string `json:"sx"`
	SchoolName string `json:"school name"`
	Zf         string `json:"zf"`
	Ky         string `json:"ky"`
	BizCode2   string `json:"biz_code2"`
	BizCode1   string `json:"biz code1"`
	BizCode    string `json:"biz_code"`
	Tl         string `json:"tl"`
	Name       string `json:"name"`
	Zs         string `json:"zs"`
	Jf         string `json:"jf"`
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
	if !fileIsExists("info.xlsx") {
		fmt.Println("info.xlsx文件不存在，此文件包含了查询的必要信息")
		<-done
	}

	if fileIsExists("./result.xlsx") {
		fmt.Println("result.xlsx 文件已经存在，说明已经运行过此程序，请及时将此文件移动到其他地方，以免数据被覆盖！")
		<-done
	}
	initResult()

	studentList := readStudentInfo()
	for _, item := range studentList {
		item.recognizeCaptcha(r)
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
		"姓名", "建档号", "班级", "总分", "语文", "数学", "英语", "物理", "化学", "道法", "历史", "生物", "地理", "体育",
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
	f, _ := excelize.OpenFile("./info.xlsx")
	sheetName := f.GetSheetName(0)
	rows, _ := f.GetRows(sheetName)
	for index := range rows {
		if index == 0 {
			continue
		}
		class, _ := f.GetCellValue(sheetName, "A"+strconv.Itoa(index+1))
		name, _ := f.GetCellValue(sheetName, "B"+strconv.Itoa(index+1))
		bmNo, _ := f.GetCellValue(sheetName, "C"+strconv.Itoa(index+1))
		jdNo, _ := f.GetCellValue(sheetName, "D"+strconv.Itoa(index+1))
		password, _ := f.GetCellValue(sheetName, "E"+strconv.Itoa(index+1))
		studentList = append(studentList, Student{
			name:     name,
			class:    class,
			bmNo:     bmNo,
			jdNo:     jdNo,
			password: password,
		})
	}
	return
}

// recognizeCaptcha 识别验证码
func (s Student) recognizeCaptcha(r *resty.Client) {
	print("正在查询" + s.name + " ")
	captchaURL := "http://121.229.42.255:8996/grade/getCaptcha"
	_, err := r.R().SetOutput("captcha.jpg").Get(captchaURL)
	if err != nil {
		panic(err)
	}
	// 识别验证码
	ocrURL := "http://sansg.luckysan.top:9898/ocr/file"
	resp, err := r.R().SetFile("image", "captcha.jpg").Post(ocrURL)
	if err != nil {
		panic(err)
	}
	print("验证码" + checkCaptcha(resp.String()))
	s.getGrade(checkCaptcha(resp.String()), r)
}

// checkCaptcha 检查验证码格式
func checkCaptcha(captcha string) (cap string) {
	cap = captcha
	if len(captcha) > 5 {
		// 删除前面字符，直到只剩下5位
		cap = captcha[len(captcha)-5:]
	}
	return
}

// getGrade 获取成绩
func (s Student) getGrade(captcha string, r *resty.Client) {
	resp := &Response{}
	queryURL := "http://121.229.42.255:8996/grade/getGrade?param=" + s.jdNo + "_" + s.bmNo +
		"_&password=" + s.password + "&captcha=" + captcha
	_, _ = r.R().SetResult(resp).Get(queryURL)
	// 判断Code
	switch resp.Code {
	case 0: // 获取成绩成功 保存
		resp.saveGrade(s)
		println("总分：" + resp.Data.Jfzf)
	case 1001: // 验证码错误
		println(" 验证码验证失败，正在重试")
		s.recognizeCaptcha(r)
		return
	default: // 其他错误 要做登记
		s.error(resp)
		println(" 错误响应:" + resp.Msg)
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
	_ = resultFile.SetCellValue("Sheet1", "D"+i, resp.Data.Jfzf) // 总分
	_ = resultFile.SetCellValue("Sheet1", "E"+i, resp.Data.Yw)   // 语文
	_ = resultFile.SetCellValue("Sheet1", "F"+i, resp.Data.Sx)   // 数学
	_ = resultFile.SetCellValue("Sheet1", "G"+i, resp.Data.Yy)   // 英语
	_ = resultFile.SetCellValue("Sheet1", "H"+i, resp.Data.Wl)   // 物理
	_ = resultFile.SetCellValue("Sheet1", "I"+i, resp.Data.Hx)   // 化学
	_ = resultFile.SetCellValue("Sheet1", "J"+i, resp.Data.Zz)   // 道法
	_ = resultFile.SetCellValue("Sheet1", "K"+i, resp.Data.Ls)   // 历史
	_ = resultFile.SetCellValue("Sheet1", "L"+i, resp.Data.Sw)   // 生物
	_ = resultFile.SetCellValue("Sheet1", "M"+i, resp.Data.Dl)   // 地理
	_ = resultFile.SetCellValue("Sheet1", "N"+i, resp.Data.Ty)   // 体育
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
