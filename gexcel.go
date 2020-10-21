package gexcel

import (
	"errors"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/gookit/validate"
	"github.com/tidwall/gjson"
	"strconv"
	"strings"
)

type GExceler interface {
	Validate(validRule string, sheetName string) error
	GetRows(sheetName string) []map[string]string
	AddCustomValidator(funcName string, cb func(val interface{}) bool) error
}

type GExcel struct {
	fileHandle *excelize.File
	excelPath  string
	header     map[string][]string
	// 原始数据
	rawRows map[string][][]string
	// 保存过滤后的数据
	filterRows map[string][]map[string]string
	// 自定义验证器
	customValidator map[string]func(val interface{}) bool
}

func (e *GExcel) init(excelPath string) error {
	var err error
	e.excelPath = excelPath
	e.fileHandle, err = excelize.OpenFile(excelPath)
	if err != nil {
		return err
	}
	e.rawRows = make(map[string][][]string)
	e.header = make(map[string][]string)
	e.filterRows = make(map[string][]map[string]string)
	e.customValidator = make(map[string]func(val interface{}) bool)
	return nil
}

// readRows 获取sheet原始数据
func (e *GExcel) readRows(sheetName string, from int64, to int64) (err error) {
	// 不会重复读取
	if _, ok := e.rawRows[sheetName]; ok {
		return
	}
	var rows [][]string
	rows, err = e.fileHandle.GetRows(sheetName)
	if err != nil {
		return
	}
	rowNum := len(rows)
	if rowNum < 1 {
		return errors.New("excel数据不允许为空")
	}
	e.header[sheetName] = rows[(from - 1):from][0]
	if to == 0 {
		e.rawRows[sheetName] = rows[from:]
	} else {
		if int64(rowNum) > to {
			e.rawRows[sheetName] = rows[from:to]
		} else {
			e.rawRows[sheetName] = rows[from:rowNum]
		}
	}
	return
}

// 验证
// 格式
// {
//    "fieldRule": [
//		{
//			"title":"名称",
//			"alias": "name",
//			"rule": "required|int|min:1|max:99",
//			"message": "int:必须是数字类型| min: 最小值为1"
//		}
//	],
//	"range": { // 包含表头，如果不传默认从第一行开始
//		"from": 1,
//		"to": 1000
//	}
//}
func (e *GExcel) Validate(validRule string, sheetName string) error {
	// 验证规则是否合法
	checkErr := e.checkRule(validRule)
	if checkErr != nil {
		return checkErr
	}
	// 解析范围规则
	from := gjson.Get(validRule, "range.from").Int()
	to := gjson.Get(validRule, "range.to").Int()
	err := e.readRows(sheetName, from, to)
	if err != nil {
		return err
	}
	fieldRuleArr := gjson.Get(validRule, "fieldRule").Array()
	var fieldRules = make(map[string]map[string]string)
	for _, v := range fieldRuleArr {
		rule := make(map[string]string)
		for k, v2 := range v.Map() {
			rule[k] = v2.String()
		}
		fieldRules[v.Get("title").String()] = rule
	}
	//fmt.Println(fieldRules, e.header)
	var filterRows = []map[string]string{}
	var validErr error
I:
	for rowNum, row := range e.rawRows[sheetName] {
		_filterRow := make(map[string]string)
		for columnNum, cell := range row {
			if v, ok := fieldRules[e.header[sheetName][columnNum]]; ok {
				//fmt.Println(v, rowNum, columnNum, cell)
				realRowNum := int(from) + rowNum
				validErr = e.valid(cell, v, realRowNum, columnNum)
				if validErr != nil {
					break I
				}
				_filterRow[v["alias"]] = cell
			}
		}
		filterRows = append(filterRows, _filterRow)
	}
	e.filterRows[sheetName] = filterRows
	return validErr
}

// checkRule 检测规则
func (e *GExcel) checkRule(validRule string) error {
	// 验证外层数据
	gj := gjson.Parse(validRule)
	m := map[string]interface{}{
		"fieldRule": gj.Get("fieldRule").String(),
		"range":     gj.Get("range").String(),
		"from":      gj.Get("range.from").Int(),
		"to":        gj.Get("range.to").Int(),
	}
	v := validate.Map(m)
	v.StringRule("fieldRule", "required|json")
	v.StringRule("range", "required|json")
	v.StringRule("from", "required|int|gt:0")
	v.StringRule("to", "int|gte:0")
	if !v.Validate() {
		return errors.New(v.Errors.One())
	}
	fields := gj.Get("fieldRule").Array()
	for _, v := range fields {
		m1 := map[string]interface{}{
			"title":   v.Get("title").String(),
			"alias":   v.Get("alias").String(),
			"rule":    v.Get("rule").String(),
			"message": v.Get("message").String(),
		}
		v := validate.Map(m1)
		v.StringRule("title", "required")
		v.StringRule("alias", "required")
		if !v.Validate() {
			return errors.New(v.Errors.One())
		}
	}
	return nil
}

// valid
func (e *GExcel) valid(cell string, validRule map[string]string, rowNum int, columnNum int) error {
	m := map[string]interface{}{
		validRule["title"]: cell,
	}
	v := validate.Map(m)
	// 如果不存在rule则不验证
	if _, ok := validRule["rule"]; !ok {
		return nil
	}
	// 也可以这样，一次添加多个验证器
	v.StringRule(validRule["title"], validRule["rule"])
	message := make(map[string]string)
	_message := strings.Split(validRule["message"], "|")
	for _, v := range _message {
		_v := strings.Split(v, ":")
		message[strings.TrimSpace(_v[0])] = strings.TrimSpace(_v[1])
	}
	v.AddMessages(message)
	// 判断自定义验证器
	if len(e.customValidator) > 0 {
		for funcName, cb := range e.customValidator {
			v.AddValidator(funcName, cb)
		}
	}
	if !v.Validate() {
		errMsg := "第" + strconv.FormatInt(int64(rowNum+1), 10) +
			"行,第" + strconv.FormatInt(int64(columnNum+1), 10) + "列:" + v.Errors.One()
		return errors.New(errMsg)
	}
	return nil
}

// GetRows 获取行数据
func (e *GExcel) GetRows(sheetName string) []map[string]string {
	if v, ok := e.filterRows[sheetName]; ok {
		return v
	}
	return nil
}

// AddCustomValidator 添加自定义验证器
func (e *GExcel) AddCustomValidator(funcName string, cb func(val interface{}) bool) error {
	e.customValidator[funcName] = cb
	return nil
}

func New(excelPath string) (GExceler, error) {
	ge := GExcel{}
	err := ge.init(excelPath)
	return &ge, err
}
