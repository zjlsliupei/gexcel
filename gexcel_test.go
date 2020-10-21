package gexcel

import (
	"fmt"
	"testing"
)

var (
	excelPath string
)

func init() {
	excelPath = "D:\\tmp\\案件导入 (1).xlsx"
}

func TestImport(t *testing.T) {
	g, err := New(excelPath)
	if err != nil {
		t.Error(excelPath, " import err:", err)
		return
	}
	valid := `{
    "fieldRule": [
		{
			"title":"案由",
			"alias": "name",
			"rule": "required|myFunc",
			"message": "required:案由必须填|myFunc:我是自定义验证"
		},
		{
			"title":"当事人1",
			"alias": "dangshiren",
			"rule": "required",
			"message": "required:当事人1必须填"
		}
	],
	"range": {
		"from": 1,
		"to": 0
	} 
}
`
	g.AddCustomValidator("myFunc", func(val interface{}) bool {
		fmt.Println("hh", val)
		return false
	})
	err = g.Validate(valid, "Sheet1")
	if err != nil {
		t.Error(err)
	}
	rows := g.GetRows("Sheet1")
	fmt.Println(rows)
}
