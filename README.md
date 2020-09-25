# gexcel
valida excel data for golang

## 安装
```go
go get github.com/zjlsliupei/gexcel
```

## 快速开始
```go
import (    
    "github.com/zjlsliupei/gexcel"
)

g, err := New("d:\\demo.xlxs")
if err != nil {
    fmt.Println(excelPath, " import err:", err)
    return
}
valid := `{
    "fieldRule": [
        {
            "title":"标题",
            "alias": "name",
            "rule": "required",
            "message": "required:标题必填"
        },
        {
            "title":"年龄",
            "alias": "age",
            "rule": "required",
            "message": "required:年龄必填"
        }
    ],
    "range": {
        "from": 1,
        "to": 2000
    } 
}
`
err = g.Validate(valid, "Sheet1")
if err != nil {
    fmt.Println(err)
}
rows := g.GetRows("Sheet1")
fmt.Println(rows)
```

## 规则说明
字段名称 | 类型| 是否必填 |说明
---|---|---|---
fieldRule |json| 是|字段验证规则
range |json| 是|excel数据范围，包含表头部分

fieldRule包说明

字段名称 | 类型| 是否必填 |说明
---|---|---|---
title |string| 是|excel表头名称
alias |string| 是|表头转换后名称
rule |string| 否|字段验证规则，验证规则参考链接，https://github.com/gookit/validate/blob/master/README.zh-CN.md#built-in-validators
message |string| 否|字段验证错误提示

range包说明

字段名称 | 类型| 是否必填 |说明
---|---|---|---
from |int| 是|数据表验证起始行，包含表头，从1开始
to |int| 否|数据表验证结束行，0：表示到底


