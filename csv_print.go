package gocsv

import (
	"fmt"
	"strings"
)

func CsvPrintln(data ...interface{}) {
	n := len(data)
	fmtstr := strings.TrimSuffix(strings.Repeat("\"%v\",", n), ",")
	fmt.Printf(fmtstr+"\n", data...)
}
