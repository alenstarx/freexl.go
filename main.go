package main

import "fmt"

func main() {
	xls := XlsHandleNew("test2.xls")
	ok := xls.Open()
	if ok {
		name := xls.GetWorksheetName(0)
		fmt.Println("#0 ", name)
		xls.Close()
	} else {
		fmt.Println("xls file open failed")
	}
}
