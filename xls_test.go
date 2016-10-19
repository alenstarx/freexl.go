package goxls

import (
	"fmt"
	"testing"
)

func TestXLS(t *testing.T) {
	xls := XlsHandleNew("test2.xls")
	ok := xls.Open()
	if ok {
		sheet, err := xls.GetSheetCount()
		if err != nil {
			t.Error("SheetCount:", err.Error())
		}
		sst, err := xls.GetStringCount()
		if err != nil {
			t.Error("StringCount:", err.Error())
		}
		format, err := xls.GetFormatCount()
		if err != nil {
			t.Error("FormatCount:", err.Error())
		}
		xf, err := xls.GetXfCount()
		if err != nil {
			t.Error("XfCount:", err.Error())
		}
		fmt.Printf("GetInfo: %d %d %d %d \n", sheet, sst, format, xf)
		for i := uint(0); i < sheet; i++ {
			name := xls.GetWorksheetName(i)
			fmt.Printf("#%02d %s\n", i, name)
			ok := xls.GetSelectActiveWorksheet(i)
			if ok {
				active := xls.GetActiveWorksheet()
				if active < 0 {
					t.Error("GetActiveWorksheet failure")
				} else {
					rows, colums, err := xls.WorksheetDimensions()
					if err != nil {
						t.Error("WorksheetDimensions failure")
					} else {
						for r := uint(0); r < rows; r++ {
							for c := uint(0); c < colums; c++ {
								cell := xls.GetCellValue(r, c)
								if cell == nil {
									t.Error("GetCellValue failure")
								} else {
									typ := cell.GetType()
									switch typ {
									case FREEXL_CELL_INT:
										v := cell.GetInt()
										fmt.Printf("\t%d", v)
									case FREEXL_CELL_DOUBLE:
										v := cell.GetDouble()
										fmt.Printf("\t%1.12f", v)
									case FREEXL_CELL_TEXT:
										fallthrough
									case FREEXL_CELL_SST_TEXT:
										fallthrough
									case FREEXL_CELL_DATE:
										fallthrough
									case FREEXL_CELL_DATETIME:
										fallthrough
									case FREEXL_CELL_TIME:
										v := cell.GetText()
										fmt.Printf("\t'%s'", v)
									case FREEXL_CELL_NULL:
									default:
										fmt.Println("unkown type")
									}
								}
							}
						}
						fmt.Printf("\n")
					}
				}

			} else {
				t.Error("SelectActiveWorksheet failure")
			}
		}
		xls.Close()
	} else {
		fmt.Println("xls file open failed")
	}
}
