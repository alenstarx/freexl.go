package main

/*
#cgo pkg-config: freexl
#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <freexl.h>

char* xls_get_cell_value(FreeXL_CellValue *cell) {
	switch (cell->type) {
		case FREEXL_CELL_INT:
        return &cell->value.int_value;
        break;
        case FREEXL_CELL_DOUBLE:
        return &cell->value.double_value;
        break;
        case FREEXL_CELL_TEXT:
        case FREEXL_CELL_SST_TEXT:
        return cell->value.text_value;
        break;
        case FREEXL_CELL_DATE:
        case FREEXL_CELL_DATETIME:
        case FREEXL_CELL_TIME:
        return cell->value.text_value;
        break;
        case FREEXL_CELL_NULL:
        default:
        break;
	}
	return 0;
}
*/
import "C"

import (
	// "runtime"
	"unsafe"
)

// Cell Type
const (
	FREEXL_CELL_NULL     = C.FREEXL_CELL_NULL
	FREEXL_CELL_INT      = C.FREEXL_CELL_INT
	FREEXL_CELL_DOUBLE   = C.FREEXL_CELL_DOUBLE
	FREEXL_CELL_TEXT     = C.FREEXL_CELL_TEXT
	FREEXL_CELL_SST_TEXT = C.FREEXL_CELL_SST_TEXT
	FREEXL_CELL_DATE     = C.FREEXL_CELL_DATE
	FREEXL_CELL_DATETIME = C.FREEXL_CELL_DATETIME
	FREEXL_CELL_TIME     = C.FREEXL_CELL_TIME
)

const (
	FREEXL_BIFF_ASCII               = C.FREEXL_BIFF_ASCII
	FREEXL_BIFF_CODEPAGE            = C.FREEXL_BIFF_CODEPAGE
	FREEXL_BIFF_CP1250              = C.FREEXL_BIFF_CP1250
	FREEXL_BIFF_CP1251              = C.FREEXL_BIFF_CP1251
	FREEXL_BIFF_CP1252              = C.FREEXL_BIFF_CP1252
	FREEXL_BIFF_CP1253              = C.FREEXL_BIFF_CP1253
	FREEXL_BIFF_CP1254              = C.FREEXL_BIFF_CP1254
	FREEXL_BIFF_CP1255              = C.FREEXL_BIFF_CP1255
	FREEXL_BIFF_CP1256              = C.FREEXL_BIFF_CP1256
	FREEXL_BIFF_CP1257              = C.FREEXL_BIFF_CP1257
	FREEXL_BIFF_CP1258              = C.FREEXL_BIFF_CP1258
	FREEXL_BIFF_CP1361              = C.FREEXL_BIFF_CP1361
	FREEXL_BIFF_CP437               = C.FREEXL_BIFF_CP437
	FREEXL_BIFF_CP720               = C.FREEXL_BIFF_CP720
	FREEXL_BIFF_CP737               = C.FREEXL_BIFF_CP737
	FREEXL_BIFF_CP775               = C.FREEXL_BIFF_CP775
	FREEXL_BIFF_CP850               = C.FREEXL_BIFF_CP850
	FREEXL_BIFF_CP852               = C.FREEXL_BIFF_CP852
	FREEXL_BIFF_CP855               = C.FREEXL_BIFF_CP855
	FREEXL_BIFF_CP857               = C.FREEXL_BIFF_CP857
	FREEXL_BIFF_CP858               = C.FREEXL_BIFF_CP858
	FREEXL_BIFF_CP860               = C.FREEXL_BIFF_CP860
	FREEXL_BIFF_CP861               = C.FREEXL_BIFF_CP861
	FREEXL_BIFF_CP862               = C.FREEXL_BIFF_CP862
	FREEXL_BIFF_CP863               = C.FREEXL_BIFF_CP863
	FREEXL_BIFF_CP864               = C.FREEXL_BIFF_CP864
	FREEXL_BIFF_CP865               = C.FREEXL_BIFF_CP865
	FREEXL_BIFF_CP866               = C.FREEXL_BIFF_CP866
	FREEXL_BIFF_CP869               = C.FREEXL_BIFF_CP869
	FREEXL_BIFF_CP874               = C.FREEXL_BIFF_CP874
	FREEXL_BIFF_CP932               = C.FREEXL_BIFF_CP932
	FREEXL_BIFF_CP936               = C.FREEXL_BIFF_CP936
	FREEXL_BIFF_CP949               = C.FREEXL_BIFF_CP949
	FREEXL_BIFF_CP950               = C.FREEXL_BIFF_CP950
	FREEXL_BIFF_DATEMODE            = C.FREEXL_BIFF_DATEMODE
	FREEXL_BIFF_DATEMODE_1900       = C.FREEXL_BIFF_DATEMODE_1900
	FREEXL_BIFF_DATEMODE_1904       = C.FREEXL_BIFF_DATEMODE_1904
	FREEXL_BIFF_FORMAT_COUNT        = C.FREEXL_BIFF_FORMAT_COUNT
	FREEXL_BIFF_ILLEGAL_SHEET_INDEX = C.FREEXL_BIFF_ILLEGAL_SHEET_INDEX
	FREEXL_BIFF_ILLEGAL_SST_INDEX   = C.FREEXL_BIFF_ILLEGAL_SST_INDEX
	FREEXL_BIFF_INVALID_BOF         = C.FREEXL_BIFF_INVALID_BOF
	FREEXL_BIFF_INVALID_SST         = C.FREEXL_BIFF_INVALID_SST
	FREEXL_BIFF_MACROMAN            = C.FREEXL_BIFF_MACROMAN
	FREEXL_BIFF_MAX_RECSIZE         = C.FREEXL_BIFF_MAX_RECSIZE
	FREEXL_BIFF_MAX_RECSZ_2080      = C.FREEXL_BIFF_MAX_RECSZ_2080
	FREEXL_BIFF_MAX_RECSZ_8224      = C.FREEXL_BIFF_MAX_RECSZ_8224
	FREEXL_BIFF_OBFUSCATED          = C.FREEXL_BIFF_OBFUSCATED
	FREEXL_BIFF_PASSWORD            = C.FREEXL_BIFF_PASSWORD
	FREEXL_BIFF_PLAIN               = C.FREEXL_BIFF_PLAIN
	FREEXL_BIFF_SHEET_COUNT         = C.FREEXL_BIFF_SHEET_COUNT
	FREEXL_BIFF_STRING_COUNT        = C.FREEXL_BIFF_STRING_COUNT
	FREEXL_BIFF_UNSELECTED_SHEET    = C.FREEXL_BIFF_UNSELECTED_SHEET
	FREEXL_BIFF_UTF16LE             = C.FREEXL_BIFF_UTF16LE
	FREEXL_BIFF_VERSION             = C.FREEXL_BIFF_VERSION
	FREEXL_BIFF_VER_2               = C.FREEXL_BIFF_VER_2
	FREEXL_BIFF_VER_3               = C.FREEXL_BIFF_VER_3
	FREEXL_BIFF_VER_4               = C.FREEXL_BIFF_VER_4
	FREEXL_BIFF_VER_5               = C.FREEXL_BIFF_VER_5
	FREEXL_BIFF_VER_8               = C.FREEXL_BIFF_VER_8
	FREEXL_BIFF_WORKBOOK_NOT_FOUND  = C.FREEXL_BIFF_WORKBOOK_NOT_FOUND
	FREEXL_BIFF_XF_COUNT            = C.FREEXL_BIFF_XF_COUNT

	FREEXL_CFBF_EMPTY_FAT_CHAIN        = C.FREEXL_CFBF_EMPTY_FAT_CHAIN
	FREEXL_CFBF_FAT_COUNT              = C.FREEXL_CFBF_FAT_COUNT
	FREEXL_CFBF_ILLEGAL_FAT_ENTRY      = C.FREEXL_CFBF_ILLEGAL_FAT_ENTRY
	FREEXL_CFBF_ILLEGAL_MINI_FAT_ENTRY = C.FREEXL_CFBF_ILLEGAL_MINI_FAT_ENTRY
	FREEXL_CFBF_INVALID_SECTOR_SIZE    = C.FREEXL_CFBF_INVALID_SECTOR_SIZE
	FREEXL_CFBF_INVALID_SIGNATURE      = C.FREEXL_CFBF_INVALID_SIGNATURE
	FREEXL_CFBF_READ_ERROR             = C.FREEXL_CFBF_READ_ERROR
	FREEXL_CFBF_SECTOR_4096            = C.FREEXL_CFBF_SECTOR_4096
	FREEXL_CFBF_SECTOR_512             = C.FREEXL_CFBF_SECTOR_512
	FREEXL_CFBF_SECTOR_SIZE            = C.FREEXL_CFBF_SECTOR_SIZE
	FREEXL_CFBF_SEEK_ERROR             = C.FREEXL_CFBF_SEEK_ERROR
	FREEXL_CFBF_VERSION                = C.FREEXL_CFBF_VERSION
	FREEXL_CFBF_VER_3                  = C.FREEXL_CFBF_VER_3
	FREEXL_CFBF_VER_4                  = C.FREEXL_CFBF_VER_4
	FREEXL_FILE_NOT_FOUND              = C.FREEXL_FILE_NOT_FOUND
	FREEXL_ILLEGAL_CELL_ROW_COL        = C.FREEXL_ILLEGAL_CELL_ROW_COL
	FREEXL_ILLEGAL_MULRK_VALUE         = C.FREEXL_ILLEGAL_MULRK_VALUE
	FREEXL_ILLEGAL_RK_VALUE            = C.FREEXL_ILLEGAL_RK_VALUE
	FREEXL_INSUFFICIENT_MEMORY         = C.FREEXL_INSUFFICIENT_MEMORY
	FREEXL_INVALID_CFBF_HEADER         = C.FREEXL_INVALID_CFBF_HEADER
	FREEXL_INVALID_CHARACTER           = C.FREEXL_INVALID_CHARACTER
	FREEXL_INVALID_HANDLE              = C.FREEXL_INVALID_HANDLE
	FREEXL_INVALID_INFO_ARG            = C.FREEXL_INVALID_INFO_ARG
	FREEXL_INVALID_MINI_STREAM         = C.FREEXL_INVALID_MINI_STREAM
	FREEXL_NULL_ARGUMENT               = C.FREEXL_NULL_ARGUMENT
	FREEXL_NULL_HANDLE                 = C.FREEXL_NULL_HANDLE
	FREEXL_OK                          = C.FREEXL_OK
	FREEXL_UNKNOWN                     = C.FREEXL_UNKNOWN
	FREEXL_UNSUPPORTED_CHARSET         = C.FREEXL_UNSUPPORTED_CHARSET
)

type XlsHandle struct {
	cptr     unsafe.Pointer
	fileName string
}

type XlsCellValue struct {
	Cell C.FreeXL_CellValue
}

func (thiz *XlsCellValue) GetType() int {
	return int(thiz.Cell._type)
}

func (thiz *XlsCellValue) GetInt() int {
	return int(*C.xls_get_cell_value(&thiz.Cell))
}

func (thiz *XlsCellValue) GetDouble() float64 {
	return float64(*C.xls_get_cell_value(&thiz.Cell))
}

func (thiz *XlsCellValue) GetText() string {
	p := C.xls_get_cell_value(&thiz.Cell)
	str := C.GoString(p)
	return str
}

func XlsHandleNew(filename string) *XlsHandle {
	return &XlsHandle{fileName: filename}
}
func (thiz *XlsHandle) Open() bool {
	p := C.CString(thiz.fileName)
	//defer C.free(p)

	rc := C.freexl_open(p, &thiz.cptr)
	if rc == C.FREEXL_OK {
		return true
	}
	return false
}

func (thiz *XlsHandle) Close() {
	C.freexl_close(thiz.cptr)
}

func (thiz *XlsHandle) OpenInfo() bool {
	p := C.CString(thiz.fileName)
	// defer C.free(p)

	rc := C.freexl_open_info(p, unsafe.Pointer(&thiz.cptr))
	if rc == C.FREEXL_OK {
		return true
	}
	return false
}

func (thiz *XlsHandle) GetFatEntry(sector_index uint) int {
	var next_sector_index C.uint
	rc := C.freexl_get_FAT_entry(thiz.cptr,
		C.uint(sector_index), &next_sector_index)
	if rc == C.FREEXL_OK {
		return int(next_sector_index)
	}
	return -1
}

func (thiz *XlsHandle) GetSstString(string_index uint16) string {
	/*
		rc := C.freexl_get_SST_string()
		if rc == C.FREEXL_OK {
			return true
		} */
	return ""
}

func (thiz *XlsHandle) GetActiveWorksheet() int {
	var index C.ushort
	rc := C.freexl_get_active_worksheet(thiz.cptr, &index)
	if rc == C.FREEXL_OK {
		return int(index)
	}
	return -1
}

func (thiz *XlsHandle) GetCellValue(row uint, column uint16) *XlsCellValue {
	cell := &XlsCellValue{}
	rc := C.freexl_get_cell_value(thiz.cptr,
		C.uint(row), C.ushort(column), &cell.Cell)
	if rc == C.FREEXL_OK {
		return cell
	}
	return nil
}

func (thiz *XlsHandle) GetInfo(what uint16) uint {
	var info C.uint
	rc := C.freexl_get_info(thiz.cptr, C.ushort(what), &info)
	if rc == C.FREEXL_OK {
		return uint(info)
	}
	return 0
}

func (thiz *XlsHandle) GetWorksheetName(sheet_index uint16) string {
	var buf *C.char
	rc := C.freexl_get_worksheet_name(thiz.cptr, C.ushort(sheet_index), &buf)
	if rc == C.FREEXL_OK {
		str := C.GoString(buf)
		return str
	}
	return ""
}

func (thiz *XlsHandle) GetSelectActiveWorksheet(sheet_index uint16) bool {
	rc := C.freexl_select_active_worksheet(thiz.cptr, C.ushort(sheet_index))
	if rc == C.FREEXL_OK {
		return true
	}
	return false
}

func (thiz *XlsHandle) WorksheetDimensions() (rows uint, colums uint16) {
	var crows C.uint
	var ccolums C.ushort
	rc := C.freexl_worksheet_dimensions(thiz.cptr, &crows, &ccolums)
	if rc == C.FREEXL_OK {
		return uint(crows), uint16(ccolums)
	}
	return 0, 0
}
