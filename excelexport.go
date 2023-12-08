package excelexport

import (
	"fmt"

	"github.com/xuri/excelize/v2"
	"go.uber.org/zap"
)

type Sheet struct {
	File    *excelize.File
	StyleID int
	Name    string
	Row     int
	Col     int
}

func (s *Sheet) GetAxis() string {
	if s.Col == 0 {
		s.Col = 1
	}
	strCol, err := excelize.ColumnNumberToName(s.Col)
	if err != nil {
		panic(err)
	}
	axis := fmt.Sprintf("%s%d", strCol, s.Row)
	return axis
}

func (s *Sheet) SetHeader(header *ExcelHeader) {
	if s.Col == 0 {
		s.Col = 1
	}
	strCol, err := excelize.ColumnNumberToName(s.Col)
	if err != nil {
		panic(err)
	}
	axis := fmt.Sprintf("%s%d", strCol, s.Row)

	if header.Width > 0.0 {
		s.File.SetColWidth(s.Name, strCol, strCol, header.Width)
	}
	s.File.SetCellStr(s.Name, axis, header.Title)
	if header.style > 0 {
		s.File.SetCellStyle(s.Name, axis, axis, header.style)
	}
	s.Col = s.Col + 1
}

func (s *Sheet) SetValue(value interface{}) {
	s.File.SetCellValue(s.Name, s.GetAxis(), value)
	s.Col = s.Col + 1
}

func (s *Sheet) SetString(value string) {
	s.File.SetCellStr(s.Name, s.GetAxis(), value)
	s.Col = s.Col + 1

}

func (s *Sheet) SetInt(value int, style ...int) {
	s.File.SetCellInt(s.Name, s.GetAxis(), value)
	s.Col = s.Col + 1
}

func (s *Sheet) SetFloat(value float64, prec, bitSize int) {
	s.File.SetCellFloat(s.Name, s.GetAxis(), value, prec, bitSize)
	s.Col = s.Col + 1
}

func (s *Sheet) NextRow() {
	s.Col = 1
	s.Row = s.Row + 1
}
func (s *Sheet) NextCell() {
	s.Col = s.Col + 1
}

type ExcelHeader struct {
	Title string
	Width float64
	Key   string //key for
	style int    //it's style ID
}

// func for how to get Value from the key defined.
type GetCellValue func(key string, data map[string]interface{}) interface{}

// simple return the value by key, if not existed, return the empty
var GetCellValueSimple = func(key string, data map[string]interface{}) interface{} {
	v, ok := data[key]
	if !ok {
		return key
	}
	return v
}

var defaultStyle = &excelize.Style{
	Alignment: &excelize.Alignment{
		Horizontal: "center",
		Vertical:   "center",
	},
	Font: &excelize.Font{
		Bold: true,
	},
}

type ExcelExport struct {
	SheetName string
	Index     bool //similar to pandas export setting if included Index.
	Columns   []*ExcelHeader
	Style     *excelize.Style
	Logger    *zap.Logger
	Mode      GetCellValue
}

func (ee *ExcelExport) Export(f *excelize.File, data []map[string]interface{}) error {
	if ee.Logger == nil {
		ee.Logger = zap.L()
	}

	if ee.Style == nil {
		ee.Style = defaultStyle
	}
	if ee.Mode == nil {
		ee.Mode = GetCellValueSimple
	}

	styleID, err := f.NewStyle(ee.Style)
	if err != nil {
		ee.Logger.Error("create style failed", zap.Error(err))
		return err
	}

	i, _ := f.GetSheetIndex(ee.SheetName)
	if i == -1 {
		f.NewSheet(ee.SheetName)
		ee.Logger.Info("create new sheet", zap.String("sheet", ee.SheetName))
	}

	st := &Sheet{
		File: f,
		Name: ee.SheetName,
		Row:  1,
	}

	index := &ExcelHeader{Title: "Index"}
	index.style = styleID
	if ee.Index {
		st.SetHeader(index)
	}

	for _, item := range ee.Columns {
		item.style = styleID
		st.SetHeader(item)
	}
	ee.Logger.Info("write headers done.")

	for index, row := range data {
		st.NextRow()
		if ee.Index {
			st.SetValue(index + 1)
		}

		for _, item := range ee.Columns {
			value := ee.Mode(item.Key, row)
			// st.SetString(fmt.Sprintf("%s", value))
			st.SetValue(value)
			ee.Logger.Debug("set cell value", zap.Int("x", st.Row), zap.Int("y", st.Col), zap.Any("value", value))
		}
		ee.Logger.Debug("write row done", zap.Int("row", index))
	}

	ee.Logger.Info("write data done.", zap.String("sheet", ee.SheetName), zap.Int("rows", st.Row))

	return nil
}
