package excelexport

import (
	"io"

	"github.com/techquest-tech/gin-shared/pkg/orm"
	"github.com/techquest-tech/gin-shared/pkg/query"
	"github.com/xuri/excelize/v2"
	"go.uber.org/zap"
	"gorm.io/gorm"
)

type ExportService struct {
	Logger  *zap.Logger
	DB      *gorm.DB
	Exports []*ExportDefine
}

type ExportDefine struct {
	Query  *query.RawQuery
	Output *ExcelExport
}

func (es *ExportService) Init() {
	for _, item := range es.Exports {
		item.Output.Logger = es.Logger
		item.Query.Sql = orm.ReplaceTablePrefix(item.Query.Sql)
	}
}

func (es *ExportService) DoExport(params map[string]interface{}, w io.Writer) error {
	f := excelize.NewFile()

	defer f.Close()
	anyData := false

	for _, item := range es.Exports {
		data, err := item.Query.Query(es.DB, params)
		if err != nil {
			es.Logger.Error("query for item failed", zap.Error(err))
			return err
		}
		es.Logger.Info("read data done.", zap.Int("len", len(data)))

		anyData = anyData || len(data) > 0

		err = item.Output.Export(f, data)
		if err != nil {
			es.Logger.Error("export to sheet failed.", zap.Error(err), zap.String("sheet", item.Output.SheetName))
			return err
		}
		es.Logger.Info("export to sheet done", zap.String("sheet", item.Output.SheetName))
	}

	sheet1 := f.GetSheetName(0)
	f.DeleteSheet(sheet1)

	err := f.Write(w)
	// err := f.SaveAs(file)
	if err != nil {
		es.Logger.Error("flush to writer failed.", zap.Error(err))
		return err
	}

	if !anyData {
		es.Logger.Warn("no data for this export job.")
	}

	es.Logger.Info("export done")

	return nil
}
