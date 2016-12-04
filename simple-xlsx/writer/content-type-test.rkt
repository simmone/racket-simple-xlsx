#lang racket

(require "../xlsx.rkt")

(require rackunit/text-ui)

(require rackunit "content-type.rkt")

(define test-content-type
  (test-suite
   "test-content-type"

   (test-case
    "test-content-type"
    
    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet "Sheet1" '((1)))
      (send xlsx add-data-sheet "Sheet2" '((1)))
      (send xlsx add-data-sheet "Sheet3" '((1)))
      (send xlsx add-line-chart-sheet "Chart1" "Chart1")
      (send xlsx add-line-chart-sheet "Chart2" "Chart2")
      (send xlsx add-line-chart-sheet "Chart3" "Chart3")

      (check-equal? (write-content-type xlsx)
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/><Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/><Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/chartsheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/><Override PartName=\"/xl/charts/chart2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/><Override PartName=\"/xl/drawings/drawing2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/chartsheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/><Override PartName=\"/xl/charts/chart3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/><Override PartName=\"/xl/drawings/drawing3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/chartsheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/></Types>"))

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet "Sheet1" '(("1")))
      (send xlsx add-data-sheet "Sheet2" '((1)))
      (send xlsx add-data-sheet "Sheet3" '((1)))
      (send xlsx add-line-chart-sheet "Chart1" "Chart1")
      (send xlsx add-line-chart-sheet "Chart2" "Chart2")
      (send xlsx add-line-chart-sheet "Chart3" "Chart3")

      (check-equal? (write-content-type xlsx)
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/><Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/><Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/chartsheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/><Override PartName=\"/xl/charts/chart2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/><Override PartName=\"/xl/drawings/drawing2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/chartsheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/><Override PartName=\"/xl/charts/chart3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/><Override PartName=\"/xl/drawings/drawing3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/chartsheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/><Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/></Types>"))

    )
   ))

(run-tests test-content-type)
