#lang racket

(require "../define.rkt")

(require rackunit/text-ui)

(require rackunit "content-type.rkt")

(define test-content-type
  (test-suite
   "test-content-type"

   (test-case
    "test-content-type"
    
    (let ([sheet_list
           (list
            (sheet-data "Sheet1" 0 'data '() (make-hash))
            (sheet-data "Sheet2" 0 'data '() (make-hash))
            (sheet-data "Sheet3" 0 'data '() (make-hash))
            (sheet-data "Chart1" 0 'chart '() (make-hash))
            (sheet-data "Chart2" 0 'chart '() (make-hash))
            (sheet-data "Chart3" 0 'chart '() (make-hash)))])

      (check-equal? (write-content-type sheet_list)
                    (string-append
                     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                     "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/><Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocume;nt.drawingml.chart+xml\"/><Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/chartsheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/><Override PartName=\"/xl/charts/chart2.xml\" ContentType=\"application/vnd.openxmlformats-officedocume;nt.drawingml.chart+xml\"/><Override PartName=\"/xl/drawings/drawing2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/chartsheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/><Override PartName=\"/xl/charts/chart3.xml\" ContentType=\"application/vnd.openxmlformats-officedocume;nt.drawingml.chart+xml\"/><Override PartName=\"/xl/drawings/drawing3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/chartsheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/><Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/></Types>")))
    )
   ))

(run-tests test-content-type)
