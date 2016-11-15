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
            (sheetData "Sheet1" 0 'data 0 '() (make-hash))
            (sheetData "Sheet2" 0 'data 0 '() (make-hash))
            (sheetData "Sheet3" 0 'data 0 '() (make-hash))
            (sheetData "Chart1" 0 'chart 0 '() (make-hash))
            (sheetData "Chart2" 0 'chart 0 '() (make-hash))
            (sheetData "Chart3" 0 'chart 0 '() (make-hash)))])

      (check-equal? (sheets->content-type sheet_list)
                    "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocume;nt.drawingml.chart+xml\"/><Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/chartsheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/><Override PartName=\"/xl/charts/chart2.xml\" ContentType=\"application/vnd.openxmlformats-officedocume;nt.drawingml.chart+xml\"/><Override PartName=\"/xl/drawings/drawing2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/chartsheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/><Override PartName=\"/xl/charts/chart3.xml\" ContentType=\"application/vnd.openxmlformats-officedocume;nt.drawingml.chart+xml\"/><Override PartName=\"/xl/drawings/drawing3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/chartsheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/>")
      
      (check-equal? (sheetData-seq (list-ref sheet_list 0)) 1)
      (check-equal? (sheetData-seq (list-ref sheet_list 1)) 2)
      (check-equal? (sheetData-seq (list-ref sheet_list 2)) 3)
      (check-equal? (sheetData-seq (list-ref sheet_list 3)) 4)
      (check-equal? (sheetData-seq (list-ref sheet_list 4)) 5)
      (check-equal? (sheetData-seq (list-ref sheet_list 5)) 6)

      (check-equal? (sheetData-typeSeq (list-ref sheet_list 0)) 1)
      (check-equal? (sheetData-typeSeq (list-ref sheet_list 1)) 2)
      (check-equal? (sheetData-typeSeq (list-ref sheet_list 2)) 3)
      (check-equal? (sheetData-typeSeq (list-ref sheet_list 3)) 1)
      (check-equal? (sheetData-typeSeq (list-ref sheet_list 4)) 2)
      (check-equal? (sheetData-typeSeq (list-ref sheet_list 5)) 3)
      )
    )
   ))

(run-tests test-content-type)
