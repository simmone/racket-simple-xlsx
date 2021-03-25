#lang racket

(require rackunit/text-ui)

(require rackunit "../../../reader.rkt")

(require "../../../xlsx/xlsx.rkt")

(require racket/runtime-path)
(define-runtime-path test1_file "test1.xlsx")

(define test-reader
  (test-suite
   "test-reader"
   
   (test-case
    "test-reader"

    (with-input-from-xlsx-file
     test1_file
     (lambda (xlsx)
      (check-equal? (XLSX-sheet_count xlsx) 10)
      
      (check-equal? (hash-count (XLSX-sheet_index_id_map xlsx)) 10)
      (check-equal? (hash-ref (XLSX-sheet_index_id_map xlsx) 0) "1")
      (check-equal? (hash-ref (XLSX-sheet_index_id_map xlsx) 9) "10")

      (check-equal? (hash-count (XLSX-sheet_index_name_map xlsx)) 10)
      (check-equal? (hash-ref (XLSX-sheet_index_name_map xlsx) 0) "DataSheet")
      (check-equal? (hash-ref (XLSX-sheet_index_name_map xlsx) 9) "PieChart3D")

      (check-equal? (hash-count (XLSX-sheet_name_index_map xlsx)) 10)
      (check-equal? (hash-ref (XLSX-sheet_name_index_map xlsx) "DataSheet") 0)
      (check-equal? (hash-ref (XLSX-sheet_name_index_map xlsx) "PieChart3D") 9)

      (check-equal? (hash-count (XLSX-sheet_index_rid_map xlsx)) 10)
      (check-equal? (hash-ref (XLSX-sheet_index_rid_map xlsx) 0) "rId1")
      (check-equal? (hash-ref (XLSX-sheet_index_rid_map xlsx) 9) "rId10")

      (check-equal? (hash-count (XLSX-sheet_rid_rel_map xlsx)) 13)
      (check-equal? (hash-count (XLSX-sheet_index_rel_map xlsx)) 13)
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map xlsx) "rId1") "worksheets/sheet1.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map xlsx) 0) "worksheets/sheet1.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map xlsx) "rId10") "chartsheets/sheet7.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map xlsx) 9) "chartsheets/sheet7.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map xlsx) "rId11") "theme/theme1.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map xlsx) 10) "theme/theme1.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map xlsx) "rId12") "styles.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map xlsx) 11) "styles.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map xlsx) "rId13") "sharedStrings.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map xlsx) 12) "sharedStrings.xml")

      (check-equal? (hash-count (XLSX-shared_strings_map xlsx)) 18)
      (check-equal? (hash-ref (XLSX-shared_strings_map xlsx) "1") "")
      (check-equal? (hash-ref (XLSX-shared_strings_map xlsx) "18") "month/brand")
      
      (check-equal? (get-sheet-names xlsx)
                    '("DataSheet"
                      "DataSheetWithStyle"
                      "DataSheetWithStyle2"
                      "LineChart1"
                      "LineChart2"
                      "LineChart3D"
                      "BarChart"
                      "BarChart3D"
                      "PieChart"
                      "PieChart3D"))
      ))
    
    )))

(run-tests test-reader)
