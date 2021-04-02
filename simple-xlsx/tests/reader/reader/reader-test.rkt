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
     (lambda ()
      (check-equal? (xlsx-sheet-count) 10)
      
      (check-equal? (hash-count (XLSX-sheet_index_id_map (*CURRENT_XLSX*))) 10)
      (check-equal? (hash-ref (XLSX-sheet_index_id_map (*CURRENT_XLSX*)) 0) "1")
      (check-equal? (hash-ref (XLSX-sheet_index_id_map (*CURRENT_XLSX*)) 9) "10")

      (check-equal? (hash-count (XLSX-sheet_index_name_map (*CURRENT_XLSX*))) 10)
      (check-equal? (hash-ref (XLSX-sheet_index_name_map (*CURRENT_XLSX*)) 0) "DataSheet")
      (check-equal? (hash-ref (XLSX-sheet_index_name_map (*CURRENT_XLSX*)) 9) "PieChart3D")

      (check-equal? (hash-count (XLSX-sheet_name_index_map (*CURRENT_XLSX*))) 10)
      (check-equal? (hash-ref (XLSX-sheet_name_index_map (*CURRENT_XLSX*)) "DataSheet") 0)
      (check-equal? (hash-ref (XLSX-sheet_name_index_map (*CURRENT_XLSX*)) "PieChart3D") 9)

      (check-equal? (hash-count (XLSX-sheet_index_rid_map (*CURRENT_XLSX*))) 10)
      (check-equal? (hash-ref (XLSX-sheet_index_rid_map (*CURRENT_XLSX*)) 0) "rId1")
      (check-equal? (hash-ref (XLSX-sheet_index_rid_map (*CURRENT_XLSX*)) 9) "rId10")

      (check-equal? (hash-count (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*))) 13)
      (check-equal? (hash-count (XLSX-sheet_index_rel_map (*CURRENT_XLSX*))) 13)
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) "rId1") "worksheets/sheet1.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) 0) "worksheets/sheet1.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) "rId10") "chartsheets/sheet7.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) 9) "chartsheets/sheet7.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) "rId11") "theme/theme1.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) 10) "theme/theme1.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) "rId12") "styles.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) 11) "styles.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) "rId13") "sharedStrings.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) 12) "sharedStrings.xml")

      (check-equal? (hash-count (XLSX-shared_strings_map (*CURRENT_XLSX*))) 18)
      (check-equal? (hash-ref (XLSX-shared_strings_map (*CURRENT_XLSX*)) "1") "")
      (check-equal? (hash-ref (XLSX-shared_strings_map (*CURRENT_XLSX*)) "18") "month/brand")
      
      (check-equal? (xlsx-sheet-names)
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
