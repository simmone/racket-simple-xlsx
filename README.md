A Open Xml File format spreadsheet(.xlsx) reader and writer
==================

# Install
    raco pkg install simple-xlsx

# Usage Example
```racket

    (require simple-xlsx)

    ;; write
    ;; data type must be string or number

    (let ([xlsx (new xlsx-data%)])
        (send xlsx add-sheet '(("chenxiao" "cx") (1 2 34 100 456.34)) "Sheet1")
        (send xlsx add-sheet '((1 2 3 4)) "Sheet2")
        (send xlsx add-sheet '(()) "Sheet3")
        (write-xlsx-file xlsx "test1.xlsx"))
        
    ;; read specific cell
        
    (with-input-from-xlsx-file "test1.xlsx"
        (lambda (xlsx)
            (load-sheet "Sheet1" xlsx)
                (get-cell-value "A1" xlsx) ; "chenxiao"
                (get-cell-value "B1" xlsx) ; "xiaomin"
                (get-cell-value "E2" xlsx) ; 456.34))

    ;; loop for row

    (with-input-from-xlsx-file "test1.xlsx"
        (lambda (xlsx)
            (load-sheet "Sheet1" xlsx)
                (with-row xlsx
                    (lambda (row)
                        (printf "~a\n" (first row)))))) ;; chenxiao 1
```

# To do list

    column format: width, color, etc;
    
