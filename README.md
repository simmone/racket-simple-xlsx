A Open Xml File format spreadsheet(.xlsx) reader and writer
==================

# Install
    raco pkg install simple-xlsx

# Usage Example
```racket

(require simple-xlsx)

;; read

(with-input-from-excel-file
  "test1.xlsx"
  (lambda ()
    (load-sheet "Sheet1")
    (get-cell-value "A2")
    (get-sheet-names)
    (get-sheet-dimension)
    (with-row
      (lambda (row)
        (printf "~a\n" row)))))

;; write

will added soon.

```
