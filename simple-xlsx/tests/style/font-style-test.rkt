#lang racket

(require rackunit/text-ui rackunit)

(require "../../style/font-style.rkt")

(define test-font-style
  (test-suite
   "test-font-style"

   (test-case
    "test-set-font-style1"

    (let ([font_style (font-style-from-hash-code "10<p>Arial<p>0000ff")])
      (check-equal? (FONT-STYLE-size font_style) 10)
      (check-equal? (FONT-STYLE-color font_style) "0000FF")
      (check-equal? (FONT-STYLE-name font_style) "Arial")
      (check-equal? (FONT-STYLE-hash_code font_style) "10<p>Arial<p>0000FF"))

    (let ([font_style (font-style-from-hash-code "")])
      (check-false font_style)))

   (test-case
    "test-font-style<=?"

    (let ([font1_style #f]
          [font2_style (font-style-from-hash-code "11<p>Arial<p>0000ff")])
      (check-true (font-style<? font1_style font2_style))
      (check-false (font-style=? font1_style font2_style)))

    (let ([font1_style (font-style-from-hash-code "11<p>Arial<p>0000ff")]
          [font2_style #f])
      (check-false (font-style<? font1_style font2_style))
      (check-false (font-style=? font1_style font2_style)))

    (let ([font1_style #f]
          [font2_style #f])
      (check-false (font-style<? font1_style font2_style))
      (check-true (font-style=? font1_style font2_style)))

    (let ([font1_style (font-style-from-hash-code "10<p>Arial<p>0000ff")]
          [font2_style (font-style-from-hash-code "10<p>Arial<p>0000ff")])
      (check-false (font-style<? font1_style font2_style))
      (check-true (font-style=? font1_style font2_style))))

    (let ([font1_style (font-style-from-hash-code "10<p>Arial<p>0000ff")]
          [font2_style (font-style-from-hash-code "11<p>Arial<p>0000ff")])
      (check-true (font-style<? font1_style font2_style))
      (check-false (font-style=? font1_style font2_style)))

    (let ([font1_style (font-style-from-hash-code "10<p>Arial<p>0000ff")]
          [font2_style (font-style-from-hash-code "10<p>Brial<p>0000ff")])
      (check-true (font-style<? font1_style font2_style))
      (check-false (font-style=? font1_style font2_style)))

    (let ([font1_style (font-style-from-hash-code "10<p>Arial<p>0000ff")]
          [font2_style (font-style-from-hash-code "10<p>Arial<p>000fff")])
      (check-true (font-style<? font1_style font2_style)))

   ))

(run-tests test-font-style)
