HOME_DIR="racket-simple-xlsx/simple-xlsx"
echo "lib";cd; cd $HOME_DIR/lib/;racket lib-test.rkt
echo "reader-test1";cd; cd $HOME_DIR/reader/tests/test1/;racket test1-test.rkt
echo "reader-test2";cd; cd $HOME_DIR/reader/tests/test2/;racket test2-test.rkt
echo "reader-test3";cd; cd $HOME_DIR/reader/tests/test3/;racket test3-test.rkt
echo "reader-test4";cd; cd $HOME_DIR/reader/tests/test4/;racket test4-test.rkt
echo "reader-test5";cd; cd $HOME_DIR/reader/tests/test5/;racket test5-test.rkt
echo "reader-test6";cd; cd $HOME_DIR/reader/tests/test6/;racket test6-test.rkt
echo "reader-test7";cd; cd $HOME_DIR/reader/tests/test7/;racket test7-test.rkt

echo "writer-content-type";cd; cd $HOME_DIR/writer/;racket content-type-test.rkt
echo "writer-rels";cd; cd $HOME_DIR/writer/_rels;racket rels-test.rkt
echo "writer-docprops-app";cd; cd $HOME_DIR/writer/docProps;racket docprops-app-test.rkt
echo "writer-docprops-core";cd; cd $HOME_DIR/writer/docProps;racket docprops-core-test.rkt
echo "writer-xl-rels-workbook-xml-rels";cd; cd $HOME_DIR/writer/xl/_rels;racket workbook-xml-rels-test.rkt
echo "writer-xl-cal-chain";cd; cd $HOME_DIR/writer/xl;racket cal-chain-test.rkt
echo "writer-shared-strings";cd; cd $HOME_DIR/writer/xl;racket sharedStrings-test.rkt
echo "writer-styles";cd; cd $HOME_DIR/writer/xl;racket styles-test.rkt
echo "writer-workbook";cd; cd $HOME_DIR/writer/xl;racket workbook-test.rkt

echo "write-read-test-normal";cd; cd $HOME_DIR/writer/tests/;racket test1-test.rkt
echo "write-read-test-10000-row";cd; cd $HOME_DIR/writer/tests/;racket test2-test.rkt
