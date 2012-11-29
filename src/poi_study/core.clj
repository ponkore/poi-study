(ns poi-study.core
  (:import
   (java.io FileOutputStream FileInputStream)
   (java.util Date Calendar)
   (org.apache.poi.hssf.usermodel HSSFWorkbook)
   (org.apache.poi.ss.usermodel Workbook Sheet Cell Row WorkbookFactory DateUtil
                                IndexedColors CellStyle Font CellValue)
   (org.apache.poi.ss.util CellReference AreaReference)))

(defn foo
  "I don't do a whole lot."
  [x]
  (println x "Hello, World!"))

(comment
    Workbook wb = new HSSFWorkbook();
    FileOutputStream fileOut = new FileOutputStream("workbook.xls");
    wb.write(fileOut);
    fileOut.close();

    Workbook wb = new XSSFWorkbook();
    FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
    wb.write(fileOut);
    fileOut.close();
)

(defn empty-xls-save
  "空のエクセルファイルを作る。"
  ([] (new-and-save "workbook-01.xls"))
  ([fname]
     (with-open [out (FileOutputStream. fname)]
       (-> (HSSFWorkbook.)
           (.write out)))))

(comment

Creating Cells

    Workbook wb = new HSSFWorkbook();
    CreationHelper createHelper = wb.getCreationHelper();
    Sheet sheet = wb.createSheet("new sheet");

    // Create a row and put some cells in it. Rows are 0 based.
    Row row = sheet.createRow((short)0);
    // Create a cell and put a value in it.
    Cell cell = row.createCell(0);
    cell.setCellValue(1);

    // Or do it on one line.
    row.createCell(1).setCellValue(1.2);
    row.createCell(2).setCellValue(
        createHelper.createRichTextString("This is a string"));
    row.createCell(3).setCellValue(true);

    // Write the output to a file
    FileOutputStream fileOut = new FileOutputStream("workbook.xls");
    wb.write(fileOut);
    fileOut.close();
)

(defn add-two-sheets-and-save
  "シートとセルを追加する。"
  ([] (new-and-save "workbook-02.xls"))
  ([fname]
     (with-open [out (FileOutputStream. fname)]
       (let [wb (HSSFWorkbook.)
             helper (.getCreationHelper wb)
             sheet (.createSheet wb "new-sheet")
             row (.createRow sheet)]
         (.setCellValue (.createCell row 1) 1.2)
         (.setCellValue (.createCell row 2) (.createRichTextString helper "This is a string"))
         (.setCellValue (.createCell row 3) true)
         (.write out)))))
