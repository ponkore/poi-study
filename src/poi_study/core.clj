(ns poi-study.core
  (:import
   (java.io FileOutputStream FileInputStream)
   (java.util Date Calendar)
   (org.apache.poi.hssf.usermodel HSSFWorkbook)
   (org.apache.poi.ss.usermodel Workbook Sheet Cell Row WorkbookFactory DateUtil
                                IndexedColors CellStyle Font CellValue)
   (org.apache.poi.ss.util CellReference AreaReference CellRangeAddress RegionUtil)))

;;; see http://poi.apache.org/spreadsheet/quick-guide.html

(defn add-two-sheets-and-save
  "シートとセルを追加する。"
  ([] (add-two-sheets-and-save "workbook-02.xls"))
  ([fname]
     (with-open [out (FileOutputStream. fname)]
       (let [wb (HSSFWorkbook.)
             helper (.getCreationHelper wb)
             sheet (.createSheet wb "new-sheet")]
         (let [row (.createRow sheet 0)]
           (-> row (.createCell 0) (.setCellValue 1.0)) ;; 1 だとコンパイルエラー
           (-> row (.createCell 1) (.setCellValue 1.2))
           (-> row (.createCell 2) (.setCellValue (.createRichTextString helper "This is a string.")))
           (-> row (.createCell 3) (.setCellValue true)))
         (.write wb out)))))

(defn draw-border-example
  "罫線を描いてみる。"
  ([] (draw-border-example "workbook-03.xls"))
  ([fname]
     (with-open [out (FileOutputStream. fname)]
       (let [workbook (HSSFWorkbook.)
             helper (.getCreationHelper workbook)
             sheet (.createSheet workbook "new-sheet")
             cellstyle (.createCellStyle workbook)]
         (doto cellstyle
           (.setBorderBottom CellStyle/BORDER_THIN)
           (.setBottomBorderColor (.getIndex IndexedColors/BLACK))
           (.setBorderLeft CellStyle/BORDER_THIN)
           (.setLeftBorderColor (.getIndex IndexedColors/GREEN))
           (.setBorderRight CellStyle/BORDER_THIN)
           (.setRightBorderColor (.getIndex IndexedColors/BLUE))
           (.setBorderTop CellStyle/BORDER_THIN)
           (.setTopBorderColor (.getIndex IndexedColors/BLACK)))
         (-> (.createRow sheet 1)
             (.createCell  1)
             (.setCellStyle cellstyle))
         (.write workbook out)))))

(defn misc-example
  "いろいろ"
  ([] (misc-example "workbook-04.xls"))
  ([fname]
     (with-open [out (FileOutputStream. fname)]
       (let [workbook (HSSFWorkbook.)
             helper (.getCreationHelper workbook)
             sheet (.createSheet workbook "新シート")
             row (.createRow sheet 1)
             row2 (.createRow sheet 2)
             cell (.createCell row 1)
             region (CellRangeAddress/valueOf "B2:E5")]
         (.setCellValue cell "これはマージのテストです。")
         (.addMergedRegion sheet region)
         (doto CellStyle/BORDER_MEDIUM_DASHED
           (RegionUtil/setBorderBottom region sheet workbook)
           (RegionUtil/setBorderTop region sheet workbook)
           (RegionUtil/setBorderLeft region sheet workbook)
           (RegionUtil/setBorderRight region sheet workbook))
         (doto (.getIndex IndexedColors/AQUA)
           (RegionUtil/setBottomBorderColor region sheet workbook)
           (RegionUtil/setTopBorderColor region sheet workbook)
           (RegionUtil/setLeftBorderColor region sheet workbook)
           (RegionUtil/setRightBorderColor region sheet workbook))
         (.write workbook out)))))

(defn templating-xls
  "テンプレートファイルを読み込んで、セルの値だけ書き換えたものを別ファイルとして出力"
  ([] (templating-xls "template-01.xls" "result-01.xls"))
  ([template output]
     (with-open [input (FileInputStream. template)]
       (with-open [out (FileOutputStream. output)]
         (let [workbook (WorkbookFactory/create input)
               helper (.getCreationHelper workbook)
               sheet (-> workbook (.getSheetAt 0))]
           (when-let [row0 (-> sheet (.getRow 0))]
             (doseq [v (iterator-seq (.cellIterator row0))]
               ;; 本当は Cell の型に応じてメソッドを使い分け
               (.setCellValue v (str "[" (-> v (.getRichStringCellValue) (.getString)) "]"))))
           (let [row (.createRow sheet 1)]
             (-> row (.createCell 0) (.setCellValue 1.0)) ;; 1 だとコンパイルエラー
             (-> row (.createCell 1) (.setCellValue 1.2))
             (-> row (.createCell 2) (.setCellValue (.createRichTextString helper "This is a string.")))
             (-> row (.createCell 3) (.setCellValue true)))
           (.write workbook out))))))

; .getRichStringCellValue ???
(defmulti read-cell #(.getCellType %))
(defmethod read-cell Cell/CELL_TYPE_BLANK   [cell] nil)
(defmethod read-cell Cell/CELL_TYPE_BOOLEAN [cell] (.getBooleanCellValue cell))
(defmethod read-cell Cell/CELL_TYPE_ERROR   [cell] (.getErrorCellValue cell))
(defmethod read-cell Cell/CELL_TYPE_FORMULA [cell] (.getCellFormula cell))
(defmethod read-cell Cell/CELL_TYPE_NUMERIC [cell] (if (DateUtil/isCellDateFormatted cell)
                                                     (.getDateCellValue cell)
                                                     (.getNumericCellValue cell)))
(defmethod read-cell Cell/CELL_TYPE_STRING  [cell] (.getStringCellValue cell))

(defmulti set-border! (fn [_ aside _] aside))
(defmethod set-border! :top    [cellstyle _ border-style] (.setBorderTop cellstyle border-style))
(defmethod set-border! :bottom [cellstyle _ border-style] (.setBorderBottom cellstyle border-style))
(defmethod set-border! :left   [cellstyle _ border-style] (.setBorderLeft cellstyle border-style))
(defmethod set-border! :right  [cellstyle _ border-style] (.setBorderRight cellstyle border-style))

(defmulti set-border-color! (fn [_ aside _] aside))
(defmethod set-border-color! :top    [cellstyle _ color] (.setBorderTopColor cellstyle (.getIndex color)))
(defmethod set-border-color! :bottom [cellstyle _ color] (.setBorderBottomColor cellstyle (.getIndex color)))
(defmethod set-border-color! :left   [cellstyle _ color] (.setBorderLeftColor cellstyle (.getIndex color)))
(defmethod set-border-color! :right  [cellstyle _ color] (.setBorderRightColor cellstyle (.getIndex color)))


(defn draw-border-example-2
  "罫線を描いてみる。"
  ([] (draw-border-example-2 "workbook-03.xls"))
  ([fname]
     (with-open [out (FileOutputStream. fname)]
       (let [workbook (HSSFWorkbook.)
             helper (.getCreationHelper workbook)
             sheet (.createSheet workbook "new-sheet")
             cellstyle (.createCellStyle workbook)]
         (doto cellstyle
           (set-border! :bottom CellStyle/BORDER_THIN)
           (set-border-color! :bottom  IndexedColors/BLACK)
           (set-border! :left CellStyle/BORDER_THIN)
           (set-border-color! :left IndexedColors/GREEN)
           (set-border! :right CellStyle/BORDER_THIN)
           (set-border-color! :right IndexedColors/BLUE)
           (set-border! :top CellStyle/BORDER_THIN)
           (set-border-color! :top IndexedColors/BLACK))
         (-> (.createRow sheet 1)
             (.createCell  1)
             (.setCellStyle cellstyle))
         (.write workbook out)))))


#_(defn set-cell! [^Cell cell value]
  (if (nil? value)
    (let [^String null nil]
      (.setCellValue cell null)) ;do not call setCellValue(Date) with null
    (let [converted-value (cond (number? value) (double value)
                                true value)]
      (.setCellValue cell converted-value)
      (if (date-or-calendar? value)
        (apply-date-format! cell "m/d/yy")))))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
#_(comment
    ;; 以下メモ

    ;; style 関連：こんな感じの関数が使いやすい？
    (set-cell-style cellstyle :border :bottom :border-thin)
    (set-cell-style cellstyle :border-color :bottom :border-thin)
    (set-cell-style cellstyle :background-color :white)

    (set-cell-value cell-or-region :number 1)
    (set-cell-value cell-or-region :text "hello, world")
    (set-cell-value cell-or-region :formula "=SUM(A1:C3)")

    ;; sheet-seq / row-seq / cell-seq みたいな階層が作れるか
    (doseq [sheet (sheet-seq workbook)]
      (doseq [row (row-seq sheet)]
        (doseq [cell (cell-seq row)]
          (cell-value cell))))
    ;; こんな感じ

    ;; あとは、cell、row、sheet、workbook の各属性をもつ map との相互変換 (to-bean、to-map 的なもの)

    ;; style 関連も整理したい
)
