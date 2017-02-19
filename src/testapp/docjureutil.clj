;; docjure util
;; 2017/2/8
(ns testapp.docjureutil)

(use 'dk.ative.docjure.spreadsheet)

;;test
(let
    [workbookfile "/home/masa/tmpmountp/プロジェクトT/2.案件用資料/5.アップデート/VmInventry2/MLtougou.xlsm"
     sheetname "MLLight03"]
  (doseq [r1 (->> (load-workbook workbookfile)
       (select-sheet sheetname)
       (select-columns {:B :搭載基盤, :C :ゲストOSNo, :L :システム名})
       (remove nil?)
       rest
       rest
       rest)] ;あたま3行スキップ
    (println (:システム名 r1))
    ))
       
;;test2
(let
    [workbookfile "/home/masa/tmpmountp/プロジェクトT/2.案件用資料/5.アップデート/spreadsheet.xlsx"
     sheetname "Price List"]
;  (->> (load-workbook workbookfile) ;;#=> リストのリスト (("name" "price") ("chikuwa" 100.0))
;       (select-sheet sheetname)
;       row-seq
;       (map cell-seq)           ;rowを順に取り出す
;       (map #(map read-cell %)) ;「cellをread-cellする」をそのrowの各cellにmapする」を各rowにmapする。
;       (map #(map println %))
;       ))
  (->> (load-workbook workbookfile)
       (select-sheet sheetname)
      ; (select-columns {:A :name  :B :price}) ;;これをかませるとエラーになる
       row-seq
       (map cell-seq)
       (map #(map read-cell %))
       (map #(map println %))
       ))

;;test3
(let
    [workbookfile "/home/masa/tmpmountp/プロジェクトT/2.案件用資料/5.アップデート/spreadsheet.xlsx"
     sheetname "Price List"]
  (->> (load-workbook workbookfile) ;;#=> リストのリスト (("name" "price") ("chikuwa" 100.0))
       (select-sheet sheetname)
       (select-columns {:A :name :B :price}) ;ここまでならOKなんや。この後にrow-seqやるとエラー？
       rest ;1行めスキップ
       first
       :name
       println
       ))

;;test4
(let
    [workbookfile "/home/masa/tmpmountp/プロジェクトT/2.案件用資料/5.アップデート/spreadsheet.xlsx"
     sheetname "Price List"]
  (let [r1 (->> (load-workbook workbookfile) ;;#=> リストのリスト (("name" "price") ("chikuwa" 100.0))
       (select-sheet sheetname)
       (select-columns {:A :name :B :price}) ;ここまでならOKなんや。この後にrow-seqやるとエラー？
       rest ;1行めスキップ
       first)]
    (println (:name r1))
    (println (:price r1))))

;;test5 roop each row test
(let
    [workbookfile "/home/masa/tmpmountp/プロジェクトT/2.案件用資料/5.アップデート/spreadsheet.xlsx"
     sheetname "Price List"]
  (doseq [r1 (->> (load-workbook workbookfile) 
       (select-sheet sheetname)
       (select-columns {:A :name :B :price}) 
       rest)] ;1行めスキップ
    (println (:name r1))
    (println (:price r1))))

