(ns testapp.core
  (:gen-class))

(use 'dk.ative.docjure.spreadsheet) ; docjure
(use 'clojess.core)                 ; clojess

;;###################
;; jackcessutil test
;;###################
(require 'testapp.jackcessutil)

;; delete-rows test実行
;;(let [db (open-db "/home/masa/tmpmountp/プロジェクトT/2.案件用資料/5.アップデート/VmInventry2/testDB.mdb")
;;      tbl-name (str "testTable")]
  ;; 全rows削除前のrow数
;;  (printf "rows before delete-rows : %d\n" (. (table db tbl-name) getRowCount))
  ;; 全rowsの削除(delete-rows) の実行
;;  (testapp.jackcessutil/delete-rows db tbl-name)
  ;; 全rows削除後のrow数
;;  (printf "rows after delete-rows : %d\n" (. (table db tbl-name) getRowCount))
;;  (. db close))

;;###########################################
;;Accessのtargettblの全レコードをあらかじめ削除する
;;###########################################
(defn delete-rows
  [dbfile targettblname]
  (let [db (open-db dbfile)
        currentrownum (. (table db targettblname) getRowCount)]
    (if (> currentrownum 0)
      ;;1行以上存在する場合全rowを削除
      (do
        (printf "テーブル %s の %d レコードを全て削除します。\n" targettblname currentrownum)
        (testapp.jackcessutil/delete-rows db targettblname)
        (printf "テーブル %s の %d レコードを全て削除しました。\n" targettblname currentrownum))
      ;;0行の場合なにもしない
      (println "指定テーブルはレコードないので予め全row削除は行いません。")
      )))





;######################################
;;　本appの呼び方
;; lein run <引数１> <引数2> <引数3>
;; <引数1> : MLtougou.xlsm のフルパス名
;; <引数2> : guestOSinfo.mdbのフルパス名
;; <引数3> : 処理対象のテーブル名 ML56=56号機、MLLight=Light号機、MLShinkiban=新基盤、MLGM=GMテーブル
;;#####################################
(defn -main
  [& args]
  (if (= (count args) 3)
    (do
      (println "MLtougou.xlsmのフルパス:" (first args))
      (println "guestOSinfo.mdbのフルパス：" (second args))
      (def targettblname (first (rest (rest args))))
      (println "対象テーブル：" targettblname)
      (if (and (.exists (clojure.java.io/as-file (first args)))
               (.exists (clojure.java.io/as-file (second args)))
               (not= targettblname nil))
        ;;2ファイルとも存在する。ここから処理本体
        (do
          (println "2ファイルとも存在します。")
          ;; 対象テーブルの現レコード一括削除
          (testapp.core/delete-rows (second args) targettblname)
          ;; MLtougou.xlsmの該当シートからguestOSinfoの対象号機のテーブルへレコードコピー
          );;処理本体の終わり
        ;;2ファイルの少なくとも1つが存在しない
        (println "2ファイルの少なくとも1つが存在しません。"))
      )
    ;; 引数不正
    (println "引数が正しくありません。2個必須です。")))



