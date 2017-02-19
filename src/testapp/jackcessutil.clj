;; jackcess util
;; 2017/1/20
(ns testapp.jackcessutil)
;;  (:import com.healthmarketscience.jackcess.CursorBuilder))

(use 'clojess.core)
(use 'clojess.util)

;; ここでできんといかんのは２つ。(1)tableの全rec削除。(2)tableにrecの逐次INSERT

;; 2017.1.30
;; delete all rows
(defn delete-rows
  [db tbl-name]
  (let [tbl (table db tbl-name)
        crsr (.getDefaultCursor tbl)]
    (if (> (.getRowCount tbl) (int 0))
      ;; true recがある
      (do
        (. crsr beforeFirst)
        (while (. crsr getNextRow) (. crsr deleteCurrentRow))
        (. db flush))
      ;; else recない
      (println "テーブル %s にはrecありません。\n" tbl-name))))


