(ns testapp.core
  (:gen-class))

(use 'dk.ative.docjure.spreadsheet) ; docjure
(use 'clojess.core)                 ; clojess

(require 'testapp.jackcessutil)

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

;;###########################################
;;MLtougou.xlmsの所定のシートを
;;guestOSinfo.mdbの指定のテーブルに書き込む
;;###########################################
(defn excel2ml
  [workbookfilename accessfilename targettblname]
  (let
      [targettbl (table (open-db accessfilename) targettblname)]
    ;;targettblname に対応する 読み込み必要なexcelのシート群
    (case targettblname
      "ML56" (let
                 [targetsheetset ["ML5-6"]]
               (doseq [targetsheetname targetsheetset]
                 (println "targetsheetname : " targetsheetname)
                 (doseq [r1 (->> (load-workbook workbookfilename)
                      (select-sheet targetsheetname)
                      (select-columns {:C :ゲストOSNo , :D :部署 , :E :室 , :F :Gr , :G :氏名 , :H :メールアドレス , :I :管理 , :J :維持 , :K :ML , :L :システム名 , :M :システム名称補足 , :N :ステータス})
                      rest
                      rest
                      rest)] ;;頭3行スキップ
                   (do
                     (printf "インベントリ名:%s 氏名:%s システム名:%s\n" (:ゲストOSNo r1) (:氏名 r1) (:システム名 r1))
                     (let
                         [rowmap1 (hash-map "InventryName" (:ゲストOSNo r1) ,
                                             "Bu"       (:部署 r1),
                                             "Sitsu"    (:室 r1),
                                             "Gr"       (:Gr r1),
                                             "Name"     (:氏名 r1),
                                             "Email"    (:メールアドレス r1),
                                             "Kanri"    (:管理 r1),
                                             "Iji"      (:維持 r1),
                                             "ML"       (:ML r1),
                                             "SystemName" (:システム名 r1),
                                             "SupplementName" (:システム名補足 r1),
                                             "Status"   (:ステータス r1))]
                       (. targettbl addRowFromMap rowmap1))
                     );;excel1行ごとの処理
                   
                   );;end-of-excel1行ごとの処理 doseq
                 );; end-of- targetsheetset doseq
               );;end-of-ML56
      
      "MLLight" (let
                 [targetsheetset ["MLLight01" "MLLight02" "MLLight03" "MLLight04" "MLLight05"]]
               (doseq [targetsheetname targetsheetset]
                 (println "targetsheetname : " targetsheetname)
                 (doseq [r1 (->> (load-workbook workbookfilename)
                      (select-sheet targetsheetname)
                      (select-columns {:C :ゲストOSNo , :D :部署 , :E :室 , :F :Gr , :G :氏名 , :H :メールアドレス , :I :管理 , :J :維持 , :K :ML , :L :システム名 , :M :システム名称補足 , :N :ステータス})
                      rest
                      rest
                      rest)] ;;頭3行スキップ
                   (do
                     (printf "インベントリ名:%s 氏名:%s システム名:%s\n" (:ゲストOSNo r1) (:氏名 r1) (:システム名 r1))
                     (let
                         [rowmap1 (hash-map "InventryName" (:ゲストOSNo r1) ,
                                             "Bu"       (:部署 r1),
                                             "Sitsu"    (:室 r1),
                                             "Gr"       (:Gr r1),
                                             "Name"     (:氏名 r1),
                                             "Email"    (:メールアドレス r1),
                                             "Kanri"    (:管理 r1),
                                             "Iji"      (:維持 r1),
                                             "ML"       (:ML r1),
                                             "SystemName" (:システム名 r1),
                                             "SupplementName" (:システム名補足 r1),
                                             "Status"   (:ステータス r1))]
                       (if (get rowmap1 "InventryName") (. targettbl addRowFromMap rowmap1)))
                     );;excel1行ごとの処理 do block
                   );;end-of-excel1行ごとの処理 doseq
                 );; end-of- targetsheetset doseq
               );;end-of-MLLight

      
      "MLShinkiban" (let
                 [targetsheetset ["ML標準#1" "ML標準#2" "ML標準#3" "ML隔離" "ML高可用" "ML開発"]]
               (doseq [targetsheetname targetsheetset]
                 (println "targetsheetname : " targetsheetname)
                 (case targetsheetname
                   "ML標準#1" (def kiban "標準1")
                   "ML標準#2" (def kiban "標準2")
                   "ML標準#3" (def kiban "標準3")
                   "ML隔離"   (def kiban "隔離")
                   "ML高可用" (def kiban "高可用")
                   "ML開発"   (def kiban "開発"))
                 (doseq [r1 (->> (load-workbook workbookfilename)
                      (select-sheet targetsheetname)
                      (select-columns {:C :ゲストOSNo , :D :部署 , :E :室 , :F :Gr , :G :氏名 , :H :メールアドレス , :I :管理 , :J :維持 , :K :ML , :L :システム名 , :M :システム名称補足 , :N :ステータス})
                      rest
                      rest
                      rest)] ;;頭3行スキップ
                   (do
                     (printf "インベントリ名:%s 氏名:%s システム名:%s\n" (:ゲストOSNo r1) (:氏名 r1) (:システム名 r1))
                     (let
                         [rowmap1 (hash-map "InventryName" (:ゲストOSNo r1) ,
                                            "Kiban" kiban ,
                                             "Bu"       (:部署 r1),
                                             "Sitsu"    (:室 r1),
                                             "Gr"       (:Gr r1),
                                             "Name"     (:氏名 r1),
                                             "Email"    (:メールアドレス r1),
                                             "Kanri"    (:管理 r1),
                                             "Iji"      (:維持 r1),
                                             "ML"       (:ML r1),
                                             "SystemName" (:システム名 r1),
                                             "SupplementName" (:システム名補足 r1),
                                             "Status"   (:ステータス r1))]
                       (if (get rowmap1 "InventryName") (. targettbl addRowFromMap rowmap1)))
                     );;excel1行ごとの処理 do block
                   );;end-of-excel1行ごとの処理 doseq
                 );; end-of- targetsheetset doseq
               );;end-of-MLLight


      
      "MLGM" (let
                 [targetsheetset ["室長・GM "]]
               (doseq [targetsheetname targetsheetset]
                 (println "targetsheetname : " targetsheetname)
                 (def currentKubun nil) ; 現rowが室長rec or GMrecかのフラグ
                 (doseq [r1 (->> (load-workbook workbookfilename)
                      (select-sheet targetsheetname)
                      (select-columns {:A :区分 , :B :部 , :C :室 , :D :Gr , :E :氏名 , :F :メールアドレス , :G :Flag34 , :H :Flag56 , :I :FlagSTD , :J :FlagISO , :K :FlagHIG , :L :FlagTDE , :M :FlagL3 , :N :FlagL4})
                      rest
                      rest
                      rest)] ;;頭3行スキップ
                   (do
                     (if (and (= currentKubun nil) (= (:区分 r1) "GM")) (def currentKubun "GM"))
                     (if (and (not= (:部 r1) "") (not= (:部 r1) nil)) (def currentBu (:部 r1)))
                     (printf "部:%s 室:%s 氏名:%s\n" currentBu (:室 r1) (:氏名 r1))
                     (let
                         [rowmap1 (hash-map  "Bu"      currentBu,
                                             "Sitsu"   (:室 r1),
                                             "Gr"      (:Gr r1),
                                             "Name"    (:氏名 r1),
                                             "Email"   (:メールアドレス r1),
                                             "Flag34"  (:Flag34 r1),
                                             "Flag56"  (:Flag56 r1),
                                             "FlagSTD" (:FlagSTD r1),
                                             "FlagISO" (:FlagISO r1),
                                             "FlagHIG" (:FlagHIG r1),
                                             "FlagTDE" (:FlagTDE r1),
                                             "FlagL3"  (:FlagL3 r1),
                                             "FlagL4"  (:FlagL4 r1))]
                       (if (and (= currentKubun "GM") (not= (:氏名 r1) "") (not= (:氏名 r1) nil)) (. targettbl addRowFromMap rowmap1)))
                     );;excel1行ごとの処理
                   
                   );;end-of-excel1行ごとの処理 doseq
                 );; end-of- targetsheetset doseq
               );;end-of-ML56


      
      (println "targettblname の指定が不正")

      );; end-of-case

    ;;targettblname に対応する 読み込み必要なexcelの該当シート群それぞれをtargettblに書き込む。
;;    (if (> (count targetsheetset) 0)
;;      ;;該当シートが正常。各該当シートをtargettblに書き込む処理
;;      (doseq [targetsheetname targetsheetset]
;;        (println "targetsheet : " targetsheetname)
;;        )
;;      ;;対象シートが不正
;;      (println "対象シートが0個です。"))

    
    )) ; end-of-defn excel2ml
    
    

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
          (testapp.core/excel2ml (first args) (second args) targettblname)
          );;処理本体の終わり
        ;;2ファイルの少なくとも1つが存在しない
        (println "2ファイルの少なくとも1つが存在しません。"))
      )
    ;; 引数不正
    (println "引数が正しくありません。2個必須です。")))



