# testapp

仮想化統合サーバ(34,56,Light1-5,新基盤）のユーザMLのExcel(MLtougou.xlsm)を
Access(guestOSinfo.mdb)のテーブルにコピーする。

## Usage

lein run testapp <引数1> <引数2> <引数3>

ここで
<引数1> : MLtougou.xlsm のフルパス名
<引数2> : guestOSinfo.mdbのフルパス名
<引数3> : 処理対象のテーブル名 ML56=56号機、MLLight=Light号機、MLShinkiban=新基盤、MLGM=GMテーブル

## leiningenを使ってjavaアプリ(=jar file)へのコンパイルの仕方(2017/3/11)

ここ参考
http://asymmetrical-view.com/2010/06/08/building-standalone-jars-wtih-leiningen.html

    [masa@localhost testapp]$ lein compile
    [masa@localhost testapp]$ lein uberjar
    Warning: specified :main without including it in :aot. 
    Implicit AOT of :main will be removed in Leiningen 3.0.0. 
    If you only need AOT for your uberjar, consider adding :aot :all into your
    :uberjar profile instead.
    Compiling testapp.core
    Created /home/masa/testapp/target/testapp-0.1.0-SNAPSHOT.jar
    Created /home/masa/testapp/target/testapp-0.1.0-SNAPSHOT-standalone.jar

このstand-alone jarを使って実行する方法

    java -jar testapp-0.1.0-SNAPSHOT-standalone.jar <引数1> <引数2> <引数3>  

## License

Copyright © 2016 FIXME

Distributed under the Eclipse Public License either version 1.0 or (at
your option) any later version.
