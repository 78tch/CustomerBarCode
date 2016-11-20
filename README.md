# CustomerBarCode
【これは何】  
MS Access の使いこなしの習作として、VBAの勉強（復習）を兼ねて、  
MS Access 2010/2013 で、カスタマーバーコード付きの宛名印字ツールを作っていきます。  
元となる住所録は、エクセル形式ファイルから取り込めるようにする。  
ファイルの取り込みの際に、どの列がどのデータ項目なのか指定できるようなUIをつける。  
かつて、似たようなツールをDAO で書いて作ったことがあるので、それをベースにしながら、可能な限りADOに書き直していくこととします。  
Access のVBA はかつてはいろいろ本も出ていたけれど、最近ではあんまり無いので、  
「オブジェクトブラウザでオブジェクトの種類やら構文やら引数の種類やらを確認しつつ」  
「[MSDN](https://msdn.microsoft.com/ja-jp/library/cc407851.aspx) で詳しい意味を調べる」  
という流れで進めていきます。  
  
![object browser](https://github.com/78tch/CustomerBarCode/blob/master/images/Peek201611191534.gif)  
