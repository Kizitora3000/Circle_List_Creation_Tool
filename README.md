サークルリスト作成ツール
====
行きたいサークルの情報を入力するだけで、行きたいサークルをまとめたリストを作ることができるツールです。リストはエクセルで作られます。kivyというアプリケーションを作るツールを用いて、GUIで情報を入力できます。
![github](https://user-images.githubusercontent.com/42823074/45738410-4f9b6600-bc2b-11e8-9c2f-d9ce6063a502.png)

## 使い方
このソフトはpython3.7.0で動作します。そのためpython3.7.0が動かせる環境と、下のほうに書いてある、ライブラリのインストールも必要です。  
main.py, Excel_Init.py, Input_to_Excel.py, quit_Excel.py, CircleListCreationToolApp.kvの五つのファイルを同じフォルダーの中に入れて、main.pyを実行してください。すると下のようなアプリが立ち上がるので、そこで操作していきます。
![github 2](https://user-images.githubusercontent.com/42823074/45791002-9e480f00-bcc2-11e8-8644-f1d7c4c7ce9e.png)

## ライブラリのインストール
pipを用いて必要なライブラリ、Kivyとopenpyxlをインストールします。
以下のコマンドをコマンドプロンプトなどに入力してください。

pip install --upgrade pip wheel setuptools

pip install --upgrade pip wheel setuptools

pip install docutils pygments pypiwin32 kivy.deps.sdl2 kivy.deps.glew

pip install kivy

pip install openpyxl

## 補足
<del>画像だとエクセルに罫線がついていますが、実際にプログラムを実行して作ったエクセルファイルは、罫線がありません。これはopenpyxlの都合上、複数のセルに罫線を引こうとしても、罫線が途中で切れてしまうためです。そのため出来上がったファイルに自分で罫線を付け足すといった対策をとる必要があります。openpyxl v2.6でこのバグが修正されるようなので、v2.6がリリースしたら罫線もpythonで引けるようにしたいです。（やる気は）ないです。</del>  

現在のプログラムを修正することで罫線を引くことができました
