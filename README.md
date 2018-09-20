サークルリスト作成ツール
====
行きたいサークルの情報を入力するだけで、行きたいサークルをまとめたリストを作ることができるツールです。リストはエクセルで作られます。kivyというアプリケーションを作るツールを用いて、GUIで情報を入力できます。
![github](https://user-images.githubusercontent.com/42823074/45738410-4f9b6600-bc2b-11e8-9c2f-d9ce6063a502.png)

## 使い方
main.py, Excel_Init.py, Input_to_Excel.py, quit_Excel.py, CircleListCreationToolApp.kvの五つのファイルを同じフォルダーの中に入れて、main.pyを実行してください。すると下のようなアプリが立ち上がるので、そこで操作していきます。

## インストール方法
pipを用いてライブラリをインストールします

kivyのインストール  
kivyとはpythonで書かれたプログラムをGUIで動かせるようにするライブラリです。

pip install --upgrade pip wheel setuptools

pip install --upgrade pip wheel setuptools

pip install docutils pygments pypiwin32 kivy.deps.sdl2 kivy.deps.glew

pip install kivy

poenpyxlのインストール  
pythonでエクセルが動かせるようにするライブラリです。
