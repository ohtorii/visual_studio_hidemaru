============================================================================
秀丸エディタから VisualStudio を制御するマクロ

http://d.hatena.ne.jp/ohtorii/
https://github.com/ohtorii/visual_studio_hidemaru
============================================================================

■これは何
    秀丸エディタから VisualStudio を制御するプログラマ向けのマクロです。

    ソースコードを秀丸エディタで書いて、ビルドする度に VisualStudio へ移動する
    のが面倒なので、「ビルド・リビルド・実行...etc」を秀丸エディタから出来るよ
    うにしました。


■出来ること
    ・コンパイル
    ・ビルドのキャンセル
    ・ソリューションのビルド
    ・ソリューションのリビルド
    ・ソリューションのクリーン
    ・プロジェクトのビルド
    ・プロジェクトのリビルド
    ・プロジェクトのクリーン
    ・デバッグ開始
    ・デバッグなしで開始
    ・デバッグの停止
    ・VisualStudioに登録されているファイルからプロジェクトファイル(.hmbook)を
      作る


■動作環境
    秀丸エディタ ver8.03（ver8系なら多分動くと思います）


■インストール
    全ファイル(*.mac/*.exe)を秀丸エディタのマクロディレクトリへコピーしてくだ
    さい。

    visual_studio_menu_simple.mac をマクロ登録してキーアサインしてください。
    マクロ登録が面倒なら「メニュー → マクロ → マクロ実行」から実行することも
    出来ます。


■試しに使ってみる
    ・VisualStudioで既存のプロジェクトを開いてください。（あなたはプログラマな
      のでプロジェクトファイルの10個や20個はあるはずです。）
    ・そのプロジェクトに含まれているソースコードを秀丸エディタで開いてください。
      （例えば、stdafx.cpp MainFrm.cpp などです。）
    ・visual_studio_menu_simple.mac を実行します。
    ・メニューがポップアップするので手始めに「コンパイル」を選択してください。
      そうするとアウトプット枠へビルド状況が表示されます。
      コンパイルエラーがあればファイル名のクリックで、該当のファイルへジャンプ
      できます。


■確認した環境
    VisualStudio 2008
    VisualStudio 2010

    C++/C#/VBプロジェクトで動作を確認しています。
    VisualStudioは複数起動していても正しく動作します。
    Express版は未対応です（本マクロでは認識しません）。


■注意
    このマクロは起動済みのVisualStudioを制御しています。VisualStudioが起動して
    いない状態だと何もしません。


■内部実装
    秀丸エディタで編集しているファイル名を含むVisual Studioを特定して、各種
    コマンド（ビルド／デバッグ・・・）を送りつけて結果を取得しています。

    例えばVisual Studioが二つ起動しているとします、
    ・VisualStudio_0
        c:\my_app\src\main.cpp  （ソリューションエクスプローラーに含まれる）

    ・VisualStudio_1
        c:\project\main.cpp

    このとき、c:\my_app\src\main.cpp を秀丸エディタでリビルドすると 
    VisualStudio_0 に対してリビルドコマンドを送り、ビルド中のメッセージを取得
    しています。

    詳細はソースコードに書いています。


■カスタマイズ
    ・「コンパイル・ビルドのキャンセル・ソリューションのビルド...etc」は１
      コマンド１ファイルになっています。visual_studio_menu_simple.mac は各
      コマンドを呼び出しているサンプル的な位置づけです。

    ・こんなキーアサインにするとVisual Studioと同じ操作になります。
        F7 ビルド       (visual_studio_cf_compile.mac)
        F5 デバッグ開始 (visual_studio_cf_debug.mac)

    ・visual_studio_hidemaru.exe がVisual Studioを制御しています、秀丸に依存し
      ないので他のテキストエディから使用できるかもしれません。


■マクロの説明
    visual_studio_menu_simple.mac           簡易メニュー

    visual_studio_cf_compile.mac            コンパイル
    visual_studio_cf_cancel.mac             ビルドのキャンセル
    visual_studio_cf_debug.mac              デバッグ開始
    visual_studio_cf_debug_stop.mac         デバッグの停止
    visual_studio_cf_project_build.mac      プロジェクトのビルド
    visual_studio_cf_project_clear.mac      プロジェクトのクリーン
    visual_studio_cf_project_rebuild.mac    プロジェクトのリビルド
    visual_studio_cf_run_without_debug.mac  デバッグなしで開始
    visual_studio_cf_solution_build.mac     ソリューションのビルド
    visual_studio_cf_solution_clear.mac     ソリューションのクリーン
    visual_studio_cf_solution_rebuild.mac   ソリューションのリビルド
    visual_studio_cf_hmbook.mac             秀丸の(.hmbook)ファイルを作る

    visual_studio_call.mac                  橋渡しをするマクロ
    visual_studio_hidemaru.exe              VisualStudioを制御する実行ファイル


■その他
    このマクロは Visual Studio に変更を加えることはしません。Visual Studioから
    情報を取り出すだけです。


以上
