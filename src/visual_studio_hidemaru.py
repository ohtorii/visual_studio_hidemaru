# -*- coding: utf-8 -*-
"""コマンドラインからVisualStudioの操作を行なうスクリプト

使い方概要
    コマンド実行
    「コマンド実行」が終わるまで待つ
    実行結果をコンソール出力かファイル出力で取得する。

実装済み：
    起動しているvcのリストを取得する
    プロセスIDからVCを特定する
    ビルド結果(出力ウインドウ)をコンソール出力する
    ビルド結果(出力ウインドウ)をファイルへ出力する
    ソリューションファイルを開く
    ソリューションファイルを閉じる
    ソリューションのビルド
    ソリューションのリビルド
    ソリューションのクリーン
    スタートアッププロジェクトをビルドする
    スタートアッププロジェクトをリビルドする
    スタートアッププロジェクトをクリーンする
    指定したファイル(cpp/c)をコンパイルする。

    VCを終了する。
    デバッグで開始
    起動しているvcのプロジェクトからファイルリスト(cpp/h)を取得する

    任意プロジェクトのビルド
    ビルドのキャンセル
    デバッグなしで実行
    デバッグ開始
    デバッグの停止
    VisualStudioのウインドウを最前面に持ってくる
    ファイル名+行番号を秀丸の形式にしてタグジャンプできるようにする。

todo:
    VisualStudioのウインドウを最小化する。
    ソリューションからVCを特定する。
    任意プロジェクトのリビルド
    任意プロジェクトのクリーン

    タスク一覧
    ビルド後のエラー一覧(EnvDTE80 名前空間が必要、後で)
    秀丸の「プロジェクト」に登録してあるソリューションをVCで開く。
    秀丸の「プロジェクト」に登録してあるソリューションをVCで閉じる。
    構成マネージャーで「リリース・デバッグ」「win32/win64」のリストを取得する。
    構成マネージャーで「リリース・デバッグ」「win32/win64」を切り替える。
    スタートアッププロジェクトに登録する
    外部で変更されたソースをVC側で自動ロードするようにする。
        ->毎回聞いてきてうっとおしい。特に外部のテキストエディと連動するときは。
    プロパティ取得時にCOMの例外が発生する、プロパティ取得は別関数にして適当にリトライしたほうが良い。
    （例外が発生しやすい箇所には手動でリトライしている。）



メモ:
    EnvDTE80 : VS2005
    EnvDTE90 : VS2008
    EnvDTE100: VS2010

コマンドラインへの返値：
    成功    0
    失敗    1

参考：
    http://msdn.microsoft.com/ja-jp/library/k3dys0y2(v=VS.90).aspx
    http://msdn.microsoft.com/ja-jp/library/ms228772.aspx#Y458

連絡先：
    http://d.hatena.ne.jp/ohtorii/
"""
import sys,os,re,time,string,pythoncom,win32com,pywintypes
import win32com.client
from optparse import OptionParser
from win32com.client.gencache import EnsureDispatch
from win32com.client import constants



"""
#タイムアウトを回避する（テスト）
import win32ui
import win32uiole
win32uiole.AfxOleInit()
win32uiole.SetRetryReply(-1)
win32uiole.SetMessagePendingDelay(1000);
win32uiole.EnableNotRespondingDialog(True);
win32uiole.EnableBusyDialog(True);
"""

#例外を再送出するかどうか
#コンソールへ例外の詳細が出力されます
#   False   デバッグ中
#   True    テキストエディタと連動する
g_exception_dont_raise = True

vsBuildStateNotStarted = 1   # Build has not yet been started.
vsBuildStateInProgress = 2   # Build is currently in progress.
vsBuildStateDone = 3         # Build has been completed

vsProjectItemKindMisc           ="{66A2671F-8FB5-11D2-AA7E-00C04F688DDE}".lower()   #"その他のファイル" プロジェクト項目。
vsProjectItemKindPhysicalFile   ="{6BB5F8EE-4483-11D3-8BCF-00C04F8EC28C}".lower()   #システム内のファイル。
vsProjectItemKindPhysicalFolder ="{6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C}".lower()   #システム内のフォルダ。
vsProjectItemKindSolutionItems  ="{66A26722-8FB5-11D2-AA7E-00C04F688DDE}".lower()   #"ソリューション項目" プロジェクト内の項目。
vsProjectItemKindSubProject     ="{EA6618E8-6E24-4528-94BE-6889FE16485C}".lower()   #プロジェクトの下のサブプロジェクト。ProjectItem.Kind によって返された場合、ProjectItem.SubProject は Project オブジェクトを返します。
vsProjectItemKindVirtualFolder  ="{6BB5F8F0-4483-11D3-8BCF-00C04F8EC28C}".lower()   #仮想フォルダ。システムに物理的には存在しない、ソリューション エクスプローラのフォルダです。

#ウインドウ
vsWindowKindKindStartPage       ="{387CB18D-6153-4156-9257-9AC3F9207BBE}".lower()   #
vsWindowKindAutoLocals          ="{F2E84780-2AF1-11D1-A7FA-00A0C9110051}".lower()   #デバッガ ウィンドウ。
vsWindowKindCallStack           ="{0504FF91-9D61-11D0-A794-00A0C9110051}".lower()   #[呼び出し履歴] ウィンドウ。
vsWindowKindClassView           ="{C9C0AE26-AA77-11D2-B3F0-0000F87570EE}".lower()   #[クラス ビュー] ウィンドウ。
vsWindowKindCommandWindow       ="{28836128-FC2C-11D2-A433-00C04F72D18A}".lower()   #コマンド ウィンドウ。
vsWindowKindDocumentOutline     ="{25F7E850-FFA1-11D0-B63F-00A0C922E851}".lower()   #[ドキュメント アウトライン] ウィンドウ。
vsWindowKindDynamicHelp         ="{66DBA47C-61DF-11D2-AA79-00C04F990343}".lower()   #[ダイナミック ヘルプ] ウィンドウ。
vsWindowKindFindReplace         ="{CF2DDC32-8CAD-11D2-9302-005345000000}".lower()   #[検索] ダイアログ ボックス、[置換] ダイアログ ボックス。
vsWindowKindFindResults1        ="{0F887920-C2B6-11D2-9375-0080C747D9A0}".lower()   #[検索結果 1] ウィンドウ。
vsWindowKindFindResults2        ="{0F887921-C2B6-11D2-9375-0080C747D9A0}".lower()   #[検索結果 2] ウィンドウ。
vsWindowKindFindSymbol          ="{53024D34-0EF5-11D3-87E0-00C04F7971A5}".lower()   #[シンボルの検索] ダイアログ ボックス。
vsWindowKindFindSymbolResults   ="{68487888-204A-11D3-87EB-00C04F7971A5}".lower()   #[シンボルの検索結果] ウィンドウ。
vsWindowKindLinkedWindowFrame   ="{9DDABE99-1D02-11D3-89A1-00C04F688DDE}".lower()   #リンク ウィンドウ フレーム。
vsWindowKindLocals              ="{9DDABE99-1D02-11D3-89A1-00C04F688DDE}".lower()   #デバッガ ウィンドウ。
vsWindowKindMacroExplorer       ="{07CD18B4-3BA1-11D2-890A-0060083196C6}".lower()   #[マクロ エクスプローラ] ウィンドウ。
vsWindowKindMainWindow          ="{9DDABE98-1D02-11D3-89A1-00C04F688DDE}".lower()   #Visual Studio .NET IDE のウィンドウ。
vsWindowKindObjectBrowser       ="{269A02DC-6AF8-11D3-BDC4-00C04F688E50}".lower()   #[オブジェクト ブラウザ] ウィンドウ。
vsWindowKindOutput              ="{34E76E81-EE4A-11D0-AE2E-00A0C90FFFC3}".lower()   #出力ウィンドウ。
vsWindowKindProperties          ="{EEFA5220-E298-11D0-8F78-00A0C9110057}".lower()   #[プロパティ] ウィンドウ。
vsWindowKindResourceView        ="{2D7728C2-DE0A-45b5-99AA-89B609DFDE73}".lower()   #リソース エディタ。
vsWindowKindServerExplorer      ="{74946827-37A0-11D2-A273-00C04F8EF4FF}".lower()   #[サーバー エクスプローラ] ウィンドウ。
vsWindowKindSolutionExplorer    ="{3AE79031-E1BC-11D0-8F78-00A0C9110057}".lower()   #[ソリューション エクスプローラ] ウィンドウ。
vsWindowKindTaskList            ="{3AE79031-E1BC-11D0-8F78-00A0C9110057}".lower()   #[タスク一覧] ウィンドウ。
vsWindowKindThread              ="{E62CE6A0-B439-11D0-A79D-00A0C9110051}".lower()   #デバッガ ウィンドウ。
vsWindowKindToolbox             ="{B1E99781-AB81-11D0-B683-00AA00A3EE26}".lower()   #ツールボックス。
vsWindowKindWatch               ="{90243340-BD7A-11D0-93EF-00A0C90F2734}".lower()   #[ウォッチ] ウィンドウ。
vsWindowKindWebBrowser          ="{E8B06F52-6D01-11D2-AA7D-00C04F990343}".lower()   #Visual Studio .NET で管理されている Web ブラウザのウィンドウ。
vsDocumentKindText              ="{8E7B96A8-E33D-11D0-A6D5-00C04FB67F6A}".lower()   #テキスト エディタで開かれたテキスト ドキュメント。
#databaseSchemaViewToolWindow   {6fe3356d-3860-495c-b0d4-6a646a907a79}  #後で調べる

#出力ウインドウ中の種類「ビルド・ビルドの順序...etc」
gGuidBuildOrder                     ="{2032b126-7c8d-48ad-8026-0e0348004fc0}".lower()
gGuidBuildOutput                    ="{1BD8A850-02D1-11d1-BEE7-00A0C913D1F8}".lower()
gGuidDebugOutput                    ="{FC076020-078A-11D1-A7DF-00A0C9110051}".lower()
gGuidGUID_OutWindowDebugPane        ="{FC076020-078A-11D1-A7DF-00A0C9110051}".lower()
gGuidGUID_OutWindowGeneralPane      ="{3c24d581-5591-4884-a571-9fe89915cd64}".lower()
gGuidSID_SVsGeneralOutputWindowPane ="{65482c72-defa-41b7-902c-11c091889c83}".lower()

g_exit_success  = 0
g_exit_error    = 1
g_cmd_result    = False

g_output_text_interval = 0.2

#あとで
g_help = "visual_studio_hidemaru.exe command arg1 arg2 ...\n"   \
+   "cmd_dte_list   print DTE list.\n"

g_codec='cp932'

def _cmp_str(s1,s2):
    if isinstance(s1,unicode):
        if isinstance(s2,unicode):
            return s1==s2
        else:
            return s1==s2.decode(g_codec)
    else:
        if isinstance(s2,unicode):
            return s1.decode(g_codec)==s2
        else:
            return s1==s2

def _to_unicode(s):
    if isinstance(s,unicode):
        return s
    else:
        return s.decode(g_codec)

def _to_mbc(s):
    if isinstance(s,unicode):
        return s.encode(g_codec)
    else:
        return s

def _vs_msg(s):
    _vs_print("vs>"+s)

def _vs_print(s):
    print(_to_mbc(s));

def _to_bool_string(a):
    return str(int(a))

def _to_bool(s):
    if isinstance(s,basestring):
        x = s.strip()
        if s.lower() in ("true","1"):
            return True
        if s.lower() in ("false","0"):
            return False
    return s

def _dte_prop(obj,name):
    """
    下記、例外が時々発生するのでその対策。（対策といっても対処療法ですが・・・）
    pywintypes.com_error: (-2147418111, '呼び出し先が呼び出しを拒否しました。', None, None)
    """
    if obj is None:
        #デバッグ目的で例外を発生させる
        return getattr(obj,name)

    try:
        return getattr(obj,name)
    except:
        pass

    #0.5秒リトライして取得できないなら諦める。
    old = time.clock()
    while(True):
        try:
            return getattr(obj,name)
        except:
            pass
        time.sleep(0.01)
        if 0.5 < (time.clock()-old):
            break;
    #デバッグ目的で例外を発生させる
    return getattr(obj,name)

_wsh = None
def _get_wsh():
    global _wsh
    if not _wsh:
        try:
            _wsh = win32com.client.Dispatch ('WScript.Shell')
        except pywintypes.com_error:
            _vim_msg ('Cannot access WScript.Shell')
    return _wsh

def _dte_activate(dte):
    MainWindow=_dte_prop(dte,"MainWindow")
    if 0:
        #VC2010
        #TypeError: 'NoneType' object is not callable
        MainWindow.Activate()
    _get_wsh().AppActivate(_dte_prop(MainWindow,"Caption"))
    return True


def _get_dte_from_pid(pid=0):
    """プロセスIDからDTEを検索する。

    pid=None    起動しているVisualStudioを全て返す
    pid!=None   pidからVisualStudioを検索して返す、見つからなければNoneを返す。

    todo:ソリューションのファイル名から探す関数を追加する。
    """
    rot = pythoncom.GetRunningObjectTable()
    try:
        lst_dte     = []
        rot_enum    = rot.EnumRunning()

        while 1:
            monikers = rot_enum.Next()
            if not monikers:
                break

            display_name = monikers[0].GetDisplayName(pythoncom.CreateBindCtx(0), None)
            if display_name.startswith ('!VisualStudio.DTE'):
                dte_pid = 0
                pos = display_name.rfind(':')
                if 0<=pos:
                    try:
                        dte_pid = int(display_name[pos+1:])
                    except ValueError:
                        dte_pid = 0

                if pid and (pid != dte_pid):
                    continue

                obj         = rot.GetObject(monikers[0])
                if 1:
                    interface   = obj.QueryInterface (pythoncom.IID_IDispatch)
                    dte = win32com.client.Dispatch(interface)
                else:
                    dte = win32com.client.Dispatch("VisualStudio.DTE")

                Solution= _dte_prop(dte,"Solution")
                dte_sln = unicode(_dte_prop(Solution,"FullName"))

                if None is dte_pid:
                    dte_pid=0

                if None is dte_sln:
                    dte_sln=""

                lst_dte.append((dte, dte_pid, dte_sln, ))
        return lst_dte
    finally:
        rot = None



def cmd_dte_list():
    """起動しているVisualStudioのリストを表示する。

    ＊フォーマット
        num_list    //(int)リスト数
        pid_0       //(int)プロセスID。タスクマネージャで確認できる。
        sln_0       //(string)空白ならソリューションが開いていない
        pid_1
            :
            :
        (num_list個続く)
        :end            //終端記号

    ＊例
        2
        5428
        C:\my_spp\build\my_app.sln
        4658
        d:\develop\hoge_app\hoge_app.sln
        :
    """
    dte_list = _get_dte_from_pid()
    _vs_print(len(dte_list))
    for o in dte_list:
        _vs_print(o[1])
        _vs_print(o[2])
    global g_cmd_result
    g_cmd_result=True
    return True

def get_dte_obj(pid_sln):
    """プロセスIDまたはソリューション名からDTEを返す。
    """
    try:
        pid_sln = int(pid_sln)
    except ValueError:
        #solution file name.
        return

    #process id.
    det_list = _get_dte_from_pid(pid_sln)
    if len(det_list):
        return det_list[0][0]
    return


def _dte_get_window(dte, ObjectKind):
    Windows = _dte_prop(dte,"Windows")
    for w in Windows:
        if ObjectKind == str(_dte_prop(w,"ObjectKind")).lower():
            return w
        Link = _dte_prop(w,"LinkedWindowFrame")
        if None is not Link:
            if ObjectKind==str(_dte_prop(Link,"ObjectKind")).lower():
                return Link

def _get_output_pane(dte):
    window = _dte_get_window(dte, vsWindowKindOutput)
    if not window:
        return
    Object=_dte_prop(window,"Object")
    OutputWindowPanes=_dte_prop(Object,"OutputWindowPanes")
    for pane in OutputWindowPanes:
        Guid=_dte_prop(pane,"Guid")
        if str(Guid).lower() == gGuidBuildOutput:
            return pane


def cmd_outputwindow_clear(pid):
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    dte.ExecuteCommand("Edit.ClearOutputWindow")
    global g_cmd_result
    g_cmd_result=True
    return True


def _reformat_output_text(txt):
    """テキスト各行の先頭に '1>' があれば削除する。
    秀丸でタグジャンプさせるために必要。
    あなたが使用しているテキストエディタがうまく処理してくれるなら、この関数を呼び出す必要はありません。
    """
    lst = []
    for line in txt.splitlines():
        if (2<=len(line)) and (line[0] in string.digits) and (line[1]==">"):
            lst.append(line[2:])
        else:
            lst.append(line)
    return '\n'.join(lst)

def _dte_output(vs_pid, con_output, fn_output="", clear=False,linetop=1):
    """出力ウインドウを内容を取得する
    vs_pid      - プロセスID
    con_output  - コンソール出力するかどうか(bool)
    fn_output   - ファイル名へ出力する、空白なら出力しない
    clear       - ウインドウの内容を消去するかどうか
    linetop     - 選択の先頭行
    返値 -  成功すれば次の行を返す
            それ以外は負値を返す
    """
    clear = _to_bool(clear)
    dte = get_dte_obj(vs_pid)
    if not dte:
        return -1

    pane = _get_output_pane(dte)
    if None is pane:
        return -1
    TextDocument=_dte_prop(pane,"TextDocument")
    sel = _dte_prop(TextDocument,"Selection")

    if None is sel:
        return -1

    if linetop<=1:
        sel.SelectAll()
    else:
        sel.GotoLine(linetop)
        sel.EndOfDocument(True)

    BottomPoint=_dte_prop(sel,"BottomPoint")
    LastLine = _dte_prop(BottomPoint,"Line")
    SelText = unicode(_dte_prop(sel,"Text"))

    if clear:
        dte.ExecuteCommand("Edit.ClearOutputWindow")
    else:
        sel.Collapse()

    if SelText.strip("\n"):
        #空白行ではない
        if 1:
            text = _reformat_output_text(SelText)
        else:
            text = '\n'.join(SelText.splitlines())

        if con_output:
            _vs_print(text)

        if fn_output:
            fp = open(fn_output, 'w')
            fp.write(text)
            fp.close()

    return LastLine

def cmd_output_console(vs_pid, clear=0):
    """出力ウインドウを内容を取得する
    vs_pid      - プロセスID
    clear       - ウインドウの内容を消去するかどうか
    """
    if 0 <= _dte_output(vs_pid, True, "", clear):
        global g_cmd_result
        g_cmd_result=True
        return True

def cmd_output_file(vs_pid, fn_output, clear=False):
    """出力ウインドウを内容を取得する
    vs_pid      - プロセスID
    fn_output   - 出力ファイル名
    clear       - ウインドウの内容を消去するかどうか
    """
    if 0 <= _dte_output(vs_pid, False, fn_output, clear):
        global g_cmd_result
        g_cmd_result=True
        return True

def _check_for_build(pid):
    """「ビルド・リビルド・クリーン・コンパイル」が終わっているか調べる
    即時復帰
    返値：  True    ビルド中
            False   「ビルド中」以外
    """
    dte = get_dte_obj(pid)
    if not dte:
        return False
    Solution        = _dte_prop(dte,"Solution")
    SolutionBuild   = _dte_prop(Solution,"SolutionBuild")
    BuildState      = _dte_prop(SolutionBuild,"BuildState")
    return BuildState == vsBuildStateInProgress

def cmd_wait_for_build(pid):
    """「ビルド・リビルド・クリーン・コンパイル」が終わるまで待つ。
    完了復帰
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False

    while _check_for_build(pid):
        time.sleep (0.1)
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_solution_close(pid,SaveFirst=True):
    """ソリューションを閉じる
    """
    SaveFirst=_to_bool(SaveFirst)
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    dte.Solution.Close(SaveFirst)
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_solution_open(pid,Filename):
    """ソリューションのファイルを開く
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    dte.Solution.Open(Filename)
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_solution_clean(pid):
    """ソリューションをクリーン
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    dte.ExecuteCommand('Build.CleanSolution')
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_solution_build(pid):
    """ソリューションをビルドする
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    dte.ExecuteCommand('Build.BuildSolution')
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_solution_rebuild(pid):
    """ソリューションをリビルドする
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    dte.ExecuteCommand('Build.RebuildSolution')
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_startup_project_clean(pid):
    """スタートアッププロジェクトのクリーン
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    dte.ExecuteCommand('Build.CleanSelection')
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_startup_project_build(pid):
    """スタートアッププロジェクトのビルド
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    dte.ExecuteCommand('Build.BuildSelection')
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_startup_project_rebuild(pid):
    """スタートアッププロジェクトのリビルド
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    dte.ExecuteCommand('Build.RebuildSelection')
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_set_startup_project(pid,project_abs_filename):
    """スタートアッププロジェクトを設定する
    project_abs_filename - プロジェクトの絶対パス名
                            (Ex.) c:/Projects/my_app/my_app/my_app.vcxproj
    """
    project_abs_filename = _to_unicode(project_abs_filename)
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    project_abs_filename = os.path.normpath(project_abs_filename).lower()
    if not os.path.isfile(project_abs_filename):
        _vs_msg("File not found."+project_abs_filename)
        return False
    dte.Solution.SolutionBuild.StartupProjects=project_abs_filename
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_debug(pid):
    """デバッグ開始
    即時復帰

    「ソースに変更あり・前回のビルドに失敗...etc」の条件でVisualStudioの
    問い合わせダイアログがポップアップします、そのときは「はい・いいえ」を
    選択するまで復帰しません。
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False

    #デバッグを開始するのでVisualStudioをアクティブにしておく。
    _dte_activate(dte)

    Debugger=_dte_prop(dte,"Debugger")
    Debugger.Go(0)
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_debug_stop(pid):
    """デバッグの停止
    即時復帰
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    #デバッグ中以外に呼び出すと例外発生。(pywintypes.com_error)
    Debugger=_dte_prop(dte,"Debugger")
    try:
        Debugger.Stop(0)
    except pywintypes.com_error:
        #デバッグ中ではない
        pass

    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_run(pid):
    """デバッグなしで実行
    即時復帰

    「ソースに変更あり・前回のビルドに失敗...etc」の条件でVisualStudioの
    問い合わせダイアログがポップアップします、そのときは「はい・いいえ」を
    選択するまで復帰しません。
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    _dte_activate(dte)
    dte.ExecuteCommand("Debug.StartWithoutDebugging")
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_activate(pid):
    """VisualStudioのウインドウを最前面にもってくる。
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    _dte_activate(dte)
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_quit(pid):
    """Visual Studioを終了する。
    VCで編集中のファイルがあっても確認なしで終了する。
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    dte.Quit()
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_project_list(pid):
    """プロジェクトリストの表示。
    ＊引数
        pid プロセスID

    ＊フォーマット
        num_list    //(int)リスト数
        full_name   //(string)プロジェクトの完全名(～.vcxproj)
        uniq_name   //(string)一意のプロジェクト名
        name        //(string)プロジェクト名
        kind        //(string)種類や型を表す GUID 文字列
        saved       //(bool)最後に保存されたとき、または開かれたとき以降、変更されているかどうかを示す値

    ＊例
        C:/Users/hoge/documents/visual studio 2010/Projects/test_app/test_app/test_app.vcxproj
        test_app/test_app.vcxproj
        test_app
        {8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}
        1
        C:/Users/hoge/documents/visual studio 2010/Projects/test_app/test_mfc_app/test_mfc_app.vcxproj
        test_mfc_app/test_mfc_app.vcxproj
        test_mfc_app
        {8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}
        1

    ＊kind
        ソリューション フォルダ {66A26720-8FB5-11D2-AA7E-00C04F688DDE}
        Visual Basic            {F184B08F-C81C-45F6-A57F-5ABD9991F28F}
        Visual C#               {FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}
        Visual C++              {8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}
        Visual J#               {E6FDF86B-F3D1-11D4-8576-0002A516ECE8}
        Web プロジェクト        {E24C65DC-7377-472b-9ABA-BC803B73C61A}
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    prjs = dte.Solution.Projects
    for o in prjs:
        _vs_print(o.FullName)
        _vs_print(o.UniqueName)
        _vs_print(o.Name)
        _vs_print(o.Kind)
        _vs_print(_to_bool_string(o.Saved))
    global g_cmd_result
    g_cmd_result=True
    return True

class project_file:
    __slots__=("name","kind","depth","saved","is_open","full_name")


def check_file_ext(filename, ext):
    if ext:
        if os.path.splitext(filename)[1].lower() in ext:
            return filename
    else:
        return filename

def iter_project_items(ProjectItems):
    """
    ProjectItems - ProjectItems コレクション
    """
    for f in ProjectItems:
        yield f
        try:
            o = f.ProjectItems
        except:
            pass
        else:
            for a in iter_project_items(o):
                yield a

def list_filenames(ProjectItems,Ext,Depth):
    """
    ProjectItems - ProjectItems コレクション
    """
    result = []
    for f in ProjectItems:
        if check_file_ext(f.Name,Ext):
            item = project_file()
            item.name = f.Name
            item.kind = f.Kind
            item.depth = Depth
            item.saved = f.Saved
            item.is_open = f.IsOpen
            if 1==f.FileCount:
                item.full_name = f.FileNames(1)
            else:
                item.full_name=""
            result.append(item)
        try:
            o = f.ProjectItems
        except:
            pass
        else:
            result = result + list_filenames(o,Ext,Depth+1)
    return result

def cmd_project_file_list(pid,project_name,ext_filter=""):
    """プロジェクトに登録されているファイルの一覧を表示する
    ＊引数
        pid             プロセスID
        project_name    プロジェクト名（my_app/test...など）
                        "*"ならば全プロジェクトを全プロジェクトを対象とする
        ext_filter      表示する拡張子(""=全て。".cpp;.h;.cs")
    ＊フォーマット
        num_list    //(int)リスト数
        filename    //(string)ファイル名
        kind        //(string)種類や型を表す GUID 文字列
        saved       //(bool)最後に保存されたとき、または開かれたとき以降、変更されているかどうかを示す値
        is_open     //(bool)特定の種類のビューで開いているかどうかを示す値
        depth       //(int)階層の深さ(ルート=0、子供=1、孫=2...)
    ＊Kind
        物理ファイル        {6BB5F8EE-4483-11D3-8BCF-00C04F8EC28C}
        物理フォルダ        {6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C}
        仮想フォルダ        {6BB5F8F0-4483-11D3-8BCF-00C04F8EC28C}
        サブプロジェクト    {EA6618E8-6E24-4528-94BE-6889FE16485C}
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False

    ext_set=None
    if(ext_filter):
        ext_set = set()
        for o in ext_filter.split(";"):
            o = o.lower().strip()
            if o:
                ext_set.add(o)

        if not ext_set:
            ext_set=None


    prjs = dte.Solution.Projects
    for o in prjs:
        if "*" == project_name:
            process=True
        else:
            process = o.Name.lower() == project_name.lower()

        if process:
            file_list = list_filenames(o.ProjectItems,ext_set,0)
            #+1はプロジェクト分
            _vs_print(1 + len(file_list))
            #プロジェクト
            _vs_print(o.Name)
            _vs_print(o.Kind)
            _vs_print(_to_bool_string(o.saved))
            _vs_print(_to_bool_string(True))
            _vs_print(0)
            #ファイル一覧
            for f in file_list:
                _vs_print(f.name)
                _vs_print(f.kind)
                _vs_print(_to_bool_string(f.saved))
                _vs_print(_to_bool_string(f.is_open))
                _vs_print(f.depth)
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_solution_configuration(pid):
    """ソリューション構成
        Debug/Release
    作成中（そもそも必要なのか・・・）
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    sc = dte.Solution.SolutionBuild.SolutionConfigurations
    for i in range(sc.Count):
        o = sc.Item(1+i)
        _vs_print(o.Name)
    global g_cmd_result
    g_cmd_result=True
    return True


def cmd_project_build(pid,project_full_name,solution_configuration=""):
    """プロジェクトをビルドする
    project_full_name       プロジェクトのフルネーム(～.test_mfc_app.vcxproj)
    solution_configuration  Release/Debug...etc
    即時復帰
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    if not solution_configuration:
        solution_configuration = unicode(dte.Solution.SolutionBuild.ActiveConfiguration.Name)
    dte.Solution.SolutionBuild.BuildProject(solution_configuration,project_full_name,False)
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_build_cancel(pid):
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    #if _check_for_build(pid)
    dte.ExecuteCommand("Build.Cancel")
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_error_list(pid):
    """エラー一覧を取得する。
    EnvDTE80 名前空間が必要、後で実装する。
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False

def cmd_hidemaru_hmbook(pid):
    """
    pid - VisualStudioのプロセス番号
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    Solution    = _dte_prop(dte,"Solution")
    prjs        = _dte_prop(Solution,"Projects")
    SolutionFullName= _dte_prop(Solution,"FullName")

    hmbook      = []
    for o in prjs:
        file_list = list_filenames(o.ProjectItems,None,0)
        #プロジェクト
        hmbook.append('"%s",group,expand' % o.Name)
        for f in file_list:
            kind    = f.kind.lower()
            fname   = unicode(f.name)
            space   = "\t" * (f.depth + 1)
            if vsProjectItemKindPhysicalFile == kind:
                hmbook.append('%s"%s",name="%s"' % (space, f.full_name, os.path.split(fname)[1]))
            else:
                hmbook.append('%s"%s",group,expand' % (space,fname))

    #utf16-le bom
    hmbook_filename = os.path.splitext(SolutionFullName)[0] + ".hmbook"
    f = open(hmbook_filename,"wb")
    try:
        f.write("\n".join(hmbook).encode("utf16"))
    except:
        f.close()
        raise

    _vs_msg("Save:"+hmbook_filename)
    global g_cmd_result
    g_cmd_result=True
    return True

def _search_file(dte, abs_filename):
    """
    abs_filename - あらかじめos.path.normpath().lower()を行なう
    return - None or [Document,Project]
    Document型 or None
    """
    prjs = dte.Solution.Projects
    for o in prjs:
        for item in iter_project_items(o.ProjectItems):
            if (str(item.Kind).lower() == vsProjectItemKindPhysicalFile) and (1==item.FileCount):
                abs_fn = _to_unicode(item.FileNames(1))
                if _cmp_str(abs_filename, os.path.normpath(abs_fn).lower()):
                    #Solution        = _dte_prop(dte,"Solution")
                    #SolutionBuild   = _dte_prop(Solution,"SolutionBuild")
                    #print SolutionBuild.StartupProjects
                    return item,o

def _get_pid_from_filename(abs_filename):
    """ファイル名からDTEを取得する
    abs_filename -  ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp
    return - None or [pid,Document,Project]
    """
    abs_filename = _to_unicode(abs_filename)
    abs_filename = os.path.normpath(abs_filename).lower()
    dte_list = _get_dte_from_pid()
    for o in dte_list:
        pid = o[1]
        dte = get_dte_obj(pid)
        if None is dte:
            break
        ret = _search_file(dte,abs_filename)
        if None is not ret:
            return pid,ret[0],ret[1]

def cmd_file_compile(pid,abs_filename,enable_check=True):
    """ファイルをコンパイルする

    abs_filename -  ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    abs_filename = _to_unicode(abs_filename)
    abs_filename = os.path.normpath(abs_filename).lower()
    if not os.path.isfile(abs_filename):
        _vs_msg("File not found."+abs_filename)
        return False

    #ソリューションに含まれていなくてもファイルを開いてビルドできてしまうので、その対策。
    if enable_check:
        ret = _search_file(dte,abs_filename)
        if None is ret:
            return False

    dte.ItemOperations.OpenFile(abs_filename)
    dte.ExecuteCommand("Build.Compile")
    global g_cmd_result
    g_cmd_result=True
    return True


def cmd_te_file_compile(abs_filename,wait=False,print_output=False,clear_output=False):
    """ファイルをコンパイルする。（テキストエディタ向け）
    ファイルが含まれるVisual Studioを探し出してビルドする。
    VisualStudioの「メニュー  ->  ビルド  ->  コンパイル」と同じ動作。

    abs_filename-   ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp

    wait        -   True    ビルド終了まで待つ（完了復帰）
                    False   即時復帰

    print_output-   True    コンパイル結果をコンソールへ表示
                    False   何もしない

    clear_output-   True    VisualStudioの出力ウインドウをクリアする
                    False   何もしない

    """
    wait        =_to_bool(wait)
    print_output=_to_bool(print_output)
    clear_output=_to_bool(clear_output)
    ret = _get_pid_from_filename(abs_filename)
    if None is ret:
        return False
    pid=ret[0]
    Project=ret[2]

    if cmd_file_compile(pid, abs_filename,False):
        if wait:
            if print_output:
                LastLine=1
                while _check_for_build(pid):
                    LastLine = _dte_output(pid,True,"",False,LastLine)
                    time.sleep(g_output_text_interval)
                LastLine = _dte_output(pid,True,"",False,LastLine)
            else:
                cmd_wait_for_build(pid)
            if clear_output:
                cmd_outputwindow_clear(pid)
        #ビルドできれば成功としておく。
        global g_cmd_result
        g_cmd_result=True
        return True
    return False

def cmd_openfile(pid,abs_filename,line_no=1,column_no=1):
    """ファイルをコンパイルする

    abs_filename -  ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp
    """
    dte = get_dte_obj(pid)
    if not dte:
        _vs_msg("Not found process")
        return False
    abs_filename = _to_unicode(abs_filename)
    abs_filename = os.path.normpath(abs_filename).lower()
    if not os.path.isfile(abs_filename):
        _vs_msg("File not found."+abs_filename)
        return False

    #ソリューションに含まれていなくてもファイルを開いてビルドできてしまうので、その対策。
    ret = _search_file(dte,abs_filename)
    if None is ret:
        return False

    dte.ItemOperations.OpenFile(abs_filename)
    dte.ActiveDocument.Selection.StartOfDocument()
    dte.ActiveDocument.Selection.MoveToLineAndOffset(int(line_no),int(column_no),False)
    global g_cmd_result
    g_cmd_result=True
    return True

def _te_main(func,abs_filename,wait,print_output,clear_output,change_startup_project=False):
    wait        =_to_bool(wait)
    print_output=_to_bool(print_output)
    clear_output=_to_bool(clear_output)

    ret = _get_pid_from_filename(abs_filename)
    if None is ret:
        return False
    pid=ret[0]
    Project=ret[2]
    if change_startup_project:
        if not cmd_set_startup_project(pid,Project.FullName):
            return False

    if func(pid):
        if wait:
            if print_output:
                LastLine=1
                while _check_for_build(pid):
                    LastLine = _dte_output(pid,True,"",False,LastLine)
                    time.sleep(g_output_text_interval)
                LastLine = _dte_output(pid,True,"",False,LastLine)
            else:
                cmd_wait_for_build(pid)
            if clear_output:
                cmd_outputwindow_clear(pid)
        global g_cmd_result
        g_cmd_result=True
        return True
    return False

def _te_main2(func,abs_filename):
    ret = _get_pid_from_filename(abs_filename)
    if None is ret:
        return False
    pid=ret[0]
    Project=ret[2]
    func(pid)
    global g_cmd_result
    g_cmd_result=True
    return True

def cmd_te_solution_build(abs_filename,wait=False,print_output=False,clear_output=False):
    """ソリューションをビルドする（テキストエディタ向け）
    ファイルが含まれるVisual Studioを探し出してソリューションをビルドする。
    VisualStudioの「メニュー  ->  ビルド  ->  ソリューションのビルド」と同じ動作。

    abs_filename-   ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp

    wait        -   True    ビルド終了まで待つ（完了復帰）
                    False   即時復帰

    print_output-   True    コンパイル結果をコンソールへ表示
                    False   何もしない

    clear_output-   True    VisualStudioの出力ウインドウをクリアする
                    False   何もしない

    """
    return _te_main(cmd_solution_build, abs_filename,wait,print_output,clear_output)

def cmd_te_solution_rebuild(abs_filename,wait=False,print_output=False,clear_output=False):
    """ソリューションをリビルドする（テキストエディタ向け）
    ファイルが含まれるVisual Studioを探し出してソリューションをビルドする。
    VisualStudioの「メニュー  ->  ビルド  ->  ソリューションのリビルド」と同じ動作。

    abs_filename-   ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp

    wait        -   True    ビルド終了まで待つ（完了復帰）
                    False   即時復帰

    print_output-   True    コンパイル結果をコンソールへ表示
                    False   何もしない

    clear_output-   True    VisualStudioの出力ウインドウをクリアする
                    False   何もしない

    """
    return _te_main(cmd_solution_rebuild, abs_filename,wait,print_output,clear_output)

def cmd_te_solution_clear(abs_filename,wait=False,print_output=False,clear_output=False):
    """ソリューションをクリアする（テキストエディタ向け）
    ファイルが含まれるVisual Studioを探し出してソリューションをビルドする。
    VisualStudioの「メニュー  ->  ビルド  ->  ソリューションのリビルド」と同じ動作。

    abs_filename-   ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp

    wait        -   True    ビルド終了まで待つ（完了復帰）
                    False   即時復帰

    print_output-   True    コンパイル結果をコンソールへ表示
                    False   何もしない

    clear_output-   True    VisualStudioの出力ウインドウをクリアする
                    False   何もしない

    """
    return _te_main(cmd_solution_clean, abs_filename,wait,print_output,clear_output)


def cmd_te_project_build(abs_filename,wait=False,print_output=False,clear_output=False):
    return _te_main(cmd_startup_project_build, abs_filename,wait,print_output,clear_output,True)

def cmd_te_project_rebuild(abs_filename,wait=False,print_output=False,clear_output=False):
    return _te_main(cmd_startup_project_rebuild, abs_filename,wait,print_output,clear_output,True)

def cmd_te_project_clean(abs_filename,wait=False,print_output=False,clear_output=False):
    return _te_main(cmd_startup_project_clean, abs_filename,wait,print_output,clear_output,True)


def cmd_te_run_without_debug(abs_filename):
    """デバッグなしで開始する（テキストエディタ向け）
    ファイルが含まれるVisual Studioを探し出してデバッグなしで開始する。

    abs_filename-   ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp
    """
    return _te_main(cmd_run, abs_filename,False,False,False)

def cmd_te_debug(abs_filename):
    """デバッグ開始（テキストエディタ向け）
    ファイルが含まれるVisual Studioを探し出してデバッグ開始する。

    abs_filename-   ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp
    """
    return _te_main(cmd_debug,abs_filename,False,False,False)

def cmd_te_debug_stop(abs_filename):
    """デバッグの停止（テキストエディタ向け）
    ファイルが含まれるVisual Studioを探し出してデバッグ停止する。

    abs_filename-   ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp
    """
    return _te_main(cmd_debug_stop,abs_filename,False,False,False)

def cmd_te_cancel(abs_filename):
    """ビルドをキャンセルする（テキストエディタ向け）
    ファイルが含まれるVisual Studioを探し出してビルドをキャンセルする。

    abs_filename-   ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp
    """
    return _te_main2(cmd_build_cancel,abs_filename)

def cmd_te_activate(abs_filename):
    """最前面に持ってくる（テキストエディタ向け）
    ファイルが含まれるVisual Studioを探し出して最前面に持ってくる。

    abs_filename-   ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp
    """
    return _te_main2(cmd_activate,abs_filename)

def cmd_te_hmbook(abs_filename):
    """VisualStudioのプロジェクト構成を秀丸エディタのプロジェクトファイル(.hmbook)へ変換する。
    (.hmbook)はソリューション(.sln)と同じディレクトリに保存されます。
    abs_filename-   ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp
    """
    return _te_main2(cmd_hidemaru_hmbook,abs_filename)

def cmd_te_switch(abs_filename, line_no=1, column_no=1):
    """ファイルをVisualStudioで開いてフォーカスを移す。
    abs_filename-   ファイル名の絶対パス
                    (Ex.) c:/project/my_app/src/main.cpp
    line_no - カーソル行
    column_no - カーソル列
    """
    ret = _get_pid_from_filename(abs_filename)
    if None is ret:
        return False
    pid=ret[0]
    Project=ret[2]
    if not cmd_openfile(pid,abs_filename,line_no,column_no):
        return False
    return cmd_activate(pid)

def escape(s):
    s = s.replace('"', '\"')
    return s.replace("\\", "\\\\")

def main():
    argv = sys.argv
    if len(argv)<2:
        #_vs_print(g_help)
        return False

    if not argv[1].startswith("cmd_"):
        #_vs_print("Invalid command name.")
        return False

    exp = argv[1] + "("
    if 2 < len(argv):
        for s in argv[2:]:
            exp = exp + '"' + escape(s) + '",'
        exp = exp[:-1]

    exp = exp + ")"
    exec(exp)


def test():
    FullName1="C:\\Users\\hoge\\documents\\visual studio 2010\\Projects\\test_20110220\\test_mfc_app\\test_mfc_app.vcxproj"
    FullName2="C:\\Users\\hoge\\documents\\visual studio 2010\\Projects\\test_20110220\\test_20110220\\test_20110220.vcxproj"

    #プロセスID（タスクマネージャーで確認する）
    g_pid= 5968

    FileName1="C:\\開発\\test\\test.cpp"
    FileName2="C:\\Users\\hoge\\documents\\visual studio 2010\\Projects\\test_20110220\\test_20110220\\test_20110220.cpp"
    cmd_file_compile(g_pid,FileName1)
    #cmd_debug_stop(g_pid)
    #cmd_debug(g_pid)
    #cmd_run(g_pid)
    #cmd_project_build(g_pid,FullName1)
    #cmd_build_cancel(g_pid)

    #cmd_project_list(g_pid)
    #cmd_project_file_list(g_pid,"test_20110220",".h");
    """
    cmd_solution_clean(g_pid)
    cmd_wait_for_build(g_pid)
    cmd_output(g_pid,True,"",True)
    """

    cmd_dte_list()


def start():
    try:
        if 1:
            main()
        else:
            test()
    except pywintypes.com_error, err:
        if g_exception_dont_raise:
            pass
        else:
            for o in err:
                print o
            raise
    except:
        if g_exception_dont_raise:
            pass
        else:
            raise

    if g_cmd_result:
        sys.exit(g_exit_success)
    else:
        sys.exit(g_exit_error)

if __name__ == "__main__":
    start()
