using HongliangSoft.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace narator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private HotKeyHelper _hotkey;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AttachThreadInput(int idAttach, int idAttachTo, bool fAttach);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);

        /// <summary>
        /// Windowフォームアクティブ化処理
        /// </summary>
        /// <param name="handle">フォームハンドル</param>
        /// <returns>true : 成功 / false : 失敗</returns>
        private static bool ForceActive(IntPtr handle)
        {
            const uint SPI_GETFOREGROUNDLOCKTIMEOUT = 0x2000;
            const uint SPI_SETFOREGROUNDLOCKTIMEOUT = 0x2001;
            const int SPIF_SENDCHANGE = 0x2;

            IntPtr dummy = IntPtr.Zero;
            IntPtr timeout = IntPtr.Zero;

            bool isSuccess = false;

            int processId;
            // フォアグラウンドウィンドウを作成したスレッドのIDを取得
            int foregroundID = GetWindowThreadProcessId(GetForegroundWindow(), out processId);
            // 目的のウィンドウを作成したスレッドのIDを取得
            int targetID = GetWindowThreadProcessId(handle, out processId);

            // スレッドのインプット状態を結び付ける
            AttachThreadInput(targetID, foregroundID, true);

            // 現在の設定を timeout に保存
            SystemParametersInfo(SPI_GETFOREGROUNDLOCKTIMEOUT, 0, timeout, 0);
            // ウィンドウの切り替え時間を 0ms にする
            SystemParametersInfo(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, dummy, SPIF_SENDCHANGE);

            // ウィンドウをフォアグラウンドに持ってくる
            isSuccess = SetForegroundWindow(handle);

            // 設定を元に戻す
            SystemParametersInfo(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, timeout, SPIF_SENDCHANGE);

            // スレッドのインプット状態を切り離す
            AttachThreadInput(targetID, foregroundID, false);

            return isSuccess;
        }

        private void ForceActiveWindow()
        {
            var helper = new System.Windows.Interop.WindowInteropHelper(this);
            // タスクバーが点滅しフォーカスはあるのに入力できない状態になるため
            for (int i = 0; i < 3; i++)
                if (ForceActive(helper.Handle)) break;
        }

        public MainWindow()
        {
            InitializeComponent();
            this._hotkey = new HotKeyHelper(this);
            this._hotkey.Register(ModifierKeys.Control,
                                  Key.S,
                                  (_, __) => { ForceActiveWindow(); });
        }

        private void text_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                System.Diagnostics.Process pro = new System.Diagnostics.Process();

                pro.StartInfo.FileName = "E:\\Users\\ユウヤ\\Documents\\workspace\\narrator\\narator\\narator\\SofTalk.exe";
                pro.StartInfo.Arguments = "/close /w:" + text.Text;               // 引数
                pro.StartInfo.CreateNoWindow = false;            // DOSプロンプトの黒い画面を非表示
                pro.StartInfo.UseShellExecute = true;          // プロセスを新しいウィンドウで起動するか否か
                pro.StartInfo.RedirectStandardOutput = false;    // 標準出力をリダイレクトして取得したい

                pro.Start();
                text.Text = "";
            }
        }
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            // HotKeyの登録解除
            this._hotkey.Dispose();
        }
    }
}
