//------------------------------------------------------------------------------
// トレースログ クラス                                                          
//------------------------------------------------------------------------------
// [NOTE]                                                                       
//   トレースログを手軽なグローバルインスタンスとして実装                       
//   トレースログファイルが増殖しないように、プログラム名をキーとして           
//   ひとつのプロセスでのみトレースログが出力可能としています                   
//   同時に複数のプロセスからログ出力可能とする場合は                           
//   ログファイル名に pid 付与などの対処をしてください                          
//                                                                              
// [注意]                                                                       
//   スタートアップルーチンで PreProcessing() を実行してからメインフォーム実行  
//   メインフォーム終了時に PostProcessing() を実行してください                 
//                                                                              
//------------------------------------------------------------------------------
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace Hoge
{
    public static class TraceLogging
    {
        #region プロパティ
        public static TraceLogObject? LogObj { get; set; }  // トレースログ本体
        #endregion

        #region 前処理/後処理
        /// <summary>
        /// 前処理 - インスタンス初期化
        /// </summary>
        public static void PreProcessing()
        {
            // トレースログ - 初期化
            LogObj = new TraceLogObject();
        }
        /// <summary>
        /// 後処理 - インスタンス開放
        /// </summary>
        public static void PostProcessing()
        {
            // トレースログ - クローズ
            LogObj?.Close();
        }
        #endregion

        #region トレースログ本体
        /// <summary>
        /// トレースログ本体
        /// </summary>
        public class TraceLogObject
        {
            #region 内部変数
            private FileStream? Stream = null; 
            private StreamWriter? Writer = null;
            #endregion

            #region 生成/消滅
            /// <summary>
            /// 生成/消滅
            /// </summary>
            public TraceLogObject()
            {
                string progname = Process.GetCurrentProcess().ProcessName;
                string terminal = GetTerminalName();

                // 同一 Windows ユーザーでRDP運用を考えてファイル名に端末名も付与
                string filename = string.Format("{0}-{1}.log", progname, terminal);

                // ワークファイル 排他オープン
                Writer = OpenWorkFile(filename, Encoding.UTF8, out Stream);
                try
                {
                    Writer?.WriteLine(string.Format("開始 - {0}", DateInfo()));
                    Writer?.WriteLine(string.Empty);
                }
                catch { /* NOP */ }
            }
            ~TraceLogObject()
            {
                Close();
            }
            /// <summary>
            /// クローズ
            /// </summary>
            public void Close()
            {
                try
                {
                    Writer?.WriteLine(string.Empty);
                    Writer?.WriteLine(string.Format("終了 - {0}", DateInfo()));
                    Writer?.Close();
                    Stream?.Close();
                }
                catch { /* NOP */ }

                Writer = null;
                Stream = null;
            }
            /// <summary>
            /// 日時情報
            /// </summary>
            private string DateInfo()
            {
                return DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            }
            #endregion

            #region トレース出力
            /// <summary>
            /// １行出力
            /// </summary>
            public void WriteLine(string msg)
            {
                try
                {
                    Writer?.WriteLine(msg);
                }
                catch { /* NOP */ }
            }
            /// <summary>
            /// データフラッシュ
            /// </summary>
            public void DataFlush()
            {
                try
                {
                    Writer?.Flush();
                }
                catch { /* NOP */ }
            }
            #endregion

            #region ワークファイル処理
            /// <summary>
            /// ワークファイル 排他オープン
            /// </summary>
            private StreamWriter? OpenWorkFile(string filename, Encoding encoding, out FileStream? stream)
            {
                // RDPの場合、GetTmpPath() はセッション付きフォルダとなるが、
                // セッション付きフォルダが存在しない場合も考慮
                string[] tmpdir = {
                    Path.GetTempPath(),
                    JoinFilePath(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Temp") };

                foreach (string tmp in tmpdir)
                {
                    string workpath = JoinFilePath(tmp, filename);

                    try
                    {
                        // 排他オープンのため、まずは FireStream 生成
                        stream = new FileStream(workpath, FileMode.Create, FileAccess.Write, FileShare.None);
                        StreamWriter writer = new StreamWriter(stream, encoding);
                        return writer;
                    }
                    catch { /* NOP */ }
                }
                stream = null;
                return null;
            }
            #endregion

            #region パス操作
            /// <summary>
            /// パス末尾のディレクトリ区切り文字削除
            /// </summary>
            private string TrimDirectoryPath(string folder)
            {
                char[] delms = { '\\', '/' };
                return folder.TrimEnd(delms);
            }
            /// <summary>
            /// パスとファイル名からフルパスの作成
            /// </summary>
            private string JoinFilePath(string basefolder, string name)
            {
                return TrimDirectoryPath(basefolder) + @"\" + name;
            }
            #endregion
        }
        #endregion

        #region 端末名取得
        /// <summary>
        /// 端末名取得 (RDPの場合は接続元の端末名)
        /// </summary>
        public static string GetTerminalName()
        {
            string? name = null;

            // RDP接続元の端末名
            try
            {
                IntPtr WTS_CURRENT_SERVER_HANDLE = IntPtr.Zero;   // ((HANDLE)NULL)
                uint WTS_CURRENT_SESSION = uint.MaxValue; // ((DWORD)-1)
                uint WTSClientName = 10;            // enum WTS_INFO_CLASS.WTSClientName (10)
                IntPtr ppBuffer;
                uint iReturned;

                if (WTSQuerySessionInformation(WTS_CURRENT_SERVER_HANDLE,
                                               WTS_CURRENT_SESSION,
                                               WTSClientName,
                                               out ppBuffer,
                                               out iReturned))
                {
                    name = Marshal.PtrToStringAnsi(ppBuffer);

                    // メモリ開放
                    WTSFreeMemory(ppBuffer);
                }

            }
            catch { /* Dll ロードエラー等 NOP */ }

            if (string.IsNullOrEmpty(name))
            {
                // 自端末名
                name = Environment.MachineName;
            }
            return name;
        }

        [DllImport("wtsapi32.dll")]
        public static extern bool WTSQuerySessionInformation(IntPtr hServer, uint sessionId, uint wtsInfoClass, out IntPtr ppBuffer, out uint iBytesReturned);

        [DllImport("wtsapi32.dll")]
        private static extern void WTSFreeMemory(IntPtr pMemory);
        #endregion

    }
}
