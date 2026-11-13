// DocTabs — 轻量文档标签栏：把 Office 的 Word / Excel 窗口当作“一个窗口的多个标签”来切换（不含 WPS）。
// 分组：默认所有文档在同一组；把标签竖直拖出标签栏即拆成单独的一组（单独的窗口），拖到别的标签栏上松手则合并回去。
//      每一组有自己的标签栏，系统分屏时左右两边都能看到各自的标签。
// 最大化：默认标签栏固定在屏幕顶部，最大化的文档窗口整体下移让出标签栏高度（其他程序不受影响）。
// 可选“功能区下方”模式：标签栏悬挂在 Word / Excel 功能区底边，不移动文档窗口。
// 小窗口：标签栏贴在窗口标题栏正上方随窗口移动；窗口顶到屏幕上沿（如系统分屏）时往下让出标签栏的高度。
// 同组联动：一个文档最大化/还原/移动/最小化，其他文档同步为相同状态和位置；点窗口 × 或 Alt+F4 整组依次关闭。
// 拖标签栏空白处（或只有一个标签时拖标签）即拖动窗口，交给系统自己的拖动流程，与拖标题栏一样流畅，也支持贴边分屏。
// Office 插件：程序本身同时是一个 COM 加载项，Word / Excel 启动时自动拉起 DocTabs，Office 全部关闭后自动退出。
// 编译：build.bat（使用系统自带的 .NET Framework 4.x csc.exe，无需安装 SDK）
// https://github.com/stevensijd-qgg/DocTabs — MIT License
// 本仓库：https://github.com/cmm101010/DocTabs — 基于上游的 MIT 修改版

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: System.Reflection.AssemblyTitle("DocTabs")]
[assembly: System.Reflection.AssemblyProduct("DocTabs")]
[assembly: System.Reflection.AssemblyDescription("Lightweight WPS-style tab bar for Microsoft Word and Excel windows")]
[assembly: System.Reflection.AssemblyCopyright("Copyright (c) 2026 stevensijd-qgg; modifications (c) 2026 cmm101010, MIT License")]
[assembly: System.Reflection.AssemblyVersion("1.5.7.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.5.7.0")]
[assembly: ComVisible(false)]

namespace DocTabs
{
    static class Program
    {
        public const string MutexName = "DocTabs.SingleInstance.8F2C41";

        [STAThread]
        static void Main(string[] args)
        {
            bool created;
            using (var mutex = new Mutex(true, MutexName, out created))
            {
                if (!created) return;
                Log.Write("DocTabs " + Application.ProductVersion + " 启动" + (args.Length > 0 ? " " + string.Join(" ", args) : ""));
                // 拖动跟随时需要稳定的 8ms 节拍，默认 15.6ms 会多出一帧的滞后
                Native.timeBeginPeriod(1);
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                Application.ThreadException += (s, e) => Log.Write("出错（已忽略）：" + e.Exception);
                AppDomain.CurrentDomain.UnhandledException += (s, e) => Log.Write("崩溃：" + e.ExceptionObject);
                // 逐显示器 DPI 感知：主副屏缩放不同时，取到的窗口坐标才是真实像素，不会被系统按主屏缩放换算
                bool perMonitor = false;
                try { perMonitor = Native.SetProcessDpiAwarenessContext(new IntPtr(-4) /* PER_MONITOR_AWARE_V2 */); }
                catch (EntryPointNotFoundException) { }
                if (!perMonitor) Native.SetProcessDPIAware();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                // 由 Office 插件拉起时带 /office：Word / Excel 全部关闭后自动退出
                bool byOffice = Array.IndexOf(args, "/office") >= 0;
                var app = new Manager(byOffice);
                Application.Run();
                GC.KeepAlive(app);
            }
        }
    }

    // ---------------- Office 插件：Word / Excel 启动时把 DocTabs 一起拉起来 ----------------

    // custom 参数必须按 VARIANT 的 SAFEARRAY 封送，否则 Office 调用 OnConnection 时 CLR 会直接崩溃
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, [In] int connectMode,
                                      [In, MarshalAs(UnmanagedType.IDispatch)] object addInInst,
                                      [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection([In] int removeMode,
                                         [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    // 运行在 Word / Excel 进程里：保证 DocTabs 在运行，并把文档的未保存状态上报给主程序
    [ComVisible(true), Guid(OfficeAddin.Clsid), ProgId(OfficeAddin.ProgId), ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2
    {
        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            Log.Write("Office 插件已加载");
            OfficeAddin.LaunchDocTabs();
            OfficeAddin.StartWatch();
            OfficeAddin.StartDirtyReport(application);
        }
        public void OnDisconnection(int removeMode, ref Array custom)
        {
            OfficeAddin.StopDirtyReport();
            OfficeAddin.StopWatch();
        }
        public void OnAddInsUpdate(ref Array custom) { }
        // Office 完全启动好之后再确认一次（开机时 Office 可能先于桌面就绪启动）
        public void OnStartupComplete(ref Array custom) { OfficeAddin.LaunchDocTabs(); }
        public void OnBeginShutdown(ref Array custom) { OfficeAddin.StopDirtyReport(); }
    }

    static class OfficeAddin
    {
        public const string Clsid = "5EBC09F9-99B3-4194-BA3C-1A0AC9DDD7F9";
        public const string ProgId = "DocTabs.Connect";
        public const string SettingsKey = @"Software\DocTabs";
        // 用户从托盘主动退出时写入；DocTabs 下次启动时清除
        public const string UserExitedValue = "UserExited";
        // Office 插件运行在 Word / Excel 进程内，直接把 Document.Saved 写到这里。
        // 主程序优先读它，绕开在某些 Office 安装上会返回 E_FAIL 的 OBJID_NATIVEOM。
        const string DirtyKey = @"Software\DocTabs\Dirty";
        const string ClsidKey = @"Software\Classes\CLSID\{" + Clsid + "}";
        static readonly string[] Apps = { "Word", "Excel" };

        static System.Threading.Timer watch;
        static int watchLaunches;
        static DateTime lastLaunch;
        static readonly object dirtySync = new object();
        static object dirtyApp;
        static System.Windows.Forms.Timer dirtyTimer;
        static Dictionary<string, string> dirtyLast = new Dictionary<string, string>();
        static bool dirtyBusy;
        static int dirtyErrorTick;

        public static void LaunchDocTabs()
        {
            try { if (!IsRunning()) Start("启动 DocTabs"); }
            catch (Exception ex) { Log.Write("启动 DocTabs 失败：" + ex.Message); } // 任何问题都不能影响 Office 本身
        }

        // Office 一直开着、而 DocTabs 中途没了（被结束、崩溃）时，几秒内把它拉回来；用户从托盘主动退出的不拉
        public static void StartWatch()
        {
            try { if (watch == null) watch = new System.Threading.Timer(delegate { Watch(); }, null, 5000, 3000); }
            catch { }
        }

        public static void StopWatch()
        {
            try { if (watch != null) { watch.Dispose(); watch = null; } }
            catch { }
        }

        // 在 Office 自己的进程里定时读 Saved。跨进程 COM 自动化在部分 Office 安装上不可用，
        // 但插件本来就在文档进程里，所以这条路不受影响。
        public static void StartDirtyReport(object application)
        {
            try
            {
                if (application == null) return;
                lock (dirtySync)
                {
                    dirtyApp = application;
                    if (dirtyTimer != null) return;
                    // OnConnection 跑在 Office 的 UI 线程上，WinForms Timer 的 Tick 也会回到这条线程。
                    // 不碰后台线程，Word 的对象模型就不会遇到跨套间调用问题。
                    dirtyLast.Clear();
                    dirtyTimer = new System.Windows.Forms.Timer();
                    dirtyTimer.Interval = 500;
                    dirtyTimer.Tick += delegate { DirtyTick(); };
                    dirtyTimer.Start();
                }
            }
            catch (Exception ex) { Log.Write("未保存上报启动失败：" + ex.Message); }
        }

        public static void StopDirtyReport()
        {
            try
            {
                lock (dirtySync)
                {
                    if (dirtyTimer != null)
                    {
                        dirtyTimer.Stop();
                        dirtyTimer.Dispose();
                        dirtyTimer = null;
                    }
                    dirtyApp = null;
                }
                RemoveOwnDirtyValues();
            }
            catch { }
        }

        static void DirtyTick()
        {
            if (dirtyBusy) return;
            dirtyBusy = true;
            try
            {
                PublishDirty(dirtyLast);
                dirtyErrorTick = 0;
            }
            catch (Exception ex)
            {
                if (unchecked(Environment.TickCount - dirtyErrorTick) > 60000)
                {
                    dirtyErrorTick = Environment.TickCount;
                    Log.Write("未保存上报失败：" + ex.GetType().Name + ":" + ex.Message);
                }
            }
            finally { dirtyBusy = false; }
        }

        static void PublishDirty(Dictionary<string, string> last)
        {
            object appObj = dirtyApp;
            if (appObj == null) return;
            string appName = null;
            try { dynamic app = appObj; appName = (string)app.Name; } catch { }
            bool excel = appName != null && appName.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) >= 0;
            string appLabel = excel ? "Excel" : "Word";
            var found = new List<KeyValuePair<string, bool>>();
            if (excel) CollectExcel(appObj, found); else CollectWord(appObj, found);

            int pid = System.Diagnostics.Process.GetCurrentProcess().Id;
            using (var key = Registry.CurrentUser.CreateSubKey(DirtyKey))
            {
                if (key == null) return;
                var seen = new HashSet<string>();
                foreach (var pair in found)
                {
                    string value = pid + "|" + (pair.Value ? "1" : "0") + "|" + (excel ? "E" : "W");
                    string old;
                    bool changed = !last.TryGetValue(pair.Key, out old) || old != value;
                    if (changed)
                    {
                        key.SetValue(pair.Key, value, RegistryValueKind.String);
                        last[pair.Key] = value;
                        Log.Write(string.Format("未保存[{0}] 0x{1}={2}", appLabel, pair.Key, pair.Value ? "未保存" : "已保存"));
                    }
                    seen.Add(pair.Key);
                }
                // 关掉的窗口立刻清掉；进程异常退出留下的值由主程序按 PID/窗口有效性过滤。
                foreach (string name in key.GetValueNames())
                {
                    if (seen.Contains(name)) continue;
                    string value = key.GetValue(name) as string;
                    if (value != null && value.StartsWith(pid + "|", StringComparison.Ordinal))
                    {
                        key.DeleteValue(name, false);
                        last.Remove(name);
                    }
                }
            }
        }

        static void CollectWord(object appObj, List<KeyValuePair<string, bool>> found)
        {
            dynamic app = appObj;
            dynamic wins = null;
            try
            {
                wins = app.Windows;
                int n = (int)wins.Count;
                for (int i = 1; i <= n; i++)
                {
                    dynamic win = null;
                    object doc = null;
                    try
                    {
                        win = wins.Item(i);
                        int hwnd = (int)win.Hwnd;
                        doc = win.Document;
                        bool saved = (bool)((dynamic)doc).Saved;
                        found.Add(new KeyValuePair<string, bool>(hwnd.ToString("X8"), !saved));
                    }
                    catch { }
                    finally { ReleaseCom(doc); ReleaseCom(win); }
                }
            }
            finally { ReleaseCom(wins); }
        }

        static void CollectExcel(object appObj, List<KeyValuePair<string, bool>> found)
        {
            dynamic app = appObj;
            dynamic books = null;
            try
            {
                books = app.Workbooks;
                int n = (int)books.Count;
                for (int i = 1; i <= n; i++)
                {
                    dynamic book = null;
                    dynamic wins = null;
                    try
                    {
                        book = books.Item(i);
                        bool saved = (bool)book.Saved;
                        wins = book.Windows;
                        int wn = (int)wins.Count;
                        for (int j = 1; j <= wn; j++)
                        {
                            dynamic win = null;
                            try
                            {
                                win = wins.Item(j);
                                int hwnd = (int)win.Hwnd;
                                found.Add(new KeyValuePair<string, bool>(hwnd.ToString("X8"), !saved));
                            }
                            catch { }
                            finally { ReleaseCom(win); }
                        }
                    }
                    catch { }
                    finally { ReleaseCom(wins); ReleaseCom(book); }
                }
            }
            finally { ReleaseCom(books); }
        }

        static void RemoveOwnDirtyValues()
        {
            try
            {
                int pid = System.Diagnostics.Process.GetCurrentProcess().Id;
                using (var key = Registry.CurrentUser.OpenSubKey(DirtyKey, true))
                {
                    if (key == null) return;
                    foreach (string name in key.GetValueNames())
                    {
                        string value = key.GetValue(name) as string;
                        if (value != null && value.StartsWith(pid + "|", StringComparison.Ordinal))
                            key.DeleteValue(name, false);
                    }
                }
            }
            catch { }
        }

        static void ReleaseCom(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                try { Marshal.ReleaseComObject(obj); }
                catch { }
            }
        }

        static void Watch()
        {
            try
            {
                if (IsRunning()) return;
                // 连着拉起 5 次都没留住（比如开机时系统还没就绪）就隔一分钟再试，不放弃
                if (watchLaunches >= 5 && (DateTime.Now - lastLaunch).TotalSeconds < 60) return;
                using (var k = Registry.CurrentUser.OpenSubKey(SettingsKey))
                    if (k != null && Equals(k.GetValue(UserExitedValue), 1)) return;
                watchLaunches++;
                Start("DocTabs 不在运行，重新启动（第 " + watchLaunches + " 次）");
            }
            catch (Exception ex) { Log.Write("重新启动 DocTabs 失败：" + ex.Message); }
        }

        static bool IsRunning()
        {
            Mutex m;
            if (!Mutex.TryOpenExisting(Program.MutexName, out m)) return false;
            m.Dispose();
            return true;
        }

        static void Start(string why)
        {
            lastLaunch = DateTime.Now;
            Log.Write(why);
            string exe = typeof(Connect).Assembly.Location;
            var psi = new System.Diagnostics.ProcessStartInfo(exe, "/office")
            {
                UseShellExecute = false,
                WorkingDirectory = System.IO.Path.GetDirectoryName(exe)
            };
            System.Diagnostics.Process.Start(psi);
        }

        static string CodeBase { get { return typeof(Connect).Assembly.CodeBase; } }
        static string AssemblyName { get { return typeof(Connect).Assembly.FullName; } }

        static string AddinKey(string app) { return @"Software\Microsoft\Office\" + app + @"\Addins\" + ProgId; }

        public static bool IsRegistered()
        {
            foreach (var app in Apps)
                using (var k = Registry.CurrentUser.OpenSubKey(AddinKey(app)))
                    if (k == null) return false;
            return true;
        }

        // Office 2016 及以后是 16.0，2013 是 15.0
        static IEnumerable<string> OfficeVersions()
        {
            yield return "16.0";
            using (var k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\15.0"))
                if (k != null) yield return "15.0";
        }

        static string ResiliencyKey(string ver, string app) { return @"Software\Microsoft\Office\" + ver + @"\" + app + @"\Resiliency"; }

        // 程序换了位置或升级了版本，或者 Office 把插件停用了（加载出错、嫌它拖慢启动），注册信息都要重写
        public static bool IsCurrent()
        {
            using (var k = Registry.CurrentUser.OpenSubKey(ClsidKey + @"\InprocServer32"))
                if (k == null || !Equals(k.GetValue("CodeBase"), CodeBase) || !Equals(k.GetValue("Assembly"), AssemblyName)) return false;
            foreach (var app in Apps)
            {
                using (var k = Registry.CurrentUser.OpenSubKey(AddinKey(app)))
                    if (k == null || !Equals(k.GetValue("LoadBehavior"), 3)) return false;
                using (var k = Registry.CurrentUser.OpenSubKey(ResiliencyKey("16.0", app) + @"\DoNotDisableAddinList"))
                    if (k == null || k.GetValue(ProgId) == null) return false;
            }
            return !ClearDisabledItems(false);
        }

        // Office 因出错而停用插件时，会在 DisabledItems 里记一条带程序路径的二进制记录。remove 为 false 时只检查
        static bool ClearDisabledItems(bool remove)
        {
            bool found = false;
            foreach (var ver in OfficeVersions())
                foreach (var app in Apps)
                    using (var k = Registry.CurrentUser.OpenSubKey(ResiliencyKey(ver, app) + @"\DisabledItems", remove))
                    {
                        if (k == null) continue;
                        foreach (var name in k.GetValueNames())
                        {
                            var data = k.GetValue(name) as byte[];
                            if (data == null || Encoding.Unicode.GetString(data).IndexOf("doctabs", StringComparison.OrdinalIgnoreCase) < 0) continue;
                            found = true;
                            if (remove) k.DeleteValue(name, false);
                        }
                    }
            return found;
        }

        // 只写当前用户的注册表，不需要管理员权限；32 / 64 位 Office 都登记
        public static void Register()
        {
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                {
                    using (var k = root.CreateSubKey(@"Software\Classes\" + ProgId))
                    {
                        k.SetValue("", ProgId);
                        using (var c = k.CreateSubKey("CLSID")) c.SetValue("", "{" + Clsid + "}");
                    }
                    using (var k = root.CreateSubKey(ClsidKey))
                    {
                        k.SetValue("", ProgId);
                        using (var p = k.CreateSubKey("ProgId")) p.SetValue("", ProgId);
                        using (var srv = k.CreateSubKey("InprocServer32"))
                        {
                            srv.SetValue("", "mscoree.dll");
                            srv.SetValue("ThreadingModel", "Both");
                            srv.SetValue("Class", typeof(Connect).FullName);
                            srv.SetValue("Assembly", AssemblyName);
                            srv.SetValue("RuntimeVersion", "v4.0.30319");
                            srv.SetValue("CodeBase", CodeBase);
                        }
                    }
                }
            foreach (var app in Apps)
                using (var k = Registry.CurrentUser.CreateSubKey(AddinKey(app)))
                {
                    k.SetValue("FriendlyName", "DocTabs 文档标签栏");
                    k.SetValue("Description", "打开 Word / Excel 时自动启动 DocTabs 标签栏");
                    k.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                }
            // 开机后第一次打开 Office 时 .NET 冷启动较慢，Office 可能把插件当成“拖慢启动”自动停用：列入“始终启用”
            foreach (var ver in OfficeVersions())
                foreach (var app in Apps)
                    using (var k = Registry.CurrentUser.CreateSubKey(ResiliencyKey(ver, app) + @"\DoNotDisableAddinList"))
                        k.SetValue(ProgId, 1, RegistryValueKind.DWord);
            ClearDisabledItems(true);
        }

        public static void Unregister()
        {
            foreach (var app in Apps) Registry.CurrentUser.DeleteSubKeyTree(AddinKey(app), false);
            foreach (var ver in OfficeVersions())
                foreach (var app in Apps)
                    using (var k = Registry.CurrentUser.OpenSubKey(ResiliencyKey(ver, app) + @"\DoNotDisableAddinList", true))
                        if (k != null) k.DeleteValue(ProgId, false);
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                {
                    root.DeleteSubKeyTree(@"Software\Classes\" + ProgId, false);
                    root.DeleteSubKeyTree(ClsidKey, false);
                }
        }
    }

    // 运行记录，排查“没有自动启动”这类问题用：%LOCALAPPDATA%\DocTabs\DocTabs.log，超过 64 KB 清空重记
    static class Log
    {
        public static void Write(string text)
        {
            try
            {
                string dir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DocTabs");
                System.IO.Directory.CreateDirectory(dir);
                string file = System.IO.Path.Combine(dir, "DocTabs.log");
                var fi = new System.IO.FileInfo(file);
                if (fi.Exists && fi.Length > 64 * 1024) fi.Delete();
                var p = System.Diagnostics.Process.GetCurrentProcess();
                System.IO.File.AppendAllText(file, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + p.ProcessName + "(" + p.Id + ")  " + text + "\r\n", Encoding.UTF8);
            }
            catch { }
        }
    }

    enum DocKind { Word, Excel }

    enum BarMode { Hidden, Maximized, Follow, AutoHide }

    class DocTab
    {
        public IntPtr Handle;
        public DocKind Kind;
        public string Title;
        public string FullTitle;
        public RectangleF Bounds;
        public RectangleF CloseBounds;
        public bool Dirty;
    }

    // ---------------- 全局：事件钩子、窗口扫描、分组、托盘 ----------------

    class Manager
    {
        const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        const string RunName = "DocTabs";

        public readonly List<TabBar> Groups = new List<TabBar>();
        // 最近一次有文档到前台的分组；新打开的文档加入这一组
        public TabBar Active;
        // 最大化后被我们下移过的窗口 / 伪最大化的窗口 / 不允许移动最大化状态的窗口
        public readonly HashSet<IntPtr> Shifted = new HashSet<IntPtr>();
        // 打开左侧导航/编辑器窗格时，为把标签栏移到窗口上方而临时下移过的窗口
        public readonly HashSet<IntPtr> PaneShifted = new HashSet<IntPtr>();
        public readonly HashSet<IntPtr> PseudoMax = new HashSet<IntPtr>();
        // 程序自己移动/最大化窗口时产生的 LOCATIONCHANGE 不能当成用户操作
        readonly Dictionary<IntPtr, int> selfChange = new Dictionary<IntPtr, int>();
        readonly Dictionary<IntPtr, bool> selfChangeZoom = new Dictionary<IntPtr, bool>();
        readonly HashSet<IntPtr> noShift = new HashSet<IntPtr>();
        // 拖动/缩放某组的当前文档时，暂时挪到屏幕外的同组其他文档，及其原来的位置。
        // 不用隐藏：Excel 处理隐藏/显示一个窗口要几百毫秒，期间拖动会卡住；挪到屏幕外只要几毫秒
        public readonly HashSet<IntPtr> Parked = new HashSet<IntPtr>();
        readonly Dictionary<IntPtr, Native.RECT> parkedFrom = new Dictionary<IntPtr, Native.RECT>();
        public TabBar DragGroup;
        DateTime dragStart;
        // 从标签栏开始的拖动：被拖的窗口、发起拖动的标签栏、松手后要切到前台的窗口
        IntPtr moveWin, activateAfterMove;
        TabBar dragSource;
        Point dragOffset, lastTarget, posAtSend;
        DateTime lastSend;
        readonly System.Windows.Forms.Timer moveWatch;
        readonly System.Windows.Forms.Timer trackTimer;

        // 用户点了某个文档窗口的 ×（或按了 Alt+F4）：该窗口真正关掉后，它所在的整组一起关闭
        IntPtr closeIntent;
        DateTime closeIntentTime;
        // 正在逐个关闭的文档队列
        readonly List<IntPtr> cascade = new List<IntPtr>();
        IntPtr cascadeCur;
        DateTime cascadeStepTime;

        readonly List<IntPtr> hooks = new List<IntPtr>();
        readonly Native.WinEventProc winEventProc;
        readonly Native.EnumWindowsProc enumProc;
        readonly Native.LowLevelProc mouseProc;
        readonly Native.LowLevelProc keyProc;
        IntPtr mouseHook, keyHook;
        readonly Dictionary<uint, string> procNames = new Dictionary<uint, string>();
        List<DocTab> scanResult;
        // 后台读取 Word / Excel 的“文档已修改未保存”状态
        DirtyWatcher dirtyWatcher;
        int dirtyApplied = -1;

        readonly System.Windows.Forms.Timer scanTimer;
        readonly System.Windows.Forms.Timer fallbackTimer;
        readonly System.Windows.Forms.Timer revealTimer;
        readonly NotifyIcon tray;
        readonly MenuItem autostartItem;
        readonly MenuItem addinItem;
        readonly bool byOffice;
        int officeGoneTicks;
        bool exiting;
        public static bool BelowRibbon;
        // 第三种布局：把标签栏画进 Word 自己的标题栏里
        public static bool TitleBarTabs;
        // 标签栏正下方的正文画布颜色；功能区下方模式用它做标签栏底色
        public static Color CanvasColor = Color.FromArgb(102, 102, 102);
        public static bool CanvasLight;
        int canvasStamp;
        readonly MenuItem placementItem;
        readonly MenuItem titleItem;

        public Manager(bool byOffice)
        {
            this.byOffice = byOffice;
            ClearUserExited();
            dirtyWatcher = new DirtyWatcher();

            scanTimer = new System.Windows.Forms.Timer { Interval = 120 };
            scanTimer.Tick += delegate { scanTimer.Stop(); Scan(); };
            fallbackTimer = new System.Windows.Forms.Timer { Interval = 2000 };
            fallbackTimer.Tick += delegate
            {
                // 兜底：漏掉了“拖动结束”通知时，别让文档一直处于隐藏状态
                if (DragGroup != null && moveWin == IntPtr.Zero && (DateTime.Now - dragStart).TotalSeconds > 30) EndWindowDrag();
                if (DragGroup == null && Parked.Count > 0 && !revealTimer.Enabled) RevealParked();
                Scan();
                if (this.byOffice) CheckOfficeGone();
            };
            fallbackTimer.Start();
            // 松手后稍等当前文档的最终位置稳定，再把其他文档放到同一位置并显示出来
            revealTimer = new System.Windows.Forms.Timer { Interval = 25 };
            revealTimer.Tick += delegate { revealTimer.Stop(); RevealParked(); };
            moveWatch = new System.Windows.Forms.Timer { Interval = 10 };
            moveWatch.Tick += delegate { MoveWatchTick(); };
            // 拖动窗口时按帧读窗口位置来摆放标签栏：只靠系统的位置变更事件会慢一拍
            trackTimer = new System.Windows.Forms.Timer { Interval = 8 };
            trackTimer.Tick += delegate
            {
                var g = DragGroup;
                if (g == null) { trackTimer.Stop(); return; }
                g.UpdatePlacement();
            };

            enumProc = EnumCallback;
            winEventProc = OnWinEvent;
            const uint flags = Native.WINEVENT_OUTOFCONTEXT | Native.WINEVENT_SKIPOWNPROCESS;
            hooks.Add(Native.SetWinEventHook(Native.EVENT_SYSTEM_FOREGROUND, Native.EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, winEventProc, 0, 0, flags));
            hooks.Add(Native.SetWinEventHook(Native.EVENT_SYSTEM_MOVESIZESTART, Native.EVENT_SYSTEM_MOVESIZEEND, IntPtr.Zero, winEventProc, 0, 0, flags));
            hooks.Add(Native.SetWinEventHook(Native.EVENT_SYSTEM_MINIMIZESTART, Native.EVENT_SYSTEM_MINIMIZEEND, IntPtr.Zero, winEventProc, 0, 0, flags));
            hooks.Add(Native.SetWinEventHook(Native.EVENT_OBJECT_DESTROY, Native.EVENT_OBJECT_HIDE, IntPtr.Zero, winEventProc, 0, 0, flags));
            hooks.Add(Native.SetWinEventHook(Native.EVENT_OBJECT_LOCATIONCHANGE, Native.EVENT_OBJECT_NAMECHANGE, IntPtr.Zero, winEventProc, 0, 0, flags));
            // Office 新打开的窗口加载期间是隐身（cloaked）的，显示出来时要立刻收进来
            hooks.Add(Native.SetWinEventHook(Native.EVENT_OBJECT_CLOAKED, Native.EVENT_OBJECT_UNCLOAKED, IntPtr.Zero, winEventProc, 0, 0, flags));

            tray = new NotifyIcon { Icon = MakeTrayIcon(Native.MonitorScale(IntPtr.Zero)), Text = "DocTabs 文档标签栏", Visible = true };
            autostartItem = new MenuItem("开机自启动", delegate { ToggleAutostart(); }) { Checked = IsAutostart() };
            addinItem = new MenuItem("打开 Word / Excel 时自动启动", delegate { ToggleAddin(); });
            BelowRibbon = ReadBelowRibbon();
            TitleBarTabs = ReadTitleBarTabs();
            // 旧版本允许两个开关同时为 1；布局代码优先采用标题栏模式，这里把冲突状态收敛掉。
            if (BelowRibbon && TitleBarTabs)
            {
                BelowRibbon = false;
                SaveLayoutFlags();
            }
            placementItem = new MenuItem("标签栏放在功能区下方", delegate { ToggleBelowRibbon(); }) { Checked = BelowRibbon, RadioCheck = true };
            titleItem = new MenuItem("标签栏内嵌标题栏（预览）", delegate { ToggleTitleBar(); }) { Checked = TitleBarTabs, RadioCheck = true };
            tray.ContextMenu = new ContextMenu(new[]
            {
                new MenuItem("DocTabs 文档标签栏") { Enabled = false },
                new MenuItem("-"),
                addinItem,
                autostartItem,
                placementItem,
                titleItem,
                new MenuItem("退出", delegate { ExitApp(true); })
            });
            InitAddin();

            SystemEvents.DisplaySettingsChanged += OnDisplayChanged;

            // 只记录“在当前文档的关闭按钮上松开左键”和“Alt+F4”这两个动作，回调里不做别的事
            mouseProc = OnMouseLL;
            keyProc = OnKeyLL;
            IntPtr mod = Native.GetModuleHandle(null);
            mouseHook = Native.SetWindowsHookEx(Native.WH_MOUSE_LL, mouseProc, mod, 0);
            keyHook = Native.SetWindowsHookEx(Native.WH_KEYBOARD_LL, keyProc, mod, 0);
            Scan();
        }

        // ---------------- 分组 ----------------

        // 记录“这次最大化/还原是我们自己下发的”，并记住下发的目标状态。
        // 只看时间会误吞用户在同一窗口上的真实操作，必须连状态一起核对。
        public void MarkSelfChange(IntPtr h, bool targetZoom)
        {
            selfChange[h] = unchecked(Environment.TickCount + 400);
            selfChangeZoom[h] = targetZoom;
        }

        public bool IsSelfChange(IntPtr h)
        {
            int until;
            if (!selfChange.TryGetValue(h, out until)) return false;
            bool target;
            if (!selfChangeZoom.TryGetValue(h, out target)) { selfChange.Remove(h); return false; }
            if (unchecked(until - Environment.TickCount) > 0)
            {
                // 状态仍是我们下发的那个：确认是自己造成的
                if (Native.IsZoomed(h) == target) return true;
                // 还没到窗口期但状态已经变了：用户自己点了最大化/还原，不能吞
                selfChange.Remove(h);
                selfChangeZoom.Remove(h);
                return false;
            }
            selfChange.Remove(h);
            selfChangeZoom.Remove(h);
            return false;
        }

        public TabBar GroupOf(IntPtr h)
        {
            if (h == IntPtr.Zero) return null;
            foreach (var g in Groups) if (g.IndexOf(h) >= 0) return g;
            return null;
        }

        public TabBar NewGroup()
        {
            var g = new TabBar(this);
            Groups.Add(g);
            return g;
        }

        void RemoveGroup(TabBar g)
        {
            Groups.Remove(g);
            if (Active == g) Active = null;
            if (DragGroup == g) DragGroup = null;
            g.Destroy();
        }

        bool IsOnAnyBar(int x, int y)
        {
            foreach (var g in Groups) if (g.Visible && g.Bounds.Contains(x, y)) return true;
            return false;
        }

        public void UpdateAll()
        {
            foreach (var g in Groups.ToArray()) g.UpdatePlacement();
        }

        // 后台线程查到新的未保存状态后，在 UI 线程刷新各个标签
        public void ApplyDirtyIfNeeded()
        {
            if (dirtyWatcher == null) return;
            int v = dirtyWatcher.Version;
            if (v == dirtyApplied) return;
            dirtyApplied = v;
            var map = dirtyWatcher.Snapshot();
            foreach (var g in Groups.ToArray())
            {
                bool changed = false;
                foreach (var t in g.Tabs)
                {
                    bool d;
                    if (!map.TryGetValue(t.Handle, out d)) d = false;
                    if (t.Dirty != d) { t.Dirty = d; changed = true; }
                }
                if (changed) g.Invalidate();
            }
        }

        // 把 src 组的文档全部并入 dst 组，src 的当前文档成为 dst 的当前文档
        public void Merge(TabBar src, TabBar dst)
        {
            IntPtr cur = src.CurDoc;
            var moved = new List<DocTab>(src.Tabs);
            src.Tabs.Clear();
            RemoveGroup(src);
            foreach (var t in moved) dst.Tabs.Add(t);
            Active = dst;
            foreach (var t in moved)
                if (t.Handle != cur && !Parked.Contains(t.Handle)) dst.ApplyGroupState(t.Handle, false);
            if (cur == IntPtr.Zero && moved.Count > 0) cur = moved[0].Handle;
            dst.LayoutTabs();
            if (cur != IntPtr.Zero) dst.ActivateDoc(cur);
            UpdateAll();
        }

        // 拖动结束时鼠标落在另一组的标签上：并进那一组。拖到屏幕最上沿是系统的“最大化”手势，不算
        void TryMergeOnDrop(TabBar src, Point c)
        {
            if (src == null || !Groups.Contains(src)) return;
            Rectangle mon = Screen.FromPoint(c).Bounds;
            if (c.Y <= mon.Top + 2) return;
            foreach (var g in Groups)
                if (g != src && g.Visible && g.TabStripContains(c)) { Merge(src, g); return; }
        }

        // ---------------- 从标签栏拖动窗口 ----------------
        // 不交给系统的拖动流程：按下发生在标签栏上，Office 的拖动流程常常要等松手才开始。
        // 鼠标一直由按下的那个标签栏捕获，每来一条鼠标移动就把窗口挪过去（窗口用异步方式挪，
        // 上一次还没生效就先不发，免得 Office 忙时请求排队、之后按鼠标轨迹“补跑”）。
        // 按着鼠标时不能把窗口切到前台（Office 会等到松手才处理），只把它提到最上面，松手后再切

        public TabBar DragSource { get { return moveWin != IntPtr.Zero ? dragSource : null; } }

        public void BeginDrag(IntPtr h, TabBar source)
        {
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            BeginDrag(h, source, new Point(r.Left, r.Top));
        }

        // winPos：窗口左上角此刻应在的位置（刚异步下发的位置可能还没生效，不能现读）
        public void BeginDrag(IntPtr h, TabBar source, Point winPos)
        {
            Point c = Cursor.Position;
            if (Native.GetForegroundWindow() != h)
            {
                Native.SetWindowPos(h, Native.HWND_TOP, 0, 0, 0, 0,
                    Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE | Native.SWP_ASYNCWINDOWPOS);
                activateAfterMove = h;
            }
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            dragOffset = new Point(winPos.X - c.X, winPos.Y - c.Y);
            lastTarget = winPos;
            posAtSend = new Point(r.Left, r.Top);
            lastSend = DateTime.MinValue;
            moveWin = h;
            dragSource = source;
            var g = GroupOf(h);
            if (g != null && g.CurDoc == h && DragGroup != g) BeginWindowDrag(g);
            moveWatch.Start();
        }

        // 跟着鼠标走一步（标签栏收到鼠标移动时调用；计时器兜底）
        public void DragMove()
        {
            IntPtr h = moveWin;
            if (h == IntPtr.Zero) return;
            Point c = Cursor.Position;
            var p = new Point(c.X + dragOffset.X, c.Y + dragOffset.Y);
            if (p == lastTarget) return;
            // 生效 = 到了上次要的位置，或者至少动过了（跨到缩放不同的屏幕时 Office 会自己微调位置）
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            var now = new Point(r.Left, r.Top);
            if (now != lastTarget && now == posAtSend && (DateTime.Now - lastSend).TotalMilliseconds < 250) return;
            posAtSend = now;
            lastTarget = p;
            lastSend = DateTime.Now;
            MoveWindowTo(h, c);
        }

        public void EndDrag()
        {
            IntPtr h = moveWin;
            if (h == IntPtr.Zero) return;
            moveWin = IntPtr.Zero;
            dragSource = null;
            moveWatch.Stop();
            Point c = Cursor.Position;
            if (Native.IsWindow(h) && !SnapOnDrop(h, c)) MoveWindowTo(h, c);
            if (DragGroup != null) EndWindowDrag();
            else TryMergeOnDrop(GroupOf(h), c);
            ActivateAfterMove();
        }

        // 计时器兜底：标签栏丢了鼠标捕获（比如被隐藏）时照样跟着走，并发现松手
        void MoveWatchTick()
        {
            IntPtr h = moveWin;
            if (h == IntPtr.Zero) { moveWatch.Stop(); return; }
            int vk = SystemInformation.MouseButtonsSwapped ? Native.VK_RBUTTON : Native.VK_LBUTTON;
            if (!Native.IsWindow(h) || (Native.GetAsyncKeyState(vk) & 0x8000) == 0) { EndDrag(); return; }
            DragMove();
        }

        // 把正在拖的窗口放到鼠标 c 处（保持按下时抓住的位置）
        void MoveWindowTo(IntPtr h, Point c)
        {
            Native.SetWindowPos(h, IntPtr.Zero, c.X + dragOffset.X, c.Y + dragOffset.Y, 0, 0,
                Native.SWP_NOSIZE | Native.SWP_NOZORDER | Native.SWP_NOACTIVATE | Native.SWP_ASYNCWINDOWPOS);
        }

        // 松手时鼠标顶到屏幕外沿：与系统贴边一样，上沿最大化、左右沿各占半屏（两块屏幕相接的那条边不算）
        bool SnapOnDrop(IntPtr h, Point c)
        {
            Screen scr = Screen.FromPoint(c);
            Rectangle mon = scr.Bounds, wa = scr.WorkingArea;
            if (c.Y <= mon.Top + 1 && !OnAnyScreen(c.X, c.Y - 4))
            {
                Native.ShowWindowAsync(h, Native.SW_MAXIMIZE);
                return true;
            }
            bool left = c.X <= mon.Left + 1 && !OnAnyScreen(c.X - 4, c.Y);
            bool right = c.X >= mon.Right - 2 && !OnAnyScreen(c.X + 4, c.Y);
            if (!left && !right) return false;
            int half = wa.Width / 2;
            Rectangle want = left ? new Rectangle(wa.Left, wa.Top, half, wa.Height) : new Rectangle(wa.Left + half, wa.Top, wa.Width - half, wa.Height);
            // 加上程序自身的隐形边框，让看得见的边框正好贴边（之后顶到上沿的窗口会自动往下让出标签栏）
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            Native.RECT f = Native.FrameBounds(h, r);
            int l = want.Left - (f.Left - r.Left), t = want.Top - (f.Top - r.Top);
            int rr = want.Right + (r.Right - f.Right), b = want.Bottom + (r.Bottom - f.Bottom);
            Native.SetWindowPos(h, IntPtr.Zero, l, t, rr - l, b - t,
                Native.SWP_NOZORDER | Native.SWP_NOACTIVATE | Native.SWP_ASYNCWINDOWPOS);
            return true;
        }

        static bool OnAnyScreen(int x, int y)
        {
            foreach (var s in Screen.AllScreens) if (s.Bounds.Contains(x, y)) return true;
            return false;
        }

        // 拖完了：把被拖的窗口切到前台
        void ActivateAfterMove()
        {
            IntPtr h = activateAfterMove;
            activateAfterMove = IntPtr.Zero;
            if (h == IntPtr.Zero || !Native.IsWindow(h) || Native.IsIconic(h) || Native.GetForegroundWindow() == h) return;
            BringToFront(h);
        }

        // 把窗口切到前台（绕过系统的前台锁）；返回是否真的切成功
        public static bool BringToFront(IntPtr h, bool prepare = true)
        {
            if (Native.GetForegroundWindow() == h) return true;
            // 先在后台把标题栏按“激活”状态画好、提到顶层但不激活，再切前台。
            // 这样 Word 的文件名/标题栏不会在切过去的一瞬间整条重绘。
            if (prepare) Native.PrepareActivation(h);
            Native.SetForegroundWindow(h);
            if (Native.GetForegroundWindow() == h) return true;
            // 直接设置被前台锁挡住时，先挂到前台线程上再激活；比模拟输入更安静，不会闪任务栏
            uint dummy;
            uint fgThread = Native.GetWindowThreadProcessId(Native.GetForegroundWindow(), out dummy);
            uint cur = Native.GetCurrentThreadId();
            bool attached = fgThread != 0 && fgThread != cur && Native.AttachThreadInput(cur, fgThread, true);
            try
            {
                Native.BringWindowToTop(h);
                Native.SetForegroundWindow(h);
            }
            finally { if (attached) Native.AttachThreadInput(cur, fgThread, false); }
            if (Native.GetForegroundWindow() == h) return true;
            // 仍然不行才发送一个空鼠标输入解锁（PowerToys 同款做法）
            var input = new Native.INPUT { type = Native.INPUT_MOUSE };
            Native.SendInput(1, new[] { input }, Marshal.SizeOf(typeof(Native.INPUT)));
            Native.SetForegroundWindow(h);
            if (Native.GetForegroundWindow() == h) return true;
            // 最后兜底：Alt+Tab 使用的切换接口
            Native.SwitchToThisWindow(h, true);
            return Native.GetForegroundWindow() == h;
        }

        // ---------------- 关闭整组 ----------------

        IntPtr OnMouseLL(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && wParam.ToInt32() == Native.WM_LBUTTONUP)
            {
                var info = (Native.MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(Native.MSLLHOOKSTRUCT));
                IntPtr fg = Native.GetForegroundWindow();
                var g = GroupOf(fg);
                if (g != null && fg == g.CurDoc && !IsOnAnyBar(info.pt.X, info.pt.Y) && IsOverCloseButton(fg, info.pt.X, info.pt.Y))
                    SetCloseIntent(fg);
            }
            return Native.CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
        }

        IntPtr OnKeyLL(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && wParam.ToInt32() == Native.WM_SYSKEYDOWN)
            {
                var info = (Native.KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(Native.KBDLLHOOKSTRUCT));
                IntPtr fg = Native.GetForegroundWindow();
                if (info.vkCode == 0x73 /* F4 */ && (info.flags & 0x20 /* ALT */) != 0 && GroupOf(fg) != null) SetCloseIntent(fg);
            }
            return Native.CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
        }

        // Office 标题栏右上角的关闭按钮区域
        static bool IsOverCloseButton(IntPtr h, int x, int y)
        {
            Native.RECT r;
            if (!Native.GetWindowRect(h, out r)) return false;
            Native.RECT f = Native.FrameBounds(h, r);
            float ws = Native.MonitorScale(h);
            // 只算标题栏那一行；再往下是“批注 / 共享”按钮
            return x >= f.Right - 50 * ws && x <= f.Right && y >= f.Top && y <= f.Top + 46 * ws;
        }

        void SetCloseIntent(IntPtr h)
        {
            closeIntent = h;
            closeIntentTime = DateTime.Now;
        }

        // 关闭一个文档窗口，与点它标题栏上的 × 相同。不能发 WM_CLOSE：Excel 收到它会退出整个程序，所有工作簿一起关掉
        public static void CloseDoc(IntPtr h)
        {
            Native.PostMessage(h, Native.WM_SYSCOMMAND, (IntPtr)Native.SC_CLOSE, IntPtr.Zero);
        }

        public void StartCascade(IEnumerable<IntPtr> handles)
        {
            cascade.Clear();
            cascade.AddRange(handles);
            CascadeNext();
        }

        // 一次只关一个，前一个真正关掉后再关下一个，保存提示会依次弹出
        void CascadeNext()
        {
            cascadeCur = IntPtr.Zero;
            while (cascade.Count > 0)
            {
                IntPtr h = cascade[0];
                cascade.RemoveAt(0);
                if (!Native.IsWindow(h) || GroupOf(h) == null) continue;
                cascadeCur = h;
                cascadeStepTime = DateTime.Now;
                CloseDoc(h);
                return;
            }
        }

        void StopCascade()
        {
            cascade.Clear();
            cascadeCur = IntPtr.Zero;
        }

        // ---------------- 窗口扫描 ----------------

        public void ScheduleScan()
        {
            scanTimer.Stop();
            scanTimer.Start();
        }

        public void Scan()
        {
            if (exiting) return;
            scanResult = new List<DocTab>();
            Native.EnumWindows(enumProc, IntPtr.Zero);
            var found = new Dictionary<IntPtr, DocTab>();
            foreach (var t in scanResult) found[t.Handle] = t;

            var changed = new HashSet<TabBar>();
            foreach (var g in Groups.ToArray())
            {
                for (int i = g.Tabs.Count - 1; i >= 0; i--)
                {
                    var t = g.Tabs[i];
                    DocTab now;
                    if (!found.TryGetValue(t.Handle, out now))
                    {
                        Log.Write(string.Format("移除标签 0x{0:X} {1}", t.Handle.ToInt64(), t.Title));
                        g.Tabs.RemoveAt(i);
                        changed.Add(g);
                        if (t.Handle == closeIntent)
                        {
                            // 用户关掉了整个窗口：同组其余文档依次关闭
                            closeIntent = IntPtr.Zero;
                            var rest = new List<IntPtr>();
                            foreach (var o in g.Tabs) rest.Add(o.Handle);
                            StartCascade(rest);
                        }
                        continue;
                    }
                    if (now.FullTitle != t.FullTitle)
                    {
                        t.Title = now.Title;
                        t.FullTitle = now.FullTitle;
                        changed.Add(g);
                    }
                    found.Remove(t.Handle);
                }
            }
            if (found.Count > 0)
            {
                // 新打开的文档加入最近使用的那一组
                TabBar target = Active != null && Groups.Contains(Active) ? Active : Groups.Count > 0 ? Groups[0] : NewGroup();
                // EnumWindows 按 Z 序从上到下，倒序追加让较早的窗口排在前面
                for (int i = scanResult.Count - 1; i >= 0; i--)
                    if (found.ContainsKey(scanResult[i].Handle))
                    {
                        Log.Write(string.Format("加入标签 0x{0:X} {1}", scanResult[i].Handle.ToInt64(), scanResult[i].Title));
                        target.Tabs.Add(scanResult[i]); changed.Add(target);
                    }
            }
            foreach (var g in Groups.ToArray()) if (g.Tabs.Count == 0) RemoveGroup(g);

            // 未保存状态按窗口逐个读：SDI 模式下一份文档一个进程，只认一个 Application 会漏掉其它进程
            if (dirtyWatcher != null) dirtyWatcher.SetTargets(scanResult);
            var live = new HashSet<uint>();
            foreach (var t in scanResult) { uint pid; Native.GetWindowThreadProcessId(t.Handle, out pid); live.Add(pid); }
            foreach (var pid in new List<uint>(procNames.Keys)) if (!live.Contains(pid)) procNames.Remove(pid);

            // 离开分组的窗口（如文档已转到另一个桌面）恢复正常的最大化位置
            foreach (var h in new List<IntPtr>(Shifted))
                if (GroupOf(h) == null) { Shifted.Remove(h); if (Native.IsWindow(h)) SetFrameTop(h, 0); }
            // 窗格模式已经不再需要的下移窗口恢复原位
            foreach (var h in new List<IntPtr>(PaneShifted))
                if (!Native.HasLeftDockPane(h) || !Native.IsZoomed(h)) RestorePaneShift(h);
            noShift.Clear();
            foreach (var g in Groups)
                foreach (var t in g.Tabs) EnsureShift(t.Handle);

            if (cascadeCur != IntPtr.Zero && GroupOf(cascadeCur) == null) CascadeNext();
            // 长时间没关掉（保存提示一直没处理）就放弃，避免之后误关
            if (cascadeCur != IntPtr.Zero && (DateTime.Now - cascadeStepTime).TotalMinutes > 3) StopCascade();
            if (closeIntent != IntPtr.Zero && (DateTime.Now - closeIntentTime).TotalMinutes > 3) closeIntent = IntPtr.Zero;

            IntPtr fg = Native.GetForegroundWindow();
            var fgGroup = GroupOf(fg);
            if (fgGroup != null) Active = fgGroup;
            foreach (var g in Groups.ToArray()) g.AfterScan(scanResult, fg, changed.Contains(g));
        }

        bool EnumCallback(IntPtr h, IntPtr l)
        {
            if (!Native.IsWindowVisible(h) && !Parked.Contains(h)) return true;
            if (Native.GetWindow(h, Native.GW_OWNER) != IntPtr.Zero) return true;
            DocKind kind;
            if (!TryKind(Native.GetClassNameOf(h), out kind)) return true;
            if (Native.IsCloaked(h)) return true;
            string title = Native.GetWindowTextOf(h);
            if (string.IsNullOrEmpty(title)) return true;
            // 只认 Office 本体，WPS 等兼容程序可能使用相同的窗口类名
            string exe = kind == DocKind.Word ? "WINWORD.EXE" : "EXCEL.EXE";
            if (!string.Equals(ProcessName(h), exe, StringComparison.OrdinalIgnoreCase)) return true;
            scanResult.Add(new DocTab { Handle = h, Kind = kind, Title = CleanTitle(title), FullTitle = title });
            return true;
        }

        static bool TryKind(string cls, out DocKind kind)
        {
            kind = DocKind.Word;
            switch (cls)
            {
                case "OpusApp": kind = DocKind.Word; return true;
                case "XLMAIN": kind = DocKind.Excel; return true;
            }
            return false;
        }

        static string CleanTitle(string t)
        {
            // "报告.docx  -  兼容模式 - Word" → "报告.docx"
            int last = t.LastIndexOf(" - ", StringComparison.Ordinal);
            if (last > 0) t = t.Substring(0, last);
            int mode = t.LastIndexOf("  -  ", StringComparison.Ordinal);
            if (mode > 0) t = t.Substring(0, mode);
            return t.Trim();
        }

        string ProcessName(IntPtr h)
        {
            uint pid;
            Native.GetWindowThreadProcessId(h, out pid);
            string name;
            if (!procNames.TryGetValue(pid, out name))
            {
                name = System.IO.Path.GetFileName(Native.GetProcessPath(pid) ?? "");
                procNames[pid] = name;
            }
            return name;
        }

        void OnWinEvent(IntPtr hook, uint ev, IntPtr hwnd, int idObject, int idChild, uint thread, uint time)
        {
            if (exiting || hwnd == IntPtr.Zero || idObject != Native.OBJID_WINDOW || idChild != 0) return;
            TabBar g;
            switch (ev)
            {
                case Native.EVENT_SYSTEM_MOVESIZESTART:
                    g = GroupOf(hwnd);
                    // 被拖的窗口不一定已经是“当前文档”（点了另一份文档直接拖）。
                    // 漏判就不会把其它窗口停到屏幕外，拖动时背景窗口会一路跟着走。
                    if (g != null && DragGroup == null)
                    {
                        if (hwnd != g.CurDoc) g.CurDoc = hwnd;
                        BeginWindowDrag(g);
                    }
                    return;
                case Native.EVENT_SYSTEM_MOVESIZEEND:
                    // 从标签栏拖动时由 EndDrag 收尾
                    if (DragGroup != null && moveWin == IntPtr.Zero) EndWindowDrag();
                    return;
                case Native.EVENT_OBJECT_LOCATIONCHANGE:
                    g = GroupOf(hwnd);
                    if (g != null) g.OnDocLocation(hwnd);
                    return;
                case Native.EVENT_SYSTEM_FOREGROUND:
                    // 保存提示里点了“取消”，文档重新回到前台：不再整组关闭
                    if (hwnd == closeIntent && (DateTime.Now - closeIntentTime).TotalMilliseconds > 400) closeIntent = IntPtr.Zero;
                    if (hwnd == cascadeCur && (DateTime.Now - cascadeStepTime).TotalMilliseconds > 1500) StopCascade();
                    g = GroupOf(hwnd);
                    if (g != null)
                    {
                        Active = g;
                        g.OnDocForeground(hwnd);
                    }
                    else
                    {
                        // 刚打开的文档还没收进来：马上扫描，别让它以未下移的最大化状态盖住标签栏
                        DocKind fgKind;
                        if (TryKind(Native.GetClassNameOf(hwnd), out fgKind)) { Scan(); return; }
                        // 已经切到别的程序了：立刻放开置顶保持，别让标签栏在其他软件上悬着
                        foreach (var gg in Groups) gg.ClearFrontHold();
                    }
                    UpdateAll();
                    ScheduleScan();
                    return;
                case Native.EVENT_SYSTEM_MINIMIZESTART:
                    // 组内任一文档被最小化就全组最小化（前台可能已先切到另一个文档，不能只看当前文档）
                    g = GroupOf(hwnd);
                    if (g != null) { g.MinimizeOthers(hwnd); g.UpdatePlacement(); }
                    return;
                case Native.EVENT_OBJECT_DESTROY:
                    if (GroupOf(hwnd) != null) ScheduleScan();
                    return;
            }
            if (Native.GetAncestor(hwnd, Native.GA_ROOT) != hwnd) return;
            DocKind kind;
            if (TryKind(Native.GetClassNameOf(hwnd), out kind)) ScheduleScan();
        }

        // ---------------- 拖动窗口时的同组处理 ----------------

        void BeginWindowDrag(TabBar g)
        {
            DragGroup = g;
            dragStart = DateTime.Now;
            trackTimer.Stop();
            trackTimer.Start();
            revealTimer.Stop();
            foreach (var t in g.Tabs)
            {
                IntPtr o = t.Handle;
                if (o == g.CurDoc || Native.IsIconic(o) || !Native.IsWindowVisible(o) || Parked.Contains(o)) continue;
                Native.RECT r;
                if (!Native.GetWindowRect(o, out r)) continue;
                Parked.Add(o);
                parkedFrom[o] = r;
                Point spot;
                if (ParkSpot(o, r, out spot))
                    Native.SetWindowPos(o, IntPtr.Zero, spot.X, spot.Y, 0, 0,
                        Native.SWP_NOSIZE | Native.SWP_NOZORDER | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER | Native.SWP_ASYNCWINDOWPOS);
                else
                    Native.ShowWindowAsync(o, Native.SW_HIDE);
            }
        }

        // 屏幕外的停放位置：就在窗口所在显示器的正上方（或正下方），离它最近，Office 不会因为换了显示器而按新缩放重排
        static bool ParkSpot(IntPtr h, Native.RECT r, out Point spot)
        {
            Rectangle mon = Screen.FromHandle(h).Bounds;
            int w = r.Right - r.Left, ht = r.Bottom - r.Top;
            foreach (int y in new[] { mon.Top - ht - 60, mon.Bottom + 60 })
            {
                var rect = new Rectangle(r.Left, y, w, ht);
                bool clear = true;
                foreach (var s in Screen.AllScreens) if (s.Bounds.IntersectsWith(rect)) { clear = false; break; }
                if (clear) { spot = rect.Location; return true; }
            }
            spot = Point.Empty;
            return false;
        }

        void EndWindowDrag()
        {
            var g = DragGroup;
            DragGroup = null;
            trackTimer.Stop();
            revealTimer.Stop();
            TryMergeOnDrop(g, Cursor.Position);
            // 松手立刻归位，不要再等计时器：等一拍就是肉眼可见的“晚跟随”
            RevealParked();
            // 松手后按当前文档的真实位置对齐整组：拖动结束时窗口的落点才是准的
            if (g != null) g.ScheduleSync();
        }

        void RevealParked()
        {
            if (Parked.Count == 0) return;
            var touched = new HashSet<TabBar>();
            foreach (var o in new List<IntPtr>(Parked))
            {
                if (!Native.IsWindow(o)) continue;
                var g = GroupOf(o);
                Native.RECT from;
                bool had = parkedFrom.TryGetValue(o, out from);
                bool groupNormal = g != null && g.HasState &&
                    (!g.GroupMax || (g.CurDoc != IntPtr.Zero && Native.IsWindow(g.CurDoc) && !Native.IsZoomed(g.CurDoc)));
                if (groupNormal)
                {
                    // 摆到与当前文档相同的位置：拖动刚结束时 GroupNormal 可能还是拖之前的坐标，
                    // 先按当前文档的真实位置校正，避免出现“先回旧位置再跳到新位置”的跟随感
                    if (g.CurDoc != IntPtr.Zero && Native.IsWindow(g.CurDoc) && !Native.IsIconic(g.CurDoc) && !Native.IsZoomed(g.CurDoc))
                        g.GroupNormal = Native.GetPlacement(g.CurDoc).rcNormalPosition;
                    var wp = Native.GetPlacement(o);
                    wp.flags = 0;
                    wp.showCmd = Native.SW_SHOWNOACTIVATE;
                    wp.rcNormalPosition = g.GroupNormal;
                    // 从最大化还原时 Windows 会播放缩放动画，看起来就是背景窗口“跟着”动了一下
                    Native.DisableTransitions(o);
                    Native.SetWindowPlacement(o, ref wp);
                    Shifted.Remove(o);
                    touched.Add(g);
                }
                else if (had)
                {
                    // 整组仍是最大化：放回原处
                    Native.SetWindowPos(o, IntPtr.Zero, from.Left, from.Top, from.Right - from.Left, from.Bottom - from.Top,
                        Native.SWP_NOZORDER | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER | Native.SWP_ASYNCWINDOWPOS);
                }
                // 没地方停放、只好隐藏的那些
                if (!Native.IsWindowVisible(o)) Native.ShowWindow(o, Native.SW_SHOWNOACTIVATE);
                // 放在当前文档后面，不要盖住它
                if (g != null && g.CurDoc != IntPtr.Zero && g.CurDoc != o && Native.IsWindow(g.CurDoc))
                    Native.SetWindowPos(o, g.CurDoc, 0, 0, 0, 0,
                        Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
            }
            Parked.Clear();
            parkedFrom.Clear();
            // 跨到缩放不同的屏幕时，Office 会按新 DPI 自行调整窗口大小，稍后再校正一次
            foreach (var g in touched) g.ScheduleSync();
            ScheduleScan();
        }

        // ---------------- 最大化窗口下移 ----------------

        // 最大化的文档窗口整体下移，把顶部让给标签栏
        public void EnsureShift(IntPtr h)
        {
            if (Manager.BelowRibbon || Manager.TitleBarTabs) return;
            if (!Native.IsZoomed(h) || Parked.Contains(h)) return;
            int bh = BarHeightFor(h);
            if (SetFrameTop(h, bh)) Shifted.Add(h);
        }

        // 功能区下方模式遇到左侧导航/编辑器窗格时，把最大化窗口下移，给标签栏让出一条
        // 位于窗口上方的独立横条；这样标签栏保持全宽，又不会压住窗格标题和关闭按钮。
        public void EnsurePaneShift(IntPtr h)
        {
            if (Manager.TitleBarTabs || !Native.IsZoomed(h) || Parked.Contains(h)) return;
            int bh = BarHeightFor(h);
            if (SetFrameTop(h, bh)) PaneShifted.Add(h);
        }

        public void RestorePaneShift(IntPtr h)
        {
            if (!PaneShifted.Remove(h)) return;
            if (Native.IsWindow(h)) SetFrameTop(h, 0);
        }

        public void RestoreAllPaneShifts()
        {
            foreach (var h in new List<IntPtr>(PaneShifted)) RestorePaneShift(h);
        }

        // 窗口所在显示器上的标签栏高度
        public static int BarHeightFor(IntPtr h)
        {
            return (int)Math.Round(30 * Native.MonitorScale(h));
        }

        // 后台把文档真正最大化。SetWindowPlacement 不会激活窗口，也不会播放“从小放大”的动画；
        // 切过去时窗口已经是最大化尺寸，只需要换前台。
        public void MaximizeNoActivate(IntPtr h)
        {
            if (!Native.IsWindow(h) || Native.IsZoomed(h) || Native.IsIconic(h)) return;
            MarkSelfChange(h, true);
            var wp = Native.GetPlacement(h);
            wp.flags = 0;
            wp.showCmd = Native.SW_MAXIMIZE;
            Native.SetWindowPlacement(h, ref wp);
        }

        public bool IsMax(IntPtr h)
        {
            return Native.IsZoomed(h);
        }

        // 把最大化窗口的可见上边沿放到“工作区顶部 + offset”，底边不动；返回是否处于偏移状态
        bool SetFrameTop(IntPtr h, int offset)
        {
            if (!Native.IsZoomed(h)) return false;
            Native.RECT r;
            if (!Native.GetWindowRect(h, out r)) return false;
            Native.RECT f = Native.FrameBounds(h, r);
            Rectangle wa = Screen.FromHandle(h).WorkingArea;
            // 按工作区算出绝对目标位置（可见边框上沿 = 工作区顶 + offset，下沿 = 工作区底），重复调用结果不变。
            // Office 最大化时会在顶部再留一段与隐形边框等高的空白（平时在屏幕外），下移时让标签栏把它也盖住，避免出现缝隙
            int border = f.Top - r.Top;
            int pad = offset != 0 ? border : 0;
            int top = wa.Top + offset - pad - border;
            int bottom = wa.Bottom + (r.Bottom - f.Bottom);
            if (Math.Abs(r.Top - top) > 1 || Math.Abs(r.Bottom - bottom) > 1)
                Native.SetWindowPos(h, IntPtr.Zero, r.Left, top, r.Right - r.Left, bottom - top,
                    Native.SWP_NOZORDER | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
            return offset != 0;
        }

        void RestoreShifted()
        {
            foreach (var h in Shifted) if (Native.IsWindow(h)) SetFrameTop(h, 0);
            Shifted.Clear();
        }

        void OnDisplayChanged(object sender, EventArgs e)
        {
            foreach (var g in Groups) foreach (var t in g.Tabs) EnsureShift(t.Handle);
            UpdateAll();
        }

        // ---------------- 托盘 / 自启动 / 插件 / 退出 ----------------

        static bool IsAutostart()
        {
            using (var k = Registry.CurrentUser.OpenSubKey(RunKey))
                return k != null && k.GetValue(RunName) != null;
        }

        void ToggleAutostart()
        {
            using (var k = Registry.CurrentUser.CreateSubKey(RunKey))
            {
                if (IsAutostart()) k.DeleteValue(RunName, false);
                else k.SetValue(RunName, "\"" + Application.ExecutablePath + "\"");
            }
            autostartItem.Checked = IsAutostart();
        }

        static void ClearUserExited()
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(OfficeAddin.SettingsKey, true))
                    if (k != null) k.DeleteValue(OfficeAddin.UserExitedValue, false);
            }
            catch { }
        }

        static bool ReadBelowRibbon()
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(OfficeAddin.SettingsKey))
                    if (k != null) return Convert.ToInt32(k.GetValue("BelowRibbon", 0)) != 0;
            }
            catch { }
            return false;
        }

        public void ToggleBelowRibbon()
        {
            BelowRibbon = !BelowRibbon;
            // 功能区下方和标题栏内嵌只能二选一。
            if (BelowRibbon) TitleBarTabs = false;
            SaveLayoutFlags();
            if (placementItem != null) placementItem.Checked = BelowRibbon;
            if (titleItem != null) titleItem.Checked = TitleBarTabs;
            RestoreAllPaneShifts();
            RestoreShifted();
            foreach (var g in Groups.ToArray())
                foreach (var t in g.Tabs) EnsureShift(t.Handle);
            UpdateAll();
        }

        static bool ReadTitleBarTabs()
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(OfficeAddin.SettingsKey))
                    if (k != null) return Convert.ToInt32(k.GetValue("TitleBarTabs", 0)) != 0;
            }
            catch { }
            return false;
        }

        public void ToggleTitleBar()
        {
            TitleBarTabs = !TitleBarTabs;
            // 标题栏内嵌和功能区下方只能二选一。
            if (TitleBarTabs) BelowRibbon = false;
            SaveLayoutFlags();
            if (placementItem != null) placementItem.Checked = BelowRibbon;
            if (titleItem != null) titleItem.Checked = TitleBarTabs;
            // 退出标题栏模式时要把之前下移的窗口放回去；进入时不能动窗口
            RestoreAllPaneShifts();
            RestoreShifted();
            foreach (var g in Groups.ToArray())
                foreach (var t in g.Tabs) EnsureShift(t.Handle);
            UpdateAll();
        }

        static void SaveLayoutFlags()
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(OfficeAddin.SettingsKey))
                {
                    k.SetValue("BelowRibbon", BelowRibbon ? 1 : 0, RegistryValueKind.DWord);
                    k.SetValue("TitleBarTabs", TitleBarTabs ? 1 : 0, RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        // 采样标签栏正下方的正文画布颜色（Word 的灰底 / 页面底色），标签栏取同一个颜色就能和页面融成一片。
        // 多点取样取众数，避开偶尔扫到的文字或滚动条；颜色明显不合理（纯黑纯白）时保留上一次的值。
        public void SampleCanvas(IntPtr doc, Rectangle bar)
        {
            if (!BelowRibbon || bar.IsEmpty) return;
            // 拖动窗口时每帧都在摆标签栏，取色会卡住跟随；拖动时画布颜色也不会变
            if (DragGroup != null) return;
            int now = Environment.TickCount;
            if (now - canvasStamp < 250) return;
            canvasStamp = now;
            Native.RECT wr;
            if (doc == IntPtr.Zero || !Native.GetWindowRect(doc, out wr)) return;
            Native.RECT fr = Native.FrameBounds(doc, wr);
            if (fr.Right - fr.Left < 200) return;
            // 一次抓取整条画布再取样：逐点 GetPixel 在 DWM 下要上百毫秒，会让标签栏一顿一顿
            var cap = new Rectangle(fr.Left, bar.Bottom + 10, fr.Right - fr.Left, 49);
            cap.Intersect(SystemInformation.VirtualScreen);
            if (cap.Width < 60 || cap.Height < 16) return;
            // 取样条被别的窗口压住时，抓屏抓到的是那个窗口的内容（灰底一样能通过中性灰过滤）。
            // 先确认取样点属于本文档的顶层窗口，否则这一轮直接放弃，保留上一次的颜色。
            IntPtr onTop = Native.WindowFromPoint(new Native.POINT { X = cap.Left + 2, Y = cap.Top + 2 });
            if (onTop == IntPtr.Zero ||
                Native.GetAncestor(onTop, Native.GA_ROOT) != Native.GetAncestor(doc, Native.GA_ROOT)) return;
            var hist = new Dictionary<int, int>();
            try
            {
                using (var bmp = new Bitmap(cap.Width, cap.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                {
                    using (var g = Graphics.FromImage(bmp))
                        g.CopyFromScreen(cap.Left, cap.Top, 0, 0, cap.Size, CopyPixelOperation.SourceCopy);
                    int[] xs = { fr.Left + 22, fr.Left + 40, fr.Left + 58, fr.Right - 26, fr.Right - 44, fr.Right - 62 };
                    for (int dy = 10; dy <= 58; dy += 16)
                    {
                        int sy = bar.Bottom + dy;
                        if (sy < cap.Top || sy >= cap.Bottom) continue;
                        foreach (int x in xs)
                        {
                            if (x < cap.Left || x >= cap.Right) continue;
                            Color c = bmp.GetPixel(x - cap.Left, sy - cap.Top);
                            // 只要中性灰：彩色的页面/正文、纯白纸、深色文字都排除
                            int mx = Math.Max(c.R, Math.Max(c.G, c.B)), mn = Math.Min(c.R, Math.Min(c.G, c.B));
                            int lum0 = (c.R * 299 + c.G * 587 + c.B * 114) / 1000;
                            if (mx - mn > 12 || lum0 < 55 || lum0 > 245) continue;
                            int key = (c.R << 16) | (c.G << 8) | c.B;
                            int n;
                            hist.TryGetValue(key, out n);
                            hist[key] = n + 1;
                        }
                    }
                }
            }
            catch { return; }
            int best = 0, bestN = 0;
            foreach (var kv in hist) if (kv.Value > bestN) { best = kv.Key; bestN = kv.Value; }
            if (bestN < 8) return;
            Color col = Color.FromArgb((best >> 16) & 0xFF, (best >> 8) & 0xFF, best & 0xFF);
            int lum = (col.R * 299 + col.G * 587 + col.B * 114) / 1000;
            if (lum < 55 || lum > 245) return;
            if (col == CanvasColor) return;
            CanvasColor = col;
            CanvasLight = lum > 150;
            foreach (var g in Groups.ToArray())
            {
                try { g.BackColor = col; g.Invalidate(); } catch { }
            }
        }
        // 第一次运行时默认登记为 Office 插件；之后尊重用户在托盘里的选择，只在程序挪了位置或升级后刷新登记信息
        void InitAddin()
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(OfficeAddin.SettingsKey))
                {
                    if (k.GetValue("AddinOffered") == null)
                    {
                        k.SetValue("AddinOffered", 1, RegistryValueKind.DWord);
                        OfficeAddin.Register();
                        tray.ShowBalloonTip(5000, "DocTabs",
                            "已登记为 Office 插件，以后打开 Word / Excel 会自动启动标签栏。可在托盘菜单中关闭。", ToolTipIcon.Info);
                    }
                    else if (OfficeAddin.IsRegistered() && !OfficeAddin.IsCurrent())
                    {
                        OfficeAddin.Register();
                        Log.Write("已更新 Office 插件登记（程序位置/版本变化，或 Office 停用过插件）");
                    }
                }
            }
            catch { }
            addinItem.Checked = OfficeAddin.IsRegistered();
        }

        void ToggleAddin()
        {
            try
            {
                if (OfficeAddin.IsRegistered()) OfficeAddin.Unregister();
                else OfficeAddin.Register();
            }
            catch (Exception ex)
            {
                MessageBox.Show("修改 Office 插件登记失败：" + ex.Message, "DocTabs", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            addinItem.Checked = OfficeAddin.IsRegistered();
        }

        void CheckOfficeGone()
        {
            bool any = false;
            foreach (var name in new[] { "WINWORD", "EXCEL" })
            {
                foreach (var p in System.Diagnostics.Process.GetProcessesByName(name))
                {
                    // 崩溃后残留、已无法访问的进程不算
                    try { if (!p.HasExited) any = true; } catch { }
                    p.Dispose();
                }
            }
            officeGoneTicks = any ? 0 : officeGoneTicks + 1;
            if (officeGoneTicks >= 3)
            {
                Log.Write("Word / Excel 已全部关闭，退出");
                ExitApp(false);
            }
        }

        static Icon MakeTrayIcon(float s)
        {
            int size = Math.Max(16, (int)Math.Round(16 * s));
            using (var bmp = new Bitmap(size, size))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    float u = size / 16f;
                    using (var back = new SolidBrush(Color.FromArgb(120, 130, 145)))
                        g.FillRectangle(back, 1 * u, 3 * u, 6 * u, 4 * u);
                    using (var front = new SolidBrush(Color.FromArgb(24, 90, 189)))
                    {
                        g.FillRectangle(front, 7 * u, 3 * u, 7 * u, 4 * u);
                        g.FillRectangle(front, 1 * u, 7 * u, 14 * u, 7 * u);
                    }
                }
                IntPtr hicon = bmp.GetHicon();
                var icon = (Icon)Icon.FromHandle(hicon).Clone();
                Native.DestroyIcon(hicon);
                return icon;
            }
        }

        // byUser：用户从菜单主动退出，Office 插件在本次 Office 运行期间不再自动把它拉起来
        public void ExitApp(bool byUser)
        {
            if (exiting) return;
            exiting = true;
            if (byUser)
            {
                Log.Write("用户选择退出");
                try
                {
                    using (var k = Registry.CurrentUser.CreateSubKey(OfficeAddin.SettingsKey))
                        k.SetValue(OfficeAddin.UserExitedValue, 1, RegistryValueKind.DWord);
                }
                catch { }
            }
            if (dirtyWatcher != null) dirtyWatcher.Stop();
            fallbackTimer.Stop();
            scanTimer.Stop();
            revealTimer.Stop();
            moveWatch.Stop();
            trackTimer.Stop();
            foreach (var h in hooks) if (h != IntPtr.Zero) Native.UnhookWinEvent(h);
            if (mouseHook != IntPtr.Zero) Native.UnhookWindowsHookEx(mouseHook);
            if (keyHook != IntPtr.Zero) Native.UnhookWindowsHookEx(keyHook);
            mouseHook = keyHook = IntPtr.Zero;
            hooks.Clear();
            SystemEvents.DisplaySettingsChanged -= OnDisplayChanged;
            DragGroup = null;
            RevealParked();
            RestoreAllPaneShifts();
            RestoreShifted();
            foreach (var g in Groups.ToArray()) g.Destroy();
            Groups.Clear();
            tray.Visible = false;
            tray.Dispose();
            Application.Exit();
        }
    }

    // ---------------- 一个分组：一个标签栏窗口 + 组内文档的联动 ----------------

    class TabBar : Form
    {
        static readonly Color BarBgDark = Color.FromArgb(54, 56, 60);
        static readonly Color BarBg2Dark = Color.FromArgb(40, 42, 45);
        static readonly Color BarLineDark = Color.FromArgb(28, 29, 32);
        // 标题栏模式：从屏幕采样 Word 标题栏底色，标签栏跟着 Word 的主题走
        public static Color TitleCaption = Color.FromArgb(243, 243, 243);
        public static bool TitleDark;
        // 标题栏底色取到一次就锁定：Office 在运行期间不会自己换标题栏颜色，反复采样只会在
        // 别的窗口压住采样点时把别人的颜色吸进来。换文档、换主题（WM_SETTINGCHANGE /
        // WM_THEMECHANGED）才重新取。
        public static bool TitleLocked;
        public static IntPtr TitleLockDoc = IntPtr.Zero;
        static Color titlePending = Color.Empty;
        static bool titlePendingDark;
        static int titlePendingCount;
        static Color BarBg { get { return Manager.TitleBarTabs ? TitleCaption : BarBgDark; } }
        static Color BarBg2 { get { return Manager.TitleBarTabs ? Blend(TitleCaption, -6) : BarBg2Dark; } }
        static Color BarLine { get { return Manager.TitleBarTabs ? Blend(TitleCaption, -18) : BarLineDark; } }
        // 功能区下方模式：整条标签栏取正文画布的颜色，和下面的页面连成一片，不再是一条深色带
        static Color BarColor { get { return Manager.TitleBarTabs ? TitleCaption : Manager.BelowRibbon ? Manager.CanvasColor : BarBgDark; } }
        static Color ActiveBg
        {
            get
            {
                if (Manager.TitleBarTabs) return TitleDark ? Blend(TitleCaption, 14) : Color.FromArgb(255, 255, 255);
                return Manager.BelowRibbon ? Manager.CanvasColor : Color.FromArgb(64, 67, 72);
            }
        }
        static Color HoverBg
        {
            get
            {
                if (Manager.TitleBarTabs) return TitleDark ? Blend(TitleCaption, 8) : Blend(TitleCaption, 55);
                return Manager.BelowRibbon ? Blend(Manager.CanvasColor, 6) : Color.FromArgb(58, 61, 66);
            }
        }
        static Color InactiveBg
        {
            get
            {
                if (Manager.TitleBarTabs) return TitleCaption;
                return Manager.BelowRibbon ? Blend(Manager.CanvasColor, -8) : BarBgDark;
            }
        }
        static bool LightCanvas { get { return (Manager.BelowRibbon && Manager.CanvasLight) || (Manager.TitleBarTabs && !TitleDark); } }
        static Color TextColor { get { return Manager.TitleBarTabs ? (TitleDark ? Color.FromArgb(242, 243, 245) : Color.FromArgb(30, 32, 36)) : LightCanvas ? Color.FromArgb(26, 28, 32) : Color.FromArgb(250, 251, 253); } }
        static Color HoverText { get { return Manager.TitleBarTabs ? (TitleDark ? Color.FromArgb(255, 255, 255) : Color.FromArgb(16, 18, 22)) : LightCanvas ? Color.FromArgb(18, 20, 24) : Color.FromArgb(238, 241, 244); } }
        static Color DimText { get { return Manager.TitleBarTabs ? (TitleDark ? Color.FromArgb(196, 198, 202) : Color.FromArgb(96, 100, 106)) : LightCanvas ? Color.FromArgb(86, 90, 96) : Color.FromArgb(188, 191, 196); } }
        static Color SepColor { get { return Manager.TitleBarTabs ? (TitleDark ? Color.FromArgb(40, 255, 255, 255) : Color.FromArgb(26, 0, 0, 0)) : LightCanvas ? Color.FromArgb(46, 0, 0, 0) : Color.FromArgb(38, 255, 255, 255); } }
        // pct > 0 往白色提亮，pct < 0 往黑色压暗
        static Color Blend(Color c, int pct)
        {
            if (pct >= 0) return Color.FromArgb(c.R + (255 - c.R) * pct / 100, c.G + (255 - c.G) * pct / 100, c.B + (255 - c.B) * pct / 100);
            int p = -pct;
            return Color.FromArgb(c.R - c.R * p / 100, c.G - c.G * p / 100, c.B - c.B * p / 100);
        }
        static readonly Color Accent = Color.FromArgb(76, 148, 255);
        // 浮动模式下标签之外的区域用这个颜色绘制，配合 TransparencyKey 变成透明且可点穿
        static readonly Color KeyColor = Color.FromArgb(201, 203, 207);

        readonly Manager mgr;
        public readonly List<DocTab> Tabs = new List<DocTab>();
        // 组内的当前文档（最近一次处于前台的文档窗口）
        public IntPtr CurDoc;
        // 同组共享的窗口状态
        public bool HasState;
        public bool GroupMax;
        public Native.RECT GroupNormal;

        BarMode mode = BarMode.Hidden;
        bool isTopmost;
        int lastInFrontTick;
        bool overlayOpen, overlayApplied;
        int overlayStamp;
        int topTicks;
        IntPtr retryDoc;
        int retryTicks;
        Rectangle lastRect;
        int titleSampleTick;
        bool reassertZ;
        bool overlayShown;
        int hoverTicks, leaveTicks;
        bool destroying;

        int hoverIndex = -1;
        bool hoverClose;
        int pressIndex = -1;
        int lastCloseTick;
        // 拖动时 DWM 的扩展边框/按钮矩形会停在旧位置：缓存相对窗口的差值，用实时的窗口矩形换算回来
        IntPtr frameRelDoc = IntPtr.Zero;
        int frameRelL, frameRelT, frameRelR, frameRelB;
        // 拖动期间整条标签栏相对窗口的矩形：只在窗口尺寸变化时重算，避免每帧高度/顶边抖一下
        int dragPlaceTick;
        IntPtr relBarDoc = IntPtr.Zero;
        int relBarLeft, relBarTop, relBarW, relBarH, relBarWinW, relBarWinH;

        IntPtr btnRelDoc = IntPtr.Zero;
        int btnRelLeftFromRight, btnRelBottomFromTop, btnRelWidth;
        // 高亮边高度缓存：屏幕取色要等 DWM 合成，拖动时不能每帧都做
        bool lastZoomed;
        int zoomFlipTick;
        IntPtr topInsetDoc = IntPtr.Zero;
        int topInsetSampleH;
        int topInsetValue;
        int topInsetTick;
        Point pressPoint;
        Point pressScreen;
        bool pressOnBar;
        bool dragging;
        // 当前文档所在显示器的缩放比例；主副屏缩放不同时随文档所在的屏幕变化
        float s;
        int barHeight;

        Font tabFont;
        Font tabFontBold;
        Font badgeFont;
        readonly System.Windows.Forms.Timer syncTimer;
        readonly System.Windows.Forms.Timer hoverTimer;
        readonly System.Windows.Forms.Timer pressTimer;
        readonly System.Windows.Forms.Timer roomTimer;
        readonly System.Windows.Forms.Timer topTimer;
        readonly System.Windows.Forms.Timer retryTimer;
        readonly System.Windows.Forms.Timer topmostTimer;
        readonly ToolTip tip;

        public TabBar(Manager mgr)
        {
            this.mgr = mgr;
            UpdateScale(IntPtr.Zero);

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            BackColor = BarBg;
            TransparencyKey = Manager.BelowRibbon ? Color.Empty : KeyColor;
            Text = "DocTabs";
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                     ControlStyles.Opaque, true);
            Bounds = new Rectangle(0, -barHeight * 3, 200, barHeight);
            CreateHandle();

            tip = new ToolTip { ShowAlways = true, InitialDelay = 600, ReshowDelay = 200 };
            // 拖动/缩放结束后再把其他文档同步到同一位置，避免拖动过程中反复重排
            syncTimer = new System.Windows.Forms.Timer { Interval = 150 };
            syncTimer.Tick += delegate { syncTimer.Stop(); SyncOthersNormal(); };
            hoverTimer = new System.Windows.Forms.Timer { Interval = 100 };
            hoverTimer.Tick += delegate { OnHoverTick(); };
            // 按住标签栏期间轮询鼠标：标签栏从不激活，鼠标一离开它就收不到鼠标消息，拖快了也不能漏判
            pressTimer = new System.Windows.Forms.Timer { Interval = 10 };
            pressTimer.Tick += delegate { PressTick(); };
            // 窗口贴到屏幕上沿后，稍等位置稳定（系统分屏动画结束）再往下让出标签栏
            roomTimer = new System.Windows.Forms.Timer { Interval = 250 };
            roomTimer.Tick += delegate { roomTimer.Stop(); MakeRoom(); };
            // 从任务栏/Alt+Tab 切进来后，系统会在几十毫秒内继续调整 Z 序；多压几次置顶
            topTimer = new System.Windows.Forms.Timer { Interval = 90 };
            topTimer.Tick += delegate
            {
                topTimer.Stop();
                reassertZ = true;
                UpdatePlacement();
                if (--topTicks > 0) topTimer.Start();
            };
            // 前台锁偶尔会让 SetForegroundWindow 失败，短暂重试
            retryTimer = new System.Windows.Forms.Timer { Interval = 60 };
            retryTimer.Tick += delegate
            {
                retryTimer.Stop();
                if (retryTicks <= 0 || retryDoc == IntPtr.Zero || !Native.IsWindow(retryDoc)) { retryDoc = IntPtr.Zero; return; }
                retryTicks--;
                if (Native.GetForegroundWindow() != retryDoc)
                {
                    Manager.BringToFront(retryDoc, false);
                    if (Native.GetForegroundWindow() != retryDoc) retryTimer.Start();
                }
            };
            // 置顶看门狗：切到别的程序后必须尽快取消置顶，不能让标签栏悬在别的软件上面
            topmostTimer = new System.Windows.Forms.Timer { Interval = 250 };
            topmostTimer.Tick += delegate
            {
                bool ov = OverlayOpen();
                if (ov != overlayApplied) { overlayApplied = ov; UpdatePlacement(); }
                else if (Visible && Tabs.Count > 0 && mgr.DragGroup == null) UpdatePlacement();
                mgr.ApplyDirtyIfNeeded();
                CheckTopmost();
            };
            topmostTimer.Start();
        }

        public void Destroy()
        {
            destroying = true;
            syncTimer.Stop();
            hoverTimer.Stop();
            pressTimer.Stop();
            roomTimer.Stop();
            topTimer.Stop();
            retryTimer.Stop();
            topmostTimer.Stop();
            tip.Dispose();
            if (tabFont != null) tabFont.Dispose();
            if (tabFontBold != null) tabFontBold.Dispose();
            if (badgeFont != null) badgeFont.Dispose();
            Dispose();
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= Native.WS_EX_TOOLWINDOW | Native.WS_EX_NOACTIVATE;
                return cp;
            }
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Native.WM_MOUSEACTIVATE) { m.Result = (IntPtr)Native.MA_NOACTIVATE; return; }
            // 标签栏的位置和大小完全由我们自己按文档所在屏幕计算，不接受系统按 DPI 建议的尺寸
            if (m.Msg == Native.WM_DPICHANGED) { m.Result = IntPtr.Zero; return; }
            // 系统主题或 Office 主题变了，标题栏底色会跟着变：解锁一次，下次静止时重新取色
            if (m.Msg == 0x001A /* WM_SETTINGCHANGE */ || m.Msg == 0x031A /* WM_THEMECHANGED */)
            {
                TitleLocked = false;
                TitleLockDoc = IntPtr.Zero;
                titlePending = Color.Empty;
                titlePendingCount = 0;
            }
            // 外部发来的关闭请求（如结束任务）：整个程序正常退出，把调整过的窗口恢复原状
            if (m.Msg == Native.WM_CLOSE && !destroying)
            {
                Log.Write("收到外部关闭消息（如关机、注销），退出");
                mgr.ExitApp(false);
                return;
            }
            base.WndProc(ref m);
        }

        public int IndexOf(IntPtr h)
        {
            if (h == IntPtr.Zero) return -1;
            for (int i = 0; i < Tabs.Count; i++) if (Tabs[i].Handle == h) return i;
            return -1;
        }

        public bool TabStripContains(Point screen)
        {
            if (!Bounds.Contains(screen)) return false;
            Point p = PointToClient(screen);
            if (Tabs.Count == 0) return false;
            return p.X <= Tabs[Tabs.Count - 1].Bounds.Right + 40 * s;
        }

        // 扫描之后：确定当前文档、刷新布局和位置
        public void AfterScan(List<DocTab> zOrder, IntPtr fg, bool changed)
        {
            if (IndexOf(CurDoc) < 0) CurDoc = IntPtr.Zero;
            bool fgHere = IndexOf(fg) >= 0;
            if (fgHere && fg != CurDoc) { OnDocForeground(fg); changed = true; }
            else if (!fgHere && mgr.DragGroup != this && !pressOnBar)
            {
                // 本组没有文档在前台（例如从资源管理器打开了新文档，但它没能抢到前台）：
                // 以 Z 序最上面、看得见的文档作为当前文档，标签栏贴在它上面，避免被它盖住
                foreach (var t in zOrder)
                {
                    if (IndexOf(t.Handle) < 0 || !Native.IsWindowVisible(t.Handle) || Native.IsIconic(t.Handle)) continue;
                    if (t.Handle != CurDoc) { CurDoc = t.Handle; changed = true; }
                    break;
                }
            }
            if (CurDoc == IntPtr.Zero && Tabs.Count > 0) { CurDoc = Tabs[0].Handle; changed = true; }
            if (!HasState && CurDoc != IntPtr.Zero) CaptureGroup(CurDoc);
            // 新文档加入、或分组状态变化时，让整组窗口提前对齐（后台完成，不激活）
            if (HasState && changed && mgr.DragGroup == null)
            {
                if (GroupMax) SyncGroupMax();
                else SyncOthersNormal();
            }
            if (changed) LayoutTabs();
            // 只在内容/关系变化时重设置顶；每次扫描都重设会让标签栏在切换时闪一下
            if (changed) reassertZ = true;
            UpdatePlacement();
            if (changed) Invalidate();
        }

        // ---------------- 同组窗口联动 ----------------

        public void OnDocForeground(IntPtr h)
        {
            // 不管是点标签、任务栏缩略图还是 Alt+Tab，只要本组文档到了前台，
            // 都先把标签栏压回最上层，并在一小段时间内持续保持，等系统切换动画结束。
            lastInFrontTick = Environment.TickCount;
            ForceFront();
            if (h == CurDoc) { UpdatePlacement(); Invalidate(); return; }
            CurDoc = h;
            // 已被最小化（例如全组最小化时前台恰好切到它）就不要再把它恢复回来
            if (Native.IsIconic(h)) { Invalidate(); return; }
            if (!HasState) CaptureGroup(h);
            else ApplyGroupState(h, true);
            UpdatePlacement();
            Invalidate();
        }

        public void ClearFrontHold()
        {
            lastInFrontTick = unchecked(Environment.TickCount - 10000);
        }

        // 文件/导出等全屏界面打开时隐藏标签栏；每 200ms 才重新枚举一次子窗口，避免拖动时每帧枚举
        bool OverlayOpen()
        {
            if (CurDoc == IntPtr.Zero || !Native.IsWindow(CurDoc) || mgr.DragGroup == this) return false;
            int now = Environment.TickCount;
            if (unchecked(now - overlayStamp) < 200) return overlayOpen;
            overlayStamp = now;
            overlayOpen = Native.HasFullpageOverlay(CurDoc);
            return overlayOpen;
        }

        // 切到别的程序后，标签栏有一小段时间要保持置顶（防止从任务栏切回 Word 时被盖住）。
        // 万一那期间没有再收到前台通知，置顶就会一直留着，悬在别的软件上面。这里兜底清掉。
        void CheckTopmost()
        {
            if (!Visible || Tabs.Count == 0 || !isTopmost) return;
            IntPtr fg = Native.GetForegroundWindow();
            if (IndexOf(fg) >= 0 || fg == Handle) { lastInFrontTick = Environment.TickCount; return; }
            if (unchecked(Environment.TickCount - lastInFrontTick) < 700) return;
            isTopmost = false;
            reassertZ = false;
            Native.SetWindowPos(Handle, Native.HWND_NOTOPMOST, 0, 0, 0, 0,
                Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
            // 取消置顶后还要压回文档后面，否则它会停在非置顶层的最上面，继续盖住资源管理器等窗口
            if (CurDoc != IntPtr.Zero && Native.IsWindow(CurDoc))
                Native.SetWindowPos(Handle, CurDoc, 0, 0, 0, 0,
                    Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
            Log.Write("标签栏取消置顶（前台已切走）");
        }

        public void ForceFront()
        {
            lastInFrontTick = Environment.TickCount;
            if (!Visible) return;
            Native.SetWindowPos(Handle, Native.HWND_TOPMOST, 0, 0, 0, 0,
                Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
            isTopmost = true;
            reassertZ = true;
            topTicks = 3;
            topTimer.Stop();
            topTimer.Start();
        }

        // Word 有时会把“最大化后的尺寸”记成还原尺寸（尤其被旧版伪最大化写过之后）。
        // 这种值拿来还原等于没还原，落成一个居中窗口才有“还原”的手感。
        Native.RECT SanitizeNormal(IntPtr h, Native.RECT np)
        {
            Rectangle mon = Screen.FromHandle(h).Bounds;
            Rectangle wa = Screen.FromHandle(h).WorkingArea;
            if (wa.Width <= 200 || wa.Height <= 200) return np;
            int ox = wa.Left - mon.Left, oy = wa.Top - mon.Top;
            float ms = Native.MonitorScale(h);
            int l = np.Left + ox, tp = np.Top + oy;
            int w = np.Right - np.Left, ht = np.Bottom - np.Top;
            int tol = (int)Math.Round(28 * ms);
            bool wide = w >= wa.Width - tol || l <= wa.Left - tol / 4 || l + w >= wa.Right + tol / 4;
            bool tall = ht >= wa.Height - tol || tp <= wa.Top - tol / 4 || tp + ht >= wa.Bottom + tol / 4;
            if (!(wide && tall)) return np;
            int nw = (int)Math.Round(wa.Width * 0.78), nh = (int)Math.Round(wa.Height * 0.82);
            int nl = wa.Left + (wa.Width - nw) / 2, nt = wa.Top + (wa.Height - nh) / 2;
            Log.Write(string.Format("还原尺寸异常(全屏)，改为 {0},{1} {2}x{3}", nl, nt, nw, nh));
            return new Native.RECT { Left = nl - ox, Top = nt - oy, Right = nl - ox + nw, Bottom = nt - oy + nh };
        }

        void CaptureGroup(IntPtr h)
        {
            if (Native.IsIconic(h)) return;
            GroupMax = mgr.IsMax(h);
            GroupNormal = Native.GetPlacement(h).rcNormalPosition;
            if (GroupMax) GroupNormal = SanitizeNormal(h, GroupNormal);
            HasState = true;
        }

        // 让窗口 h 与分组状态一致：最大化 / 相同的还原位置
        public void ApplyGroupState(IntPtr h, bool activating)
        {
            if (!HasState)
            {
                if (Native.IsIconic(h)) Native.ShowWindow(h, Native.SW_RESTORE);
                return;
            }
            if (GroupMax)
            {
                // 用真正的最大化，组内每个窗口状态一致；后台完成，不激活、不播动画
                if (!Native.IsZoomed(h) && !Native.IsIconic(h)) mgr.MaximizeNoActivate(h);
                if (!Manager.BelowRibbon) mgr.EnsureShift(h);
            }
            else
            {
                var wp = Native.GetPlacement(h);
                if (Native.IsZoomed(h) || (!Native.IsIconic(h) && !wp.rcNormalPosition.Equals(GroupNormal)))
                {
                    wp.flags = 0;
                    wp.showCmd = Native.SW_SHOWNOACTIVATE;
                    wp.rcNormalPosition = GroupNormal;
                    mgr.MarkSelfChange(h, false);
                    Native.DisableTransitions(h);
                    Native.SetWindowPlacement(h, ref wp);
                    mgr.Shifted.Remove(h);
                }
            }
            // 还原/最大化后台窗口时不要把当前文档盖住；否则看起来像“点了标签却没切过去”
            if (!activating && CurDoc != IntPtr.Zero && CurDoc != h && Native.IsWindow(CurDoc))
                Native.SetWindowPos(h, CurDoc, 0, 0, 0, 0,
                    Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
        }

        public void OnDocLocation(IntPtr h)
        {
            if (Native.IsIconic(h)) return;
            bool zoomed = Native.IsZoomed(h);
            if (h != CurDoc)
            {
                // 不是当前文档：多数是我们自己同步过去的，只保证最大化时让出标签栏
                if (mgr.IsSelfChange(h) || !HasState || mgr.Parked.Contains(h) || mgr.DragGroup == this)
                {
                    if (zoomed) mgr.EnsureShift(h);
                    return;
                }
                // 用户在这个窗口上单独点了最大化/还原（点击会先让它到前台）：
                // 以前这里直接忽略，结果前台窗口还原了、后台还留着最大化，一大一小。
                if (zoomed != GroupMax && Native.GetForegroundWindow() == h)
                {
                    CurDoc = h;
                    GroupMax = zoomed;
                    if (zoomed)
                    {
                        Log.Write(string.Format("整组最大化 0x{0:X}（非当前文档触发）", h.ToInt64()));
                        SyncGroupMax();
                        mgr.EnsureShift(h);
                    }
                    else
                    {
                        GroupNormal = SanitizeNormal(h, Native.GetPlacement(h).rcNormalPosition);
                        Log.Write(string.Format("整组还原 0x{0:X} -> {1},{2} {3}x{4}（非当前文档触发）",
                            h.ToInt64(), GroupNormal.Left, GroupNormal.Top,
                            GroupNormal.Right - GroupNormal.Left, GroupNormal.Bottom - GroupNormal.Top));
                        SyncOthersNormal();
                    }
                    UpdatePlacement();
                    Invalidate();
                    return;
                }
                if (zoomed) mgr.EnsureShift(h);
                return;
            }
            // 这一条位置变化是我们自己刚刚最大化/还原造成的，不是用户点了最大化/还原
            if (mgr.IsSelfChange(h))
            {
                if (zoomed) mgr.EnsureShift(h);
                UpdatePlacement();
                return;
            }
            if (mgr.DragGroup == this)
            {
                // 拖动标签栏之外的位置（例如拖 Word 标题栏）把窗口从最大化拉下来时，
                // 以前整条记录被忽略，松手后同组其它窗口又按最大化放回去，变成一大一小。
                if (!zoomed && GroupMax)
                {
                    GroupMax = false;
                    GroupNormal = SanitizeNormal(h, Native.GetPlacement(h).rcNormalPosition);
                    Log.Write(string.Format("整组还原 0x{0:X}（拖动触发） -> {1},{2} {3}x{4}",
                        h.ToInt64(), GroupNormal.Left, GroupNormal.Top,
                        GroupNormal.Right - GroupNormal.Left, GroupNormal.Bottom - GroupNormal.Top));
                }
                HasState = true;
                // 拖动中每次位置变化都跟着挪一下：WM_TIMER 的实际节拍约 15ms，只靠定时器竖向拖动会落后半帧，
                // 底部就会漏出一条 Word 标题栏（账户头像正好在那儿）。这里按事件流跟随，4ms 节流防抖。
                int nowTick2 = Environment.TickCount;
                if (nowTick2 - dragPlaceTick >= 4)
                {
                    dragPlaceTick = nowTick2;
                    UpdatePlacement();
                }
                return;
            }
            if (zoomed)
            {
                if (!GroupMax)
                {
                    GroupNormal = SanitizeNormal(h, GroupNormal);
                    GroupMax = true;
                    Log.Write(string.Format("整组最大化 0x{0:X}", h.ToInt64()));
                    SyncGroupMax();
                }
                mgr.EnsureShift(h);
            }
            else if (GroupMax)
            {
                // 当前文档从最大化变回普通：整组跟着还原。rcNormalPosition 是最大化前的还原位置，
                // 不会像窗口当前矩形那样在动画中途被读成“半大不小”的尺寸。
                GroupMax = false;
                var rawNormal = Native.GetPlacement(h).rcNormalPosition;
                GroupNormal = SanitizeNormal(h, rawNormal);
                Log.Write(string.Format("整组还原 0x{0:X} -> {1},{2} {3}x{4}",
                    h.ToInt64(), GroupNormal.Left, GroupNormal.Top,
                    GroupNormal.Right - GroupNormal.Left, GroupNormal.Bottom - GroupNormal.Top));
                // Word 自己记的还原尺寸就是全屏时，把当前窗口也落到同一个正常尺寸，
                // 否则只有其他文档缩了，当前窗口看起来还是“没还原”
                if (!GroupNormal.Equals(rawNormal))
                {
                    var wp = Native.GetPlacement(h);
                    wp.flags = 0;
                    wp.showCmd = Native.SW_SHOWNOACTIVATE;
                    wp.rcNormalPosition = GroupNormal;
                    mgr.MarkSelfChange(h, false);
                    Native.SetWindowPlacement(h, ref wp);
                }
                SyncOthersNormal();
            }
            else
            {
                var np = Native.GetPlacement(h).rcNormalPosition;
                bool changed = !np.Equals(GroupNormal);
                GroupNormal = np;
                if (changed)
                {
                    // 缩放过程中实时带动其他文档，结束后再由 syncTimer 精确校正
                    LiveSync(h);
                    ScheduleSync();
                }
                // 窗口顶到了屏幕上沿（例如系统分屏）：稍后往下让出标签栏
                if (NeedsRoom(h)) { roomTimer.Stop(); roomTimer.Start(); }
            }
            HasState = true;
            UpdatePlacement();
        }

        public void ScheduleSync()
        {
            syncTimer.Stop();
            syncTimer.Start();
        }

        // 组内其他文档提前真正最大化（不激活、无动画）：切过去是瞬间切换
        void SyncGroupMax()
        {
            foreach (var t in Tabs)
            {
                IntPtr o = t.Handle;
                if (o == CurDoc || !Native.IsWindow(o) || Native.IsIconic(o) || mgr.Parked.Contains(o)) continue;
                ApplyGroupState(o, false);
            }
        }

        void LiveSync(IntPtr h)
        {
            Native.RECT r;
            if (!Native.GetWindowRect(h, out r)) return;
            Native.RECT f = Native.FrameBounds(h, r);
            foreach (var t in Tabs)
            {
                IntPtr o = t.Handle;
                if (o == h || Native.IsIconic(o) || Native.IsZoomed(o) || mgr.Parked.Contains(o)) continue;
                Native.RECT ro;
                if (!Native.GetWindowRect(o, out ro)) continue;
                Native.RECT fo = Native.FrameBounds(o, ro);
                int l = f.Left - (fo.Left - ro.Left), tp = f.Top - (fo.Top - ro.Top);
                int rr = f.Right + (ro.Right - fo.Right), b = f.Bottom + (ro.Bottom - fo.Bottom);
                if (l == ro.Left && tp == ro.Top && rr == ro.Right && b == ro.Bottom) continue;
                // 异步下发，Office 重排再慢也不会卡住标签栏自己
                Native.SetWindowPos(o, IntPtr.Zero, l, tp, rr - l, b - tp,
                    Native.SWP_NOZORDER | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER | Native.SWP_ASYNCWINDOWPOS);
                // 被带动的窗口保持在当前文档下面
                Native.SetWindowPos(o, h, 0, 0, 0, 0,
                    Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
            }
        }

        void SyncOthersNormal()
        {

            if (!HasState || GroupMax || mgr.DragGroup == this) return;
            // 拖动/缩放结束后以当前文档的实际位置为准，否则其它文档会退回拖动前的坐标
            if (CurDoc != IntPtr.Zero && Native.IsWindow(CurDoc) && !Native.IsIconic(CurDoc) && !Native.IsZoomed(CurDoc))
                GroupNormal = Native.GetPlacement(CurDoc).rcNormalPosition;
            foreach (var t in Tabs)
            {
                IntPtr h = t.Handle;
                if (h == CurDoc || Native.IsIconic(h) || mgr.Parked.Contains(h)) continue;
                ApplyGroupState(h, false);
            }
        }

        // 把同组里仍在最大化的其他窗口拉回普通状态（不激活、不动 Z 序）。
        // 从最大化状态拖出窗口时如果只还原当前窗口，背景会留一个最大化窗口，
        // 看起来就是前面一个窗口、后面一个大窗口的错位画面。
        void SyncSiblingsOutOfMax(IntPtr h, TabBar g, Native.RECT n)
        {
            foreach (var t in g.Tabs)
            {
                IntPtr o = t.Handle;
                if (o == h || !Native.IsWindow(o) || Native.IsIconic(o) || mgr.Parked.Contains(o)) continue;
                if (!Native.IsZoomed(o)) continue;
                var wp = Native.GetPlacement(o);
                wp.flags = 0;
                wp.showCmd = Native.SW_SHOWNOACTIVATE;
                wp.rcNormalPosition = n;
                mgr.MarkSelfChange(o, false);
                Native.SetWindowPlacement(o, ref wp);
                mgr.Shifted.Remove(o);
            }
        }


        public void MinimizeOthers(IntPtr h)
        {
            foreach (var t in new List<DocTab>(Tabs))
                if (t.Handle != h && !Native.IsIconic(t.Handle))
                    Native.ShowWindow(t.Handle, Native.SW_SHOWMINNOACTIVE);
        }

        // 窗口上方放不下标签栏（顶到了屏幕上沿）
        bool NeedsRoom(IntPtr h)
        {
            if (Manager.BelowRibbon || Manager.TitleBarTabs) return false;
            Native.RECT r;
            if (!Native.GetWindowRect(h, out r)) return false;
            Native.RECT f = Native.FrameBounds(h, r);
            Rectangle wa = Screen.FromHandle(h).WorkingArea;
            int overlap = (int)Math.Round(2 * s);
            return f.Top < wa.Top + Manager.BarHeightFor(h) - overlap - (int)Math.Round(3 * s);
        }

        // 往下让出标签栏的高度（底边不动）；窗口太矮时保持自动隐藏
        void MakeRoom()
        {
            if (Manager.BelowRibbon || Manager.TitleBarTabs) return;
            IntPtr h = CurDoc;
            if (h == IntPtr.Zero || mgr.DragGroup != null || pressOnBar || Native.IsIconic(h) || mgr.IsMax(h)) return;
            if (!NeedsRoom(h)) return;
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            Native.RECT f = Native.FrameBounds(h, r);
            Rectangle wa = Screen.FromHandle(h).WorkingArea;
            int want = wa.Top + Manager.BarHeightFor(h) - (int)Math.Round(2 * s);
            int dy = want - f.Top;
            if (f.Bottom - want < 240 * s) return;
            Native.SetWindowPos(h, IntPtr.Zero, r.Left, r.Top + dy, r.Right - r.Left, r.Bottom - r.Top - dy,
                Native.SWP_NOZORDER | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
        }

        // 标签栏按当前文档所在显示器的缩放比例绘制
        void UpdateScale(IntPtr h)
        {
            float ns = Native.MonitorScale(h);
            if (tabFont != null && Math.Abs(ns - s) < 0.01f) return;
            s = ns;
            barHeight = (int)Math.Round(30 * s);
            // 按像素建字体：进程是逐显示器 DPI 感知的，按“磅”建的字体在不同缩放的屏幕上大小会不对
            if (tabFont != null) tabFont.Dispose();
            if (tabFontBold != null) tabFontBold.Dispose();
            if (badgeFont != null) badgeFont.Dispose();
            tabFont = new Font("Microsoft YaHei UI", 12f * s, GraphicsUnit.Pixel);
            tabFontBold = new Font("Microsoft YaHei UI", 12f * s, FontStyle.Bold, GraphicsUnit.Pixel);
            badgeFont = new Font("Segoe UI", 10f * s, FontStyle.Bold, GraphicsUnit.Pixel);
            LayoutTabs();
            Invalidate();
        }

        // ---------------- 标签栏位置 ----------------

        bool IsFullscreen(IntPtr h)
        {
            if (h == IntPtr.Zero || h == Handle) return false;
            // 拖动中的窗口一定带标题栏，不可能是全屏；这段每帧都要取类名和样式，拖动时直接跳过
            if (mgr.DragGroup != null) return false;
            string cls = Native.GetClassNameOf(h);
            if (cls == "Progman" || cls == "WorkerW" || cls == "Shell_TrayWnd") return false;
            int style = Native.GetWindowLong(h, Native.GWL_STYLE);
            if ((style & Native.WS_CAPTION) == Native.WS_CAPTION) return false;
            Native.RECT r;
            if (!Native.GetWindowRect(h, out r)) return false;
            Rectangle m = Screen.FromHandle(h).Bounds;
            return r.Left <= m.Left && r.Top <= m.Top && r.Right >= m.Right && r.Bottom >= m.Bottom;
        }

        // 标题栏模式：把标签栏放进 Word 自己的标题栏，左侧避开快速访问工具栏，右侧避开窗口按钮
        bool TryTitleBar(IntPtr doc, out Rectangle rect)
        {
            rect = Rectangle.Empty;
            UpdateScale(doc);
            Native.RECT wr;
            if (!Native.GetWindowRect(doc, out wr)) return false;
            Native.RECT fr = Native.FrameBounds(doc, wr);
            // 拖窗口时 DWM 报的扩展边框会滞后一帧甚至停在原位，标签栏右缘就跟着卡住、边拖边露出 Word 标题栏。
            // GetWindowRect 永远是实时的，改用缓存的差值从它换算。
            bool moving = mgr.DragGroup != null;
            if (moving && frameRelDoc == doc)
            {
                fr.Left = wr.Left + frameRelL;
                fr.Top = wr.Top + frameRelT;
                fr.Right = wr.Right + frameRelR;
                fr.Bottom = wr.Bottom + frameRelB;
            }
            else
            {
                frameRelDoc = doc;
                frameRelL = fr.Left - wr.Left;
                frameRelT = fr.Top - wr.Top;
                frameRelR = fr.Right - wr.Right;
                frameRelB = fr.Bottom - wr.Bottom;
            }
            Native.RECT btn;
            bool haveBtn = Native.GetCaptionButtons(doc, out btn);
            // 同理：按钮矩形也存成“距窗口右边多远 / 比窗口顶低多少”，拖动时按当前窗口位置还原
            // Word 自绘标题栏，DWM 给的这个矩形经常半旧半新：横向还停在旧位置、纵向出现 Top > Bottom。
            // 只做横向检查会放过这种数据，标签栏右缘就会被钉死在一个旧的绝对 x 上。
            bool btnOk = false;
            if (haveBtn && btn.Right > btn.Left && btn.Bottom > btn.Top)
            {
                // Word 自绘标题栏时 DWM 会返回“窗口内相对坐标”（左上角=窗口左上角），
                // 这时 Top 会明显高于窗口顶边，用窗口矩形把它换算回屏幕坐标
                if (btn.Top < fr.Top - (int)Math.Round(4 * s) && btn.Top >= -2)
                {
                    btn.Left += wr.Left;
                    btn.Right += wr.Left;
                    btn.Top += wr.Top;
                    btn.Bottom += wr.Top;
                }
                int relLeft = fr.Right - btn.Left;
                // 既要落在标题栏范围内，右侧留白也要是“三枚按钮”的量级，否则就是脏数据
                btnOk = btn.Left > fr.Left && btn.Left < fr.Right
                    && btn.Top >= fr.Top - (int)Math.Round(4 * s)
                    && btn.Bottom > fr.Top
                    && btn.Bottom <= fr.Top + (int)Math.Round(110 * s)
                    && relLeft >= (int)Math.Round(60 * s)
                    && relLeft <= (int)Math.Round(200 * s);
            }
            if (btnOk)
            {
                btnRelDoc = doc;
                btnRelLeftFromRight = fr.Right - btn.Left;
                btnRelBottomFromTop = btn.Bottom - fr.Top;
                btnRelWidth = btn.Right - btn.Left;
            }
            else if (btnRelDoc == doc && btnRelWidth > 0
                && btnRelLeftFromRight >= (int)Math.Round(60 * s) && btnRelLeftFromRight <= (int)Math.Round(200 * s))
            {
                btn.Left = fr.Right - btnRelLeftFromRight;
                btn.Right = fr.Right;
                btn.Top = fr.Top;
                btn.Bottom = fr.Top + btnRelBottomFromTop;
                haveBtn = true;
            }
            if (moving && btnRelDoc == doc)
            {
                // 拖动期间完全不用 DWM 的值，否则右缘会停在旧位置
                btn.Left = fr.Right - btnRelLeftFromRight;
                btn.Right = fr.Right;
                btn.Top = fr.Top;
                btn.Bottom = fr.Top + btnRelBottomFromTop;
                haveBtn = true;
            }
            // 用 DWM 的窗口按钮矩形反推标题栏高度；拿不到时按系统标题栏高度估算
            Native.RECT tb;
            tb.Left = fr.Left;
            tb.Right = fr.Right;
            tb.Top = fr.Top;
            tb.Bottom = haveBtn && btn.Bottom > fr.Top + (int)Math.Round(10 * s)
                ? btn.Bottom
                : fr.Top + (int)Math.Round(32 * s);
            // Word 把账户头像画到标题栏下沿之外约 1.5 个逻辑像素（2 倍缩放下是 3 像素），
            // 标签栏只贴到按钮底边就会露出一条头像底边，最大化与还原状态都一样，统一多压一点
            tb.Bottom += Math.Max(2, (int)Math.Round(1.5 * s));
            int h = tb.Bottom - tb.Top;
            if (h < 16 || h > 140 * s) return false;
            // 窗口最上方是系统画的 1~2 像素高亮边，标签栏压上去会在两侧“断线”。
            // 量出它有几像素，把整条标签栏下移同样的高度，让系统自己把这条边画完。
            // 屏幕取色会等 DWM 合成，拖动时每帧取一次会把 8ms 的节拍拖垮、标签栏追不上窗口，所以只在不动时量，量到就缓存
            int topInset = topInsetValue;
            int nowTick = Environment.TickCount;
            if (mgr.DragGroup == null && (topInsetDoc != doc || topInsetSampleH != h || nowTick - topInsetTick > 500))
            {
                int v = 0;
                Color capProbe = Native.ScreenPixel(tb.Left + 2, tb.Top + h / 2);
                if (capProbe != Color.Empty)
                {
                    for (int k = 0; k < 4 && k < h - 12; k++)
                    {
                        Color px = Native.ScreenPixel(tb.Left + 2, tb.Top + k);
                        if (px == Color.Empty) break;
                        if (Math.Abs(px.R - capProbe.R) + Math.Abs(px.G - capProbe.G) + Math.Abs(px.B - capProbe.B) > 12) v = k + 1;
                        else break;
                    }
                }
                topInset = topInsetValue = v;
                topInsetDoc = doc;
                topInsetSampleH = h;
                topInsetTick = nowTick;
            }
            if (topInset > 0)
            {
                tb.Top += topInset;
                h -= topInset;
                if (h < 16) return false;
            }
            // 左边留给快速访问工具栏，右边收在最小化/最大化/关闭按钮之前
            int left = tb.Left + (int)Math.Round(150 * s);
            int right = tb.Right - (int)Math.Round(3 * s);
            if (haveBtn && btn.Left > tb.Left + (int)Math.Round(200 * s))
                right = Math.Min(right, btn.Left - (int)Math.Round(4 * s));
            else if (!haveBtn)
            {
                // 拿不到按钮矩形时按三枚按钮的宽度留白，宁可有缝也不盖住按钮。
                // 优先用这个窗口量到过的真实宽度（Word 的按钮比系统标准宽一些）
                int reserve = btnRelDoc == doc && btnRelWidth > 0 ? btnRelWidth : (int)Math.Round(146 * s);
                right = Math.Min(right, tb.Right - reserve);
            }
            // 拖动窗口时 DWM 的矩形可能滞后一帧，右侧再兜一次
            if (mgr.DragGroup != null)
            {
                int reserve = btnRelDoc == doc && btnRelWidth > 0 ? btnRelWidth : (int)Math.Round(146 * s);
                right = Math.Min(right, tb.Right - reserve);
            }
            if (right - left < (int)Math.Round(260 * s))
            {
                left = tb.Left + (int)Math.Round(70 * s);
                right = tb.Right - (int)Math.Round(70 * s);
            }
            if (right - left < (int)Math.Round(120 * s)) return false;
            if (barHeight != h)
            {
                barHeight = h;
                LayoutTabs();
            }
            // 跟随 Word 标题栏的浅色/深色主题：取快速访问工具栏左下的空白处
            int now = Environment.TickCount;
            // 只在缩放动画结束、且没在拖动时取色：动画中途 h/Top 都是插值，取到的可能是正文或功能区的底色，
            // 整条标签栏会闪一下变成那个颜色再跳回来
            bool stable = now - zoomFlipTick > 700;
            if (stable && mgr.DragGroup == null && now - titleSampleTick > 400 &&
                !(TitleLocked && TitleLockDoc == doc))
            {
                titleSampleTick = now;
                // 取标题栏竖直中线（快速访问工具栏左侧的空白处），贴着底边取容易碰到下面的功能区
                int sampleX = tb.Left + 4, sampleY = tb.Top + h / 2;
                // 采样点被别的窗口压住时，GetPixel 拿到的是那个窗口的颜色。先问一句这个点属于谁，
                // 不是本文档所在的顶层窗口就整帧跳过，保留上一次的颜色。
                IntPtr atPoint = Native.WindowFromPoint(new Native.POINT { X = sampleX, Y = sampleY });
                bool sameWindow = atPoint != IntPtr.Zero &&
                    Native.GetAncestor(atPoint, Native.GA_ROOT) == Native.GetAncestor(doc, Native.GA_ROOT);
                Color c = sameWindow ? Native.ScreenPixel(sampleX, sampleY) : Color.Empty;
                if (c != Color.Empty)
                {
                    bool dark = (c.R * 299 + c.G * 587 + c.B * 114) / 1000 < 128;
                    if (c == titlePending && dark == titlePendingDark) ++titlePendingCount;
                    else { titlePending = c; titlePendingDark = dark; titlePendingCount = 1; }
                    // 连读两次一致才认，避免刚好扫到一格过渡色
                    if (titlePendingCount >= 2)
                    {
                        TitleLocked = true;
                        TitleLockDoc = doc;
                        if (c != TitleCaption || dark != TitleDark)
                        {
                            Log.Write(string.Format("标题栏底色 #{0:X2}{1:X2}{2:X2}（{3}），本次运行不再采样",
                                c.R, c.G, c.B, dark ? "深色" : "浅色"));
                            TitleCaption = c;
                            TitleDark = dark;
                            Invalidate();
                        }
                    }
                }
            }
            int winW = wr.Right - wr.Left, winH = wr.Bottom - wr.Top;
            if (moving && relBarDoc == doc && relBarWinW == winW && relBarWinH == winH)
            {
                // 拖动时窗口尺寸没变，整条标签栏的形状也不会变：直接按缓存的相对位置跟着窗口走，
                // 免得每帧重算出的顶边/高度差一两像素，看起来像底边在闪
                rect = new Rectangle(wr.Left + relBarLeft, wr.Top + relBarTop, relBarW, relBarH);
                return true;
            }
            rect = new Rectangle(left, tb.Top, right - left, h);
            relBarDoc = doc;
            relBarLeft = rect.Left - wr.Left;
            relBarTop = rect.Top - wr.Top;
            relBarW = rect.Width;
            relBarH = rect.Height;
            relBarWinW = winW;
            relBarWinH = winH;
            return true;
        }

        public void UpdatePlacement()
        {
            if (destroying) return;
            IntPtr fg = Native.GetForegroundWindow();
            // 直接以前台文档为准定位，漏掉前台切换事件时 250ms 看门狗也能把标签栏拉回当前窗口
            IntPtr placeDoc = CurDoc;
            if (IndexOf(fg) >= 0) placeDoc = fg;
            if (placeDoc != CurDoc && !Native.IsIconic(placeDoc) && mgr.DragGroup == null && !pressOnBar)
            {
                CurDoc = placeDoc;
                LayoutTabs();
                Invalidate();
            }
            mgr.RestorePaneShift(placeDoc);
            // 最大化/还原动画期间 DWM 的矩形和按钮位置都在插值，这段时间量到的颜色和尺寸都不可信
            bool zoomedNow = Native.IsZoomed(placeDoc);
            if (zoomedNow != lastZoomed) { lastZoomed = zoomedNow; zoomFlipTick = Environment.TickCount; }
            BarMode m;
            Rectangle rect = Rectangle.Empty;
            bool overlay = OverlayOpen();
            overlayApplied = overlay;
            if (placeDoc == IntPtr.Zero || Tabs.Count == 0 || !Native.IsWindow(placeDoc) || Native.IsIconic(placeDoc) ||
                !Native.IsWindowVisible(placeDoc) || (fg == placeDoc && IsFullscreen(fg)) || overlay)
            {
                m = BarMode.Hidden;
            }
            // 拖动期间 Word 仍可能报告最大化，这时必须按窗口真实矩形跟随，否则标签栏会停在原处
            else if (Manager.TitleBarTabs && TryTitleBar(placeDoc, out rect))
            {
                m = BarMode.Follow;
            }
            else if (mgr.IsMax(placeDoc) && mgr.DragGroup != this)
            {
                UpdateScale(placeDoc);
                Rectangle wa = Screen.FromHandle(placeDoc).WorkingArea;
                m = BarMode.Maximized;
                int top = wa.Top;
                if (Manager.BelowRibbon)
                {
                    top = Native.BelowRibbonTop(placeDoc, wa.Top + Manager.BarHeightFor(placeDoc));
                    top = Math.Max(wa.Top, Math.Min(top, wa.Bottom - barHeight));
                }
                rect = new Rectangle(wa.Left, top, wa.Width, barHeight);
            }
            else
            {
                UpdateScale(placeDoc);
                Native.RECT wr;
                Native.GetWindowRect(placeDoc, out wr);
                Native.RECT f = Native.FrameBounds(placeDoc, wr);
                Rectangle wa = Screen.FromHandle(placeDoc).WorkingArea;
                int left = f.Left, right = f.Right;
                if (Manager.BelowRibbon)
                {
                    // 有导航窗格/编辑器等任务窗格时，只铺在正文视图上方，避免挡住窗格标题栏的关闭按钮
                    Native.RECT view;
                    if (Native.GetDocViewRect(placeDoc, out view))
                    {
                        int vl = Math.Max(f.Left, view.Left), vr = Math.Min(f.Right, view.Right);
                        if (vr - vl >= (int)Math.Round(120 * s)) { left = vl; right = vr; }
                    }
                }
                int w = Manager.BelowRibbon
                    ? Math.Max(1, right - left)
                    : Math.Max(1, Math.Min(f.Right - f.Left, DesiredWidth()));
                if (Manager.BelowRibbon)
                {
                    int fallback = Math.Max(wa.Top + Manager.BarHeightFor(placeDoc), f.Top + Manager.BarHeightFor(placeDoc));
                    int top = Native.BelowRibbonTop(placeDoc, fallback);
                    top = Math.Max(f.Top, Math.Min(top, f.Bottom - barHeight));
                    m = BarMode.Follow;
                    rect = new Rectangle(left, top, w, barHeight);
                }
                else
                {
                    // 向下压住标题栏顶部的边框线，让当前标签和标题栏连成一体
                    int overlap = (int)Math.Round(2 * s);
                    if (f.Top - barHeight + overlap >= wa.Top - (int)Math.Round(3 * s))
                    {
                        m = BarMode.Follow;
                        rect = new Rectangle(f.Left, Math.Max(wa.Top, f.Top - barHeight + overlap), w, barHeight);
                    }
                    else
                    {
                        m = BarMode.AutoHide;
                        rect = new Rectangle(f.Left, f.Top, w, barHeight);
                    }
                }
            }
            if (m != BarMode.AutoHide)
            {
                overlayShown = false;
                hoverTimer.Stop();
            }
            else if (!hoverTimer.Enabled) { hoverTicks = 0; leaveTicks = 0; hoverTimer.Start(); }

            bool modeChanged = m != mode;
            mode = m;
            SyncTransparency(m);
            if (m == BarMode.Hidden || (m == BarMode.AutoHide && !overlayShown))
            {
                if (Visible) Hide();
                lastRect = Rectangle.Empty;
                return;
            }

            // 本组文档在前台时置顶，Office 被点击激活也盖不住标签栏；切到其他程序后取消置顶，只紧贴在文档上面一层
            bool docInFront = IndexOf(fg) >= 0 || fg == Handle;
            if (docInFront) lastInFrontTick = Environment.TickCount;
            // 从任务栏缩略图 / Alt+Tab 进 Word 时，前台事件会先于系统把窗口提到顶层到达。
            // 这段时间必须一直保持置顶，否则标签栏会被 Word 盖住，点标签自然就没反应了。
            bool recentlyInFront = Environment.TickCount - lastInFrontTick < 1500;
            bool wantTopmost = m == BarMode.AutoHide || docInFront || recentlyInFront;
            bool actuallyTopmost = (Native.GetWindowLong(Handle, Native.GWL_EXSTYLE) & Native.WS_EX_TOPMOST) != 0;
            bool moved = rect != lastRect || !Visible;
            if (!Visible) Show();
            // 采样标签栏下方的正文画布色，让标签栏和页面融为一体
            if (Manager.BelowRibbon && docInFront) mgr.SampleCanvas(CurDoc, rect);
            const uint moveFlags = Native.SWP_NOACTIVATE | Native.SWP_NOZORDER;
            if (wantTopmost)
            {
                // 已经置顶时只移动位置，不重复调整 Z 序（拖动窗口时每帧都会走到这里）
                if (!isTopmost || !actuallyTopmost)
                    Native.SetWindowPos(Handle, Native.HWND_TOPMOST, rect.Left, rect.Top, rect.Width, rect.Height, Native.SWP_NOACTIVATE);
                else if (reassertZ)
                    Native.SetWindowPos(Handle, Native.HWND_TOPMOST, 0, 0, 0, 0,
                        Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
                else if (moved)
                    Native.SetWindowPos(Handle, IntPtr.Zero, rect.Left, rect.Top, rect.Width, rect.Height, moveFlags);
            }
            else
            {
                if (isTopmost)
                    Native.SetWindowPos(Handle, Native.HWND_NOTOPMOST, 0, 0, 0, 0,
                        Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
                // 本组文档不在前台：把标签栏压到文档后面并对齐位置。
                // 以前这里是“紧贴文档上面一层”，会停在非置顶层最上面，盖住资源管理器等其它窗口。
                if (CurDoc != IntPtr.Zero && Native.IsWindow(CurDoc))
                    Native.SetWindowPos(Handle, CurDoc, rect.Left, rect.Top, rect.Width, rect.Height,
                        Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
                else if (moved)
                    Native.SetWindowPos(Handle, IntPtr.Zero, rect.Left, rect.Top, rect.Width, rect.Height, moveFlags);
            }
            isTopmost = wantTopmost;
            reassertZ = false;
            lastRect = rect;
            if (modeChanged) { LayoutTabs(); Invalidate(); }
        }

        int DesiredWidth()
        {
            int n = Math.Max(1, Tabs.Count);
            return (int)Math.Ceiling(12 * s + n * 210 * s + (n - 1) * 2 * s);
        }

        bool Floating
        {
            get { return !Manager.BelowRibbon && !Manager.TitleBarTabs && (mode == BarMode.Follow || mode == BarMode.AutoHide); }
        }

        void SyncTransparency(BarMode m)
        {
            Color want = (!Manager.BelowRibbon && !Manager.TitleBarTabs && (m == BarMode.Follow || m == BarMode.AutoHide)) ? KeyColor : Color.Empty;
            if (TransparencyKey != want) TransparencyKey = want;
        }

        // 小窗口且上方没空间时：鼠标停在窗口上沿约 0.2 秒滑出标签栏，离开 0.5 秒后收起
        void OnHoverTick()
        {
            if (mode != BarMode.AutoHide || CurDoc == IntPtr.Zero) { hoverTimer.Stop(); return; }
            Point p = Cursor.Position;
            if (!overlayShown)
            {
                Native.RECT wr;
                Native.GetWindowRect(CurDoc, out wr);
                Native.RECT f = Native.FrameBounds(CurDoc, wr);
                int zone = (int)(4 * s);
                bool inZone = p.X >= f.Left && p.X < f.Right && p.Y >= f.Top - zone && p.Y <= f.Top + zone &&
                              Control.MouseButtons == MouseButtons.None;
                hoverTicks = inZone ? hoverTicks + 1 : 0;
                if (hoverTicks >= 2) { overlayShown = true; leaveTicks = 0; UpdatePlacement(); }
            }
            else
            {
                Rectangle r = Bounds;
                r.Inflate((int)(6 * s), (int)(6 * s));
                leaveTicks = (r.Contains(p) || pressOnBar) ? 0 : leaveTicks + 1;
                if (leaveTicks >= 5) { overlayShown = false; hoverTicks = 0; UpdatePlacement(); }
            }
        }

        // ---------------- 绘制 ----------------

        public void LayoutTabs()
        {
            float pad = Floating || Manager.BelowRibbon || Manager.TitleBarTabs ? 6 * s : 4 * s, gap = Floating || Manager.BelowRibbon || Manager.TitleBarTabs ? 2 * s : 1 * s, top = Manager.BelowRibbon || Manager.TitleBarTabs ? 0 : 4 * s;
            float h = barHeight - top;
            int n = Tabs.Count;
            if (n == 0) return;
            float avail = ClientSize.Width - pad * 2;
            float w = Math.Max(1, Math.Min(190 * s, (avail - gap * (n - 1)) / n));
            float x = pad;
            foreach (var t in Tabs)
            {
                t.Bounds = new RectangleF(x, top, w, h);
                float cs = 16 * s;
                t.CloseBounds = new RectangleF(x + w - cs - 6 * s, top + (h - cs) / 2, cs, cs);
                x += w + gap;
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            LayoutTabs();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            bool floating = Floating;
            if (floating)
            {
                g.Clear(KeyColor);
            }
            else if (Manager.TitleBarTabs)
            {
                using (var br = new SolidBrush(TitleCaption))
                    g.FillRectangle(br, ClientRectangle);
            }
            else if (Manager.BelowRibbon)
            {
                using (var br = new SolidBrush(BarColor))
                    g.FillRectangle(br, ClientRectangle);
            }
            else
            {
                using (var br = new LinearGradientBrush(ClientRectangle, BarBg, BarBg2, LinearGradientMode.Vertical))
                    g.FillRectangle(br, ClientRectangle);
            }
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

            for (int i = 0; i < Tabs.Count; i++)
            {
                var t = Tabs[i];
                bool active = t.Handle == CurDoc;
                bool hover = i == hoverIndex;
                var b = t.Bounds;

                if (Manager.TitleBarTabs)
                {
                    // 标题栏内：白底页签，活动页签底部一条蓝色指示线，和 Word 自己的功能区选中态一致
                    if (active || hover)
                    {
                        using (var path = TopRounded(b, 4 * s))
                        {
                            using (var br = new SolidBrush(active ? ActiveBg : HoverBg))
                                g.FillPath(br, path);
                            using (var pen = new Pen(active ? Color.FromArgb(16, 0, 0, 0) : Color.FromArgb(10, 0, 0, 0), Math.Max(1f, s)))
                                g.DrawPath(pen, path);
                        }
                    }
                    if (active)
                    {
                        float ay = b.Bottom - Math.Max(1.2f, 1.1f * s);
                        using (var pen = new Pen(Color.FromArgb(76, 148, 255), Math.Max(1.6f, 1.5f * s)))
                            g.DrawLine(pen, b.Left + 6 * s, ay, b.Right - 6 * s, ay);
                    }
                }
                else if (floating || (!Manager.BelowRibbon && (active || hover)))
                {
                    using (var path = Manager.BelowRibbon ? BottomRounded(b, 5 * s) : TopRounded(b, 5 * s))
                    {
                        if (!floating && active && !Manager.BelowRibbon)
                        {
                            var sb = new RectangleF(b.X, b.Y + 1.5f * s, b.Width, b.Height);
                            using (var sp = Manager.BelowRibbon ? BottomRounded(sb, 5 * s) : TopRounded(sb, 5 * s))
                            using (var shadow = new SolidBrush(Color.FromArgb(55, 0, 0, 0)))
                                g.FillPath(shadow, sp);
                        }
                        using (var br = new SolidBrush(active ? ActiveBg : hover ? HoverBg : BarBg))
                            g.FillPath(br, path);
                        if (!floating && !Manager.BelowRibbon && (active || hover))
                        {
                            using (var pen = new Pen(active ? Color.FromArgb(88, 91, 97) : Color.FromArgb(72, 75, 81), Math.Max(1f, s)))
                                g.DrawPath(pen, path);
                        }
                    }
                }
                if (!floating && Manager.BelowRibbon && (active || hover))
                {
                    // 活动标签：一层很淡的提亮/压暗 + 顶部蓝线，既能一眼看出，又不会把整条切断
                    using (var path = Rounded(b, 4 * s))
                    using (var br = new SolidBrush(active
                        ? (LightCanvas ? Color.FromArgb(26, 0, 0, 0) : Color.FromArgb(40, 255, 255, 255))
                        : (LightCanvas ? Color.FromArgb(12, 0, 0, 0) : Color.FromArgb(14, 255, 255, 255))))
                        g.FillPath(br, path);
                }

                else if (!active && !hover && i + 1 < Tabs.Count && Tabs[i + 1].Handle != CurDoc && i + 1 != hoverIndex)
                {
                    using (var pen = new Pen(SepColor, Math.Max(1f, s)))
                        g.DrawLine(pen, b.Right, b.Top + b.Height * 0.30f, b.Right, b.Bottom - b.Height * 0.30f);
                }

                // 应用标识：浅色底 + 同色字母，比纯色块更轻
                float bw = (Manager.TitleBarTabs ? 13 : 15) * s, bh = bw;
                var badge = new RectangleF(b.Left + 8 * s, b.Top + (b.Height - bh) / 2, bw, bh);
                Color kc = KindColor(t.Kind);
                using (var path = Rounded(badge, 4 * s))
                {
                    using (var br = new SolidBrush(Color.FromArgb(active ? 42 : 28, kc)))
                        g.FillPath(br, path);
                    using (var pen = new Pen(Color.FromArgb(active ? 120 : 70, kc), Math.Max(1f, s * 0.8f)))
                        g.DrawPath(pen, path);
                }
                using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                using (var tb = new SolidBrush(kc))
                    g.DrawString(t.Kind == DocKind.Word ? "W" : "X", badgeFont, tb, badge, sf);

                // 标题
                bool showClose = active || hover;
                bool dirtyDot = t.Dirty && (Manager.TitleBarTabs || Manager.BelowRibbon);
                float textLeft = badge.Right + 7 * s;
                float textRight = showClose
                    ? t.CloseBounds.Left - (dirtyDot ? 16 : 2) * s
                    : b.Right - (dirtyDot ? 20 : 10) * s;
                var textRect = Rectangle.Round(new RectangleF(textLeft, b.Top, Math.Max(0, textRight - textLeft), b.Height));
                TextRenderer.DrawText(g, t.Title, active ? tabFontBold : tabFont, textRect, active ? TextColor : hover ? HoverText : DimText,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine |
                    TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);

                // 关闭按钮
                if (showClose)
                {
                    var c = t.CloseBounds;
                    if (hover && hoverClose)
                        using (var br = new SolidBrush(LightCanvas ? Color.FromArgb(26, 0, 0, 0) : Color.FromArgb(30, 255, 255, 255))) g.FillEllipse(br, c);
                    float d = 3.6f * s;
                    float cx = c.Left + c.Width / 2, cy = c.Top + c.Height / 2;
                    using (var pen = new Pen(hover ? TextColor : DimText, 1.15f * s))
                    {
                        g.DrawLine(pen, cx - d, cy - d, cx + d, cy + d);
                        g.DrawLine(pen, cx + d, cy - d, cx - d, cy + d);
                    }
                }

                // 未保存：标题栏/功能区下方模式用右侧小圆点，避免和活动标签的蓝色下划线打架
                if (t.Dirty)
                {
                    if (Manager.TitleBarTabs || Manager.BelowRibbon)
                    {
                        float d = 5 * s;
                        float dx = (active || hover) ? t.CloseBounds.Left - 16 * s : b.Right - 15 * s;
                        var dot = new RectangleF(dx, b.Top + (b.Height - d) / 2, d, d);
                        using (var db = new SolidBrush(Color.FromArgb(230, 162, 60)))
                            g.FillEllipse(db, dot);
                    }
                    else
                    {
                        float dw = Math.Min(b.Width - 24 * s, 30 * s);
                        if (dw > 6 * s)
                        {
                            var dr = new RectangleF(b.Left + (b.Width - dw) / 2, b.Bottom - 5 * s, dw, 2.4f * s);
                            using (var dp = Rounded(dr, 1.2f * s))
                            using (var db = new SolidBrush(Color.FromArgb(74, 144, 255)))
                                g.FillPath(db, dp);
                        }
                    }
                }
            }

            if (floating) return;
            // 顶部模式在底部压一条暗线；功能区下方模式要和正文融成一片，不能再画底边
            if (!Manager.BelowRibbon && !Manager.TitleBarTabs)
            {
                using (var pen = new Pen(Color.FromArgb(20, 21, 23), Math.Max(1f, s)))
                {
                    float yb = ClientSize.Height - pen.Width / 2;
                    g.DrawLine(pen, 0, yb, ClientSize.Width, yb);
                }
            }
            using (var pen = new Pen(Manager.TitleBarTabs ? TitleCaption : Manager.BelowRibbon ? Blend(Manager.CanvasColor, -32) : BarLine, Math.Max(1f, s)))
            {
                float y = Manager.BelowRibbon ? pen.Width / 2 : ClientSize.Height - pen.Width / 2;
                int ai = IndexOf(CurDoc);
                if (ai < 0) g.DrawLine(pen, 0, y, ClientSize.Width, y);
                else
                {
                    var ab = Tabs[ai].Bounds;
                    g.DrawLine(pen, 0, y, ab.Left, y);
                    g.DrawLine(pen, ab.Right, y, ClientSize.Width, y);
                }
            }
        }

        static Color KindColor(DocKind k)
        {
            return k == DocKind.Word ? Color.FromArgb(88, 156, 255) : Color.FromArgb(78, 201, 137);
        }

        static GraphicsPath TopRounded(RectangleF r, float rad)
        {
            var p = new GraphicsPath();
            float d = rad * 2;
            p.AddLine(r.Left, r.Bottom, r.Left, r.Top + rad);
            p.AddArc(r.Left, r.Top, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Top, d, d, 270, 90);
            p.AddLine(r.Right, r.Top + rad, r.Right, r.Bottom);
            p.CloseFigure();
            return p;
        }

        static GraphicsPath BottomRounded(RectangleF r, float rad)
        {
            var p = new GraphicsPath();
            float d = rad * 2;
            p.AddLine(r.Left, r.Top, r.Left, r.Bottom - rad);
            p.AddArc(r.Left, r.Bottom - d, d, d, 180, -90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 90, -90);
            p.AddLine(r.Right, r.Bottom - rad, r.Right, r.Top);
            p.CloseFigure();
            return p;
        }
        static GraphicsPath Rounded(RectangleF r, float rad)
        {
            var p = new GraphicsPath();
            float d = rad * 2;
            p.AddArc(r.Left, r.Top, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Top, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        // ---------------- 鼠标交互 ----------------

        int HitTest(Point p)
        {
            for (int i = 0; i < Tabs.Count; i++) if (Tabs[i].Bounds.Contains(p)) return i;
            return -1;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (TryTitleBarDrag()) return;
            // 正在从这个标签栏拖动窗口：鼠标每动一下就跟一下
            if (mgr.DragSource == this) { mgr.DragMove(); return; }
            if (pressOnBar && e.Button == MouseButtons.Left && CheckStartDrag(Cursor.Position)) return;
            if (pressOnBar && pressIndex >= 0 && e.Button == MouseButtons.Left)
            {
                // 水平拖动标签：调整顺序
                if (!dragging && Math.Abs(e.X - pressPoint.X) > 6 * s) dragging = true;
                if (dragging)
                {
                    int target = HitTest(new Point(e.X, (int)Tabs[pressIndex].Bounds.Top + 2));
                    if (target >= 0 && target != pressIndex)
                    {
                        var t = Tabs[pressIndex];
                        Tabs.RemoveAt(pressIndex);
                        Tabs.Insert(target, t);
                        pressIndex = target;
                        hoverIndex = target;
                        LayoutTabs();
                        Invalidate();
                    }
                    return;
                }
            }
            int idx = HitTest(e.Location);
            bool overClose = idx >= 0 && Tabs[idx].CloseBounds.Contains(e.Location);
            if (idx != hoverIndex || overClose != hoverClose)
            {
                if (idx != hoverIndex) tip.SetToolTip(this, idx >= 0 ? Tabs[idx].FullTitle : null);
                hoverIndex = idx;
                hoverClose = overClose;
                Invalidate();
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (!pressOnBar) { hoverIndex = -1; hoverClose = false; Invalidate(); }
        }

        // 标题栏空白区拖动：直接走本程序自己的移动逻辑。
        // 以前是把 WM_NCLBUTTONDOWN(HTCAPTION) 转给 Word 走系统原生拖动，但 Office 的自绘标题栏
        // 不接受这种合成消息，实测窗口完全不动；这里统一用与其它布局相同的拖动实现。
        bool TryTitleBarDrag()
        {
            if (!Manager.TitleBarTabs || !pressOnBar || pressIndex >= 0 || CurDoc == IntPtr.Zero) return false;
            Point c = Cursor.Position;
            if (Math.Abs(c.X - pressScreen.X) <= 4 * s && Math.Abs(c.Y - pressScreen.Y) <= 4 * s) return false;
            BeginBarMove(c);
            return true;
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            int idx = HitTest(e.Location);
            if (e.Button == MouseButtons.Left)
            {
                pressOnBar = true;
                pressScreen = Cursor.Position;
                pressIndex = idx;
                pressPoint = e.Location;
                dragging = false;
                pressTimer.Start();
            }
            else if (e.Button == MouseButtons.Middle && idx >= 0) CloseTab(Tabs[idx]);
            else if (e.Button == MouseButtons.Right) ShowTabMenu(idx, e.Location);
        }

        // 双击标签：关闭该标签；双击标签栏空白处：最大化 / 还原，与双击标题栏一致
        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            if (e.Button != MouseButtons.Left) return;
            if (Environment.TickCount - lastCloseTick < 500) return;
            int idx = HitTest(e.Location);
            if (idx >= 0)
            {
                if (!Tabs[idx].CloseBounds.Contains(e.Location)) CloseTab(Tabs[idx]);
                return;
            }
            if (IndexOf(CurDoc) < 0) return;
            ActivateDoc(CurDoc);
            if (Native.IsZoomed(CurDoc)) Native.ShowWindow(CurDoc, Native.SW_RESTORE);
            else Native.ShowWindow(CurDoc, Native.SW_MAXIMIZE);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button != MouseButtons.Left) return;
            if (mgr.DragSource == this) { mgr.EndDrag(); return; }
            bool wasPress = pressOnBar;
            int pressed = pressIndex;
            bool wasDrag = dragging;
            ResetPress();
            if (!wasPress || wasDrag || pressed < 0 || pressed >= Tabs.Count) return;
            var t = Tabs[pressed];
            if (!t.Bounds.Contains(e.Location)) return;
            if (t.CloseBounds.Contains(e.Location) && (t.Handle == CurDoc || pressed == hoverIndex)) CloseTab(t);
            else ActivateDoc(t.Handle);
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            if (Tabs.Count == 0) return;
            int cur = IndexOf(CurDoc);
            int next = e.Delta > 0 ? cur - 1 : cur + 1;
            if (cur < 0) next = 0;
            next = (next + Tabs.Count) % Tabs.Count;
            ActivateDoc(Tabs[next].Handle);
        }

        void ShowTabMenu(int idx, Point p)
        {
            var items = new List<MenuItem>();
            if (idx >= 0)
            {
                var t = Tabs[idx];
                items.Add(new MenuItem("关闭", delegate { CloseTab(t); }));
                items.Add(new MenuItem("关闭其他标签", delegate
                {
                    var rest = new List<IntPtr>();
                    foreach (var o in Tabs) if (o != t) rest.Add(o.Handle);
                    mgr.StartCascade(rest);
                }));
                items.Add(new MenuItem("关闭全部", delegate
                {
                    var all = new List<IntPtr>();
                    foreach (var o in Tabs) if (o != t) all.Add(o.Handle);
                    all.Add(t.Handle);
                    mgr.StartCascade(all);
                }));
                if (Tabs.Count > 1)
                    items.Add(new MenuItem("拆分为单独窗口", delegate { int i = IndexOf(t.Handle); if (i >= 0) DetachToSide(i); }));
                if (mgr.Groups.Count > 1)
                    items.Add(new MenuItem("合并所有窗口", delegate { MergeAllInto(); }));
                items.Add(new MenuItem("-"));
            }
            items.Add(new MenuItem(Manager.BelowRibbon ? "标签栏移回顶部" : "标签栏放在功能区下方", delegate { mgr.ToggleBelowRibbon(); }));
            items.Add(new MenuItem("退出 DocTabs", delegate { mgr.ExitApp(true); }));
            var menu = new ContextMenu(items.ToArray());
            Native.SetForegroundWindow(Handle);
            menu.Show(this, p);
        }

        // ---------------- 拖动：拆分标签 / 移动窗口 ----------------

        // 按住标签栏期间每 10ms 一次：判断是否开始拖动、发现在别处松手
        void PressTick()
        {
            if (!pressOnBar) { pressTimer.Stop(); return; }
            // 读物理按键状态：在别处松手时消息不会发给标签栏
            int vk = SystemInformation.MouseButtonsSwapped ? Native.VK_RBUTTON : Native.VK_LBUTTON;
            bool down = (Native.GetAsyncKeyState(vk) & 0x8000) != 0;
            Point c = Cursor.Position;
            if (!down)
            {
                // 在标签栏上松手的交给 OnMouseUp 处理（点击切换）；在别处松手的直接收尾
                if (!Bounds.Contains(c)) ResetPress();
                return;
            }
            if (TryTitleBarDrag()) return;
            CheckStartDrag(c);
        }

        // 多个标签时把标签竖直拖出标签栏：拆成单独的窗口；水平拖动标签：调整顺序（在 OnMouseMove 里）。
        // 拖空白处、或只有一个标签时拖标签：移动窗口
        bool CheckStartDrag(Point c)
        {
            if (!pressOnBar) return false;
            int dx = Math.Abs(c.X - pressScreen.X), dy = Math.Abs(c.Y - pressScreen.Y);
            if (pressIndex >= 0 && pressIndex < Tabs.Count && Tabs.Count > 1)
            {
                bool detach = dragging ? dy > barHeight * 0.8f : dy > barHeight * 0.6f && dy > dx;
                if (!detach) return false;
                Detach(pressIndex, c);
                return true;
            }
            if (Math.Max(dx, dy) <= 4 * s) return false;
            BeginBarMove(c);
            return true;
        }

        void ResetPress()
        {
            pressOnBar = false;
            pressIndex = -1;
            dragging = false;
            pressTimer.Stop();
        }

        // 移动整个窗口：最大化时先在鼠标下还原，然后跟着鼠标走
        void BeginBarMove(Point c)
        {
            IntPtr h = pressIndex >= 0 && pressIndex < Tabs.Count ? Tabs[pressIndex].Handle : CurDoc;
            ResetPress();
            if (IndexOf(h) < 0 || Native.IsIconic(h)) return;
            hoverIndex = -1;
            hoverClose = false;
            tip.Hide(this);
            // 按着鼠标时不切前台（见 BeginDrag），只更新分组状态，松手后再切
            if (h != CurDoc) { ApplyGroupState(h, true); CurDoc = h; }
            mgr.Active = this;
            Invalidate();
            if (mgr.IsMax(h)) mgr.BeginDrag(h, this, RestoreUnderCursor(h, c, pressScreen.X - Left, pressScreen.Y - Top));
            else mgr.BeginDrag(h, this);
            Capture = true;
        }

        // 从最大化状态拖出：还原成普通大小，并让鼠标仍落在标签栏里原来抓住的位置；返回窗口左上角
        Point RestoreUnderCursor(IntPtr h, Point c, int grabX, int grabY)
        {
            var wp = Native.GetPlacement(h);
            Native.RECT n = HasState ? GroupNormal : wp.rcNormalPosition;
            int w = Math.Max(200, n.Right - n.Left), ht = Math.Max(150, n.Bottom - n.Top);
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            Native.RECT f = Native.FrameBounds(h, r);
            int side = f.Left - r.Left; // 左侧隐形边框，最大化和普通状态下一样宽
            int frameW = w - side * 2;
            Rectangle wa = Screen.FromPoint(c).WorkingArea;
            if (grabX > frameW - 60 * s) grabX = (int)((long)frameW * (c.X - wa.Left) / Math.Max(1, wa.Width));
            return PlaceNormal(h, this, c.X - grabX - side, c.Y - grabY, w, ht, c);
        }

        // 把窗口 h 以普通状态放到屏幕坐标 (left, barTop + 标签栏高度) 处，并记为分组 g 的还原位置；返回窗口左上角
        Point PlaceNormal(IntPtr h, TabBar g, int left, int barTop, int w, int ht, Point c)
        {
            Screen scr = Screen.FromPoint(c);
            Rectangle wa = scr.WorkingArea, mon = scr.Bounds;
            int overlap = (int)Math.Round(2 * s);
            int top = Math.Max(wa.Top + barHeight - overlap, barTop + barHeight - overlap);
            // rcNormalPosition 用的是工作区坐标（相对于工作区左上角）
            int ox = wa.Left - mon.Left, oy = wa.Top - mon.Top;
            var n = new Native.RECT { Left = left - ox, Top = top - oy, Right = left - ox + w, Bottom = top - oy + ht };
            mgr.Shifted.Remove(h);
            g.GroupMax = false;
            g.GroupNormal = n;
            g.HasState = true;
            // 从最大化拖出时，同组其他窗口一起回到普通状态，避免背景留一个最大化窗口
            SyncSiblingsOutOfMax(h, g, n);
            // 拖动刚开始，不能让标签栏等 Office：本来就是普通窗口的，异步挪过去；最大化/最小化的才需要同步还原。
            // 都不激活（按着鼠标时激活，Office 会等到松手才处理后面的拖动）
            if (!Native.IsZoomed(h) && !Native.IsIconic(h))
            {
                Native.SetWindowPos(h, IntPtr.Zero, left, top, w, ht,
                    Native.SWP_NOZORDER | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER | Native.SWP_ASYNCWINDOWPOS);
                return new Point(left, top);
            }
            var wp = Native.GetPlacement(h);
            wp.flags = 0;
            wp.showCmd = Native.SW_SHOWNOACTIVATE;
            wp.rcNormalPosition = n;
            Native.SetWindowPlacement(h, ref wp);
            return new Point(left, top);
        }

        // 把第 index 个标签拆成单独一组，窗口出现在鼠标下面并继续跟着鼠标走
        void Detach(int index, Point c)
        {
            var tab = Tabs[index];
            IntPtr h = tab.Handle;
            float grabX = Math.Max(10 * s, Math.Min(pressPoint.X - tab.Bounds.Left, 180 * s));
            int grabY = Math.Max(0, Math.Min(pressPoint.Y, barHeight - 1));
            ResetPress();
            hoverIndex = -1;
            hoverClose = false;
            tip.Hide(this);

            int w, ht;
            WindowSizeFor(h, out w, out ht);
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            Native.RECT f = Native.FrameBounds(h, r);
            int side = f.Left - r.Left;

            var ng = SplitOff(index);
            // 新的标签栏只有一个标签，它的左边距是 6 像素（按缩放）
            int frameLeft = c.X - (int)(6 * s + grabX);
            Point pos = PlaceNormal(h, ng, frameLeft - side, c.Y - grabY, w, ht, c);
            // 不在这里切前台（见 BeginDrag），松手后再切；鼠标仍由本标签栏捕获，拖动消息照样发到这里
            mgr.BeginDrag(h, this, pos);
            ng.UpdatePlacement();
            Capture = true;
        }

        // 右键菜单“拆分为单独窗口”：在原窗口旁边错开一点放
        void DetachToSide(int index)
        {
            var tab = Tabs[index];
            IntPtr h = tab.Handle;
            int w, ht;
            WindowSizeFor(h, out w, out ht);
            Rectangle wa = Screen.FromHandle(CurDoc != IntPtr.Zero ? CurDoc : h).WorkingArea;
            int off = (int)(40 * s);
            int left = GroupMax ? wa.Left + wa.Width / 4 : 0, barTop = 0;
            if (!GroupMax && HasState)
            {
                Native.RECT cr;
                Native.GetWindowRect(CurDoc, out cr);
                left = cr.Left + off;
                barTop = cr.Top - barHeight + off;
            }
            else barTop = wa.Top + wa.Height / 6;
            var ng = SplitOff(index);
            PlaceNormal(h, ng, left, barTop, w, ht, new Point(left + w / 2, barTop + barHeight));
            Manager.BringToFront(h);
            ng.UpdatePlacement();
        }

        // 从本组移出第 index 个标签，放进新建的一组
        TabBar SplitOff(int index)
        {
            var tab = Tabs[index];
            Tabs.RemoveAt(index);
            if (CurDoc == tab.Handle) CurDoc = Tabs[Math.Min(index, Tabs.Count - 1)].Handle;
            LayoutTabs();
            Invalidate();
            var ng = mgr.NewGroup();
            ng.Tabs.Add(tab);
            ng.CurDoc = tab.Handle;
            mgr.Active = ng;
            ng.LayoutTabs();
            UpdatePlacement();
            return ng;
        }

        // 拆出去的窗口用原分组还原时的大小
        void WindowSizeFor(IntPtr h, out int w, out int ht)
        {
            Native.RECT n;
            if (HasState) n = GroupNormal;
            else n = Native.GetPlacement(h).rcNormalPosition;
            w = Math.Max((int)(400 * s), n.Right - n.Left);
            ht = Math.Max((int)(300 * s), n.Bottom - n.Top);
        }

        void MergeAllInto()
        {
            foreach (var g in mgr.Groups.ToArray()) if (g != this) mgr.Merge(g, this);
        }

        // ---------------- 窗口操作 ----------------

        public void ActivateDoc(IntPtr h)
        {
            if (!Native.IsWindow(h)) return;
            // 先把标签栏压到最上层，再抬 Word 窗口，避免切换瞬间标签栏被新窗口盖一下
            lastInFrontTick = Environment.TickCount;
            ForceFront();
            if (Native.IsIconic(h)) Native.ShowWindow(h, Native.SW_RESTORE);
            // 先改当前文档，再同步状态：同步引起的 LOCATIONCHANGE 会算到这个文档头上
            CurDoc = h;
            mgr.Active = this;
            // 先把目标文档调成与当前文档相同的状态/位置，再切过去，看起来就像同一个窗口换了内容
            ApplyGroupState(h, true);
            bool ok = Manager.BringToFront(h);
            Log.Write(string.Format("切换标签 0x{0:X} ok={1} fg=0x{2:X} max={3}",
                h.ToInt64(), ok, Native.GetForegroundWindow().ToInt64(), GroupMax));
            if (!ok)
            {
                // 前台锁偶尔会晚一点放开：短时间重试，不要直接放弃
                retryDoc = h;
                retryTicks = 3;
                retryTimer.Stop();
                retryTimer.Start();
            }
            UpdatePlacement();
            Invalidate();
        }

        void CloseTab(DocTab t)
        {
            lastCloseTick = Environment.TickCount;
            // Office 会自行弹出“是否保存”提示
            Manager.CloseDoc(t.Handle);
            mgr.ScheduleScan();
        }
    }

    // 后台读取 Word / Excel 的文档“未保存”状态。只读访问，不修改任何文档。
    // 放在独立 STA 线程里，Office 忙或弹窗时不会卡住标签栏。
    class DirtyWatcher
    {
        readonly Thread thread;
        readonly Dictionary<IntPtr, bool> map = new Dictionary<IntPtr, bool>();
        volatile bool running = true;
        int version;
        int lastDiagnostic;
        // 当前要跟踪的文档窗口。由 UI 线程在每次 Scan 后整体替换，读者只取一次引用。
        Target[] targets = new Target[0];

        struct Target { public IntPtr Handle; public DocKind Kind; }

        public int Version { get { return version; } }

        public void SetTargets(System.Collections.Generic.List<DocTab> docs)
        {
            var list = new Target[docs.Count];
            for (int i = 0; i < docs.Count; i++) list[i] = new Target { Handle = docs[i].Handle, Kind = docs[i].Kind };
            targets = list; // 引用赋值是原子的，读线程不需要锁
        }

        public DirtyWatcher()
        {
            thread = new Thread(Loop);
            thread.IsBackground = true;
            thread.Name = "DocTabs dirty watcher";
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        public void Stop() { running = false; }

        public Dictionary<IntPtr, bool> Snapshot()
        {
            lock (map) return new Dictionary<IntPtr, bool>(map);
        }

        void Loop()
        {
            Native.CoInitializeEx(IntPtr.Zero, 2 /* COINIT_APARTMENTTHREADED */);
            try
            {
                while (running)
                {
                    int t0 = Environment.TickCount;
                    Dictionary<IntPtr, bool> next = null;
                    try { next = Query(); }
                    catch { }
                    if (next != null && running)
                    {
                        bool changed;
                        Dictionary<IntPtr, bool> before;
                        lock (map)
                        {
                            before = new Dictionary<IntPtr, bool>(map);
                            changed = !Same(map, next);
                            if (changed)
                            {
                                map.Clear();
                                foreach (var kv in next) map[kv.Key] = kv.Value;
                            }
                        }
                        if (changed)
                        {
                            Interlocked.Increment(ref version);
                            Log.Write("未保存标记 " + Describe(before, next));
                        }
                    }
                    // 插件上报走注册表，读一次很便宜；上一轮要是卡在旧式 COM 调用上，就多歇两秒
                    int extra = unchecked(Environment.TickCount - t0) > 1500 ? 20 : 0;
                    for (int i = 0; i < 10 + extra && running; i++) Thread.Sleep(100);
                }
            }
            finally { Native.CoUninitialize(); }
        }

        static readonly Guid IID_Dispatch = new Guid("00020400-0000-0000-C000-000000000046");
        const string DirtyKey = @"Software\DocTabs\Dirty";

        struct PluginState { public int Pid; public bool Dirty; }

        Dictionary<IntPtr, bool> Query()
        {
            Dictionary<IntPtr, bool> prev;
            lock (map) prev = new Dictionary<IntPtr, bool>(map);
            var result = new Dictionary<IntPtr, bool>();
            var list = targets;
            int wordHits = 0, excelHits = 0, wordTargets = 0, excelTargets = 0;
            string firstError = null;
            // Office 插件在文档进程内读取 Saved，结果放在注册表里。它比跨进程 COM 可靠：
            // 这台机器上 AccessibleObjectFromWindow 直接返回 E_FAIL，ROT 也取不到实例。
            var plugin = ReadPluginStates();
            var pluginPids = new HashSet<int>();
            foreach (var state in plugin.Values) pluginPids.Add(state.Pid);
            foreach (var x in list)
            {
                if (x.Kind == DocKind.Word) wordTargets++; else excelTargets++;
                uint ownerPid;
                Native.GetWindowThreadProcessId(x.Handle, out ownerPid);
                PluginState ps;
                if (plugin.TryGetValue(x.Handle, out ps) && ps.Pid == unchecked((int)ownerPid))
                {
                    result[x.Handle] = ps.Dirty;
                    if (x.Kind == DocKind.Word) wordHits++; else excelHits++;
                    continue;
                }
                if (pluginPids.Contains(unchecked((int)ownerPid)))
                {
                    // 插件正在上报这个进程，但新窗口还没到下一拍：沿用旧值，别退回会失败的 COM
                    if (x.Kind == DocKind.Word) wordHits++; else excelHits++;
                    bool oldPluginValue;
                    if (prev.TryGetValue(x.Handle, out oldPluginValue)) result[x.Handle] = oldPluginValue;
                    continue;
                }
                bool dirty;
                try
                {
                    if (TryWindowSaved(x.Handle, x.Kind, out dirty, out firstError))
                    {
                        result[x.Handle] = dirty;
                        if (x.Kind == DocKind.Word) wordHits++; else excelHits++;
                        continue;
                    }
                }
                catch (Exception ex) { if (firstError == null) firstError = ex.GetType().Name + ":" + ex.Message; }
                // 这一轮没读到（Office 忙、正在弹窗）：沿用上次的值，别让圆点闪掉
                bool oldValue;
                if (prev.TryGetValue(x.Handle, out oldValue)) result[x.Handle] = oldValue;
            }
            // 逐个窗口拿不到对象模型时，把 ROT 里所有 Word / Excel 实例都问一遍
            if (list.Length > 0)
            {
                if (wordHits == 0)
                {
                    try { QueryRot("Word.Application", DocKind.Word, result); }
                    catch (Exception ex) { if (firstError == null) firstError = "ROT:" + ex.Message; }
                }
                if (excelHits == 0)
                {
                    try { QueryRot("Excel.Application", DocKind.Excel, result); }
                    catch (Exception ex) { if (firstError == null) firstError = "ROT:" + ex.Message; }
                }
            }
            if (list.Length > 0 && result.Count == 0 && pluginPids.Count == 0 &&
                unchecked(Environment.TickCount - lastDiagnostic) > 10000)
            {
                lastDiagnostic = Environment.TickCount;
                Log.Write(string.Format("未保存读取失败：目标 {0} 个（Word {1} / Excel {2}）{3}",
                    list.Length, wordTargets, excelTargets, firstError == null ? "" : " " + firstError));
            }
            return result;
        }

        // 读 Office 插件写的状态：值名是窗口句柄（十六进制），值是 "进程ID|0/1|W/E"。
        static Dictionary<IntPtr, PluginState> ReadPluginStates()
        {
            var states = new Dictionary<IntPtr, PluginState>();
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(DirtyKey))
                {
                    if (key == null) return states;
                    foreach (string name in key.GetValueNames())
                    {
                        long raw;
                        if (!long.TryParse(name, System.Globalization.NumberStyles.HexNumber,
                            System.Globalization.CultureInfo.InvariantCulture, out raw)) continue;
                        string value = key.GetValue(name) as string;
                        if (value == null) continue;
                        int bar = value.IndexOf('|');
                        if (bar <= 0) continue;
                        int pid;
                        if (!int.TryParse(value.Substring(0, bar), out pid)) continue;
                        int bar2 = value.IndexOf('|', bar + 1);
                        string dirty = bar2 > bar ? value.Substring(bar + 1, bar2 - bar - 1) : value.Substring(bar + 1);
                        states[new IntPtr(raw)] = new PluginState { Pid = pid, Dirty = dirty == "1" };
                    }
                }
            }
            catch { }
            return states;
        }

        // 直接问“这个窗口所属进程”要它的文档对象，SDI 多进程下也能读到各自的已保存状态
        static bool TryWindowSaved(IntPtr hwnd, DocKind kind, out bool dirty, out string error)
        {
            dirty = false;
            error = null;
            object om = null;
            try
            {
                Guid iid = IID_Dispatch;
                int hr = Native.AccessibleObjectFromWindow(hwnd, Native.OBJID_NATIVEOM, ref iid, out om);
                if (hr != 0 || om == null) { error = "hr=0x" + hr.ToString("X8"); return false; }
                dynamic win = om;
                if (kind == DocKind.Word)
                {
                    // 取到的可能是 Window（有 Document），也可能是 Application（只有 ActiveDocument）；
                    // SDI 下一个进程就一份文档，两者指向同一份。
                    object doc = null;
                    try { doc = win.Document; } catch { }
                    if (doc == null) doc = win.ActiveDocument;
                    if (doc == null) { error = "no document"; return false; }
                    dirty = !(bool)((dynamic)doc).Saved;
                    Release(doc);
                    return true;
                }
                dynamic sheet = win.ActiveSheet;   // Excel：窗口里当前的工作表
                dynamic book = sheet.Parent;       // → 工作簿
                dirty = !(bool)book.Saved;
                Release(book);
                Release(sheet);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.GetType().Name + ":" + ex.Message;
                return false;
            }
            finally { if (om != null) Release(om); }
        }

        // 遍历运行对象表：每个 Word / Excel 实例都会挂在这里，SDI 下也不会漏
        static void QueryRot(string progId, DocKind kind, Dictionary<IntPtr, bool> result)
        {
            System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot;
            if (Native.GetRunningObjectTable(0, out rot) != 0 || rot == null) return;
            try
            {
                System.Runtime.InteropServices.ComTypes.IEnumMoniker en;
                rot.EnumRunning(out en);
                if (en == null) return;
                var one = new System.Runtime.InteropServices.ComTypes.IMoniker[1];
                IntPtr fetched = IntPtr.Zero;
                while (en.Next(1, one, fetched) == 0 && one[0] != null)
                {
                    var mon = one[0];
                    one[0] = null;
                    string name = null;
                    try { mon.GetDisplayName(null, null, out name); } catch { }
                    if (string.IsNullOrEmpty(name) || !name.StartsWith(progId, StringComparison.OrdinalIgnoreCase)) continue;
                    object appObj = null;
                    try
                    {
                        Guid iid = IID_Dispatch;
                        mon.BindToObject(null, null, ref iid, out appObj);
                        if (appObj == null) continue;
                        dynamic app = appObj;
                        if (kind == DocKind.Word)
                        {
                            dynamic wins = app.Windows;
                            int n = (int)wins.Count;
                            for (int i = 1; i <= n; i++)
                            {
                                try
                                {
                                    dynamic win = wins.Item(i);
                                    int hwnd = (int)win.Hwnd;
                                    bool saved = (bool)win.Document.Saved;
                                    result[new IntPtr(hwnd)] = !saved;
                                }
                                catch { }
                            }
                            Release(wins);
                        }
                        else
                        {
                            dynamic books = app.Workbooks;
                            int n = (int)books.Count;
                            for (int i = 1; i <= n; i++)
                            {
                                try
                                {
                                    dynamic wb = books.Item(i);
                                    bool saved = (bool)wb.Saved;
                                    dynamic ws = wb.Windows;
                                    int wn = (int)ws.Count;
                                    for (int j = 1; j <= wn; j++)
                                    {
                                        try
                                        {
                                            dynamic win = ws.Item(j);
                                            int hwnd = (int)win.Hwnd;
                                            result[new IntPtr(hwnd)] = !saved;
                                        }
                                        catch { }
                                    }
                                    Release(ws);
                                    Release(wb);
                                }
                                catch { }
                            }
                            Release(books);
                        }
                    }
                    catch { }
                    finally { Release(appObj); }
                }
            }
            finally { try { Marshal.ReleaseComObject(rot); } catch { } }
        }

        static string Describe(Dictionary<IntPtr, bool> before, Dictionary<IntPtr, bool> after)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var kv in after)
            {
                bool old;
                if (before.TryGetValue(kv.Key, out old) && old == kv.Value) continue;
                if (sb.Length > 0) sb.Append(' ');
                sb.AppendFormat("0x{0:X}={1}", kv.Key.ToInt64(), kv.Value ? "未保存" : "已保存");
            }
            foreach (var kv in before)
            {
                if (after.ContainsKey(kv.Key)) continue;
                if (sb.Length > 0) sb.Append(' ');
                sb.AppendFormat("0x{0:X}=关闭", kv.Key.ToInt64());
            }
            return sb.Length == 0 ? "(无)" : sb.ToString();
        }

        static void QueryWord(Dictionary<IntPtr, bool> result)
        {
            object appObj = null;
            try { appObj = Marshal.GetActiveObject("Word.Application"); }
            catch { return; }
            try
            {
                dynamic app = appObj;
                dynamic wins = app.Windows;
                int n = (int)wins.Count;
                for (int i = 1; i <= n; i++)
                {
                    try
                    {
                        dynamic win = wins.Item(i);
                        int hwnd = (int)win.Hwnd;
                        bool saved = (bool)win.Document.Saved;
                        result[new IntPtr(hwnd)] = !saved;
                    }
                    catch { }
                }
            }
            catch { }
            finally { Release(appObj); }
        }

        static void QueryExcel(Dictionary<IntPtr, bool> result)
        {
            object appObj = null;
            try { appObj = Marshal.GetActiveObject("Excel.Application"); }
            catch { return; }
            try
            {
                dynamic app = appObj;
                dynamic books = app.Workbooks;
                int n = (int)books.Count;
                for (int i = 1; i <= n; i++)
                {
                    try
                    {
                        dynamic wb = books.Item(i);
                        bool saved = (bool)wb.Saved;
                        dynamic wins = wb.Windows;
                        int wn = (int)wins.Count;
                        for (int j = 1; j <= wn; j++)
                        {
                            try
                            {
                                dynamic win = wins.Item(j);
                                int hwnd = (int)win.Hwnd;
                                result[new IntPtr(hwnd)] = !saved;
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
            }
            catch { }
            finally { Release(appObj); }
        }

        static void Release(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                try { Marshal.ReleaseComObject(obj); }
                catch { }
            }
        }

        static bool Same(Dictionary<IntPtr, bool> a, Dictionary<IntPtr, bool> b)
        {
            if (a.Count != b.Count) return false;
            foreach (var kv in a)
            {
                bool v;
                if (!b.TryGetValue(kv.Key, out v) || v != kv.Value) return false;
            }
            return true;
        }
    }
    static class Native
    {
        public const int WS_EX_TOOLWINDOW = 0x80, WS_EX_NOACTIVATE = 0x08000000;
        public const int WS_CAPTION = 0x00C00000, GWL_STYLE = -16, GWL_EXSTYLE = -20, WS_EX_TOPMOST = 0x8;
        public const int WM_MOUSEACTIVATE = 0x21, MA_NOACTIVATE = 3, WM_CLOSE = 0x10, WM_NCACTIVATE = 0x86;
        public const uint RDW_INVALIDATE = 1, RDW_NOCHILDREN = 0x40, RDW_UPDATENOW = 0x100, RDW_FRAME = 0x400;
        public const uint SMTO_ABORTIFHUNG = 2;
        public const int DWMWA_TRANSITIONS_FORCEDISABLED = 3;
        public const int WM_SYSCOMMAND = 0x112, SC_CLOSE = 0xF060;
        public const uint GW_OWNER = 4, GW_HWNDPREV = 3, GA_ROOT = 2, GW_HWNDNEXT = 2, GW_CHILD = 5;
        public const int SW_SHOWNORMAL = 1, SW_MAXIMIZE = 3, SW_SHOWNOACTIVATE = 4, SW_SHOWMINNOACTIVE = 7, SW_RESTORE = 9;
        public const uint SWP_NOSIZE = 1, SWP_NOMOVE = 2, SWP_NOZORDER = 4, SWP_NOACTIVATE = 0x10, SWP_NOOWNERZORDER = 0x200, SWP_ASYNCWINDOWPOS = 0x4000;
        public static readonly IntPtr HWND_TOP = IntPtr.Zero, HWND_TOPMOST = new IntPtr(-1), HWND_NOTOPMOST = new IntPtr(-2);
        public const uint EVENT_SYSTEM_FOREGROUND = 3, EVENT_SYSTEM_MOVESIZESTART = 0x0A, EVENT_SYSTEM_MOVESIZEEND = 0x0B;
        public const uint EVENT_SYSTEM_MINIMIZESTART = 0x16, EVENT_SYSTEM_MINIMIZEEND = 0x17;
        public const int SW_HIDE = 0;
        public const uint EVENT_OBJECT_DESTROY = 0x8001, EVENT_OBJECT_HIDE = 0x8003;
        public const uint EVENT_OBJECT_LOCATIONCHANGE = 0x800B, EVENT_OBJECT_NAMECHANGE = 0x800C;
        public const uint EVENT_OBJECT_CLOAKED = 0x8017, EVENT_OBJECT_UNCLOAKED = 0x8018;
        public const uint WINEVENT_OUTOFCONTEXT = 0, WINEVENT_SKIPOWNPROCESS = 2;
        public const int OBJID_WINDOW = 0;
        public const uint INPUT_MOUSE = 0;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left, Top, Right, Bottom;
            public bool Equals(RECT o) { return Left == o.Left && Top == o.Top && Right == o.Right && Bottom == o.Bottom; }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT { public int X, Y; }


        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public POINT ptMinPosition;
            public POINT ptMaxPosition;
            public RECT rcNormalPosition;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT
        {
            public uint type;
            public MOUSEINPUT mi;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEINPUT
        {
            public int dx, dy;
            public uint mouseData, dwFlags, time;
            public IntPtr dwExtraInfo;
        }

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        public delegate void WinEventProc(IntPtr hook, uint ev, IntPtr hwnd, int idObject, int idChild, uint thread, uint time);
        public delegate IntPtr LowLevelProc(int code, IntPtr wParam, IntPtr lParam);

        public const int WH_KEYBOARD_LL = 13, WH_MOUSE_LL = 14, WM_LBUTTONUP = 0x202, WM_SYSKEYDOWN = 0x104;
        public const int VK_LBUTTON = 1, VK_RBUTTON = 2;
        [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int vk);

        [StructLayout(LayoutKind.Sequential)]
        public struct MSLLHOOKSTRUCT { public POINT pt; public uint mouseData, flags, time; public IntPtr extra; }

        [StructLayout(LayoutKind.Sequential)]
        public struct KBDLLHOOKSTRUCT { public uint vkCode, scanCode, flags, time; public IntPtr extra; }

        [DllImport("user32.dll")] public static extern IntPtr SetWindowsHookEx(int id, LowLevelProc proc, IntPtr mod, uint thread);
        [DllImport("user32.dll")] public static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")] public static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr GetModuleHandle(string name);

        [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
        [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr ctx);
        [DllImport("user32.dll")] static extern IntPtr MonitorFromWindow(IntPtr h, uint flags);
        [DllImport("user32.dll")] static extern IntPtr MonitorFromPoint(POINT pt, uint flags);
        [DllImport("shcore.dll")] static extern int GetDpiForMonitor(IntPtr mon, int type, out uint dpiX, out uint dpiY);
        public const int WM_DPICHANGED = 0x02E0;

        // 窗口所在显示器的缩放比例（96 DPI = 1.0）；h 为空时取主显示器
        public static float MonitorScale(IntPtr h)
        {
            try
            {
                IntPtr mon = h != IntPtr.Zero ? MonitorFromWindow(h, 2 /* NEAREST */) : MonitorFromPoint(new POINT(), 1 /* PRIMARY */);
                uint x, y;
                if (mon != IntPtr.Zero && GetDpiForMonitor(mon, 0 /* EFFECTIVE */, out x, out y) == 0 && x > 0) return x / 96f;
            }
            catch (DllNotFoundException) { }
            catch (EntryPointNotFoundException) { }
            using (var g = Graphics.FromHwnd(IntPtr.Zero)) return g.DpiX / 96f;
        }
        [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
        [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc cb, IntPtr lParam);
        [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
        [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
        [DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr h);
        [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr h, uint cmd);
        [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint flags);
        [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT pt);
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
        [DllImport("user32.dll")] public static extern void SwitchToThisWindow(IntPtr h, bool altTab);
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
        [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr h, int cmd);
        [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);

        [DllImport("user32.dll")] static extern bool GetWindowPlacement(IntPtr h, ref WINDOWPLACEMENT wp);
        [DllImport("user32.dll")] public static extern bool SetWindowPlacement(IntPtr h, ref WINDOWPLACEMENT wp);
        [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr h, int index);
        [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] public static extern bool ReleaseCapture();
        [DllImport("ole32.dll")] public static extern int CoInitializeEx(IntPtr reserved, uint coInit);
        [DllImport("ole32.dll")] public static extern void CoUninitialize();
        [DllImport("ole32.dll")]
        public static extern int GetRunningObjectTable(int reserved, out System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot);
        // 取窗口所属进程内的 Office 对象模型（OBJID_NATIVEOM）。Word/Excel 的 SDI 模式下一份文档
        // 一个进程，只认一个 Application 实例会漏掉其它进程里的文档。
        public const uint OBJID_NATIVEOM = 0xFFFFFFF0u;
        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId, ref Guid iid,
            [MarshalAs(UnmanagedType.Interface)] out object obj);
        [DllImport("winmm.dll")] public static extern uint timeBeginPeriod(uint ms);
        [DllImport("winmm.dll")] public static extern uint timeEndPeriod(uint ms);
        [DllImport("user32.dll")] public static extern bool RedrawWindow(IntPtr h, IntPtr rc, IntPtr rgn, uint flags);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr SendMessageTimeout(IntPtr h, int msg, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr result);
        [DllImport("user32.dll")] static extern IntPtr GetDC(IntPtr h);
        [DllImport("user32.dll")] static extern int ReleaseDC(IntPtr h, IntPtr dc);
        [DllImport("gdi32.dll")] static extern uint GetPixel(IntPtr dc, int x, int y);

        // 文档正文视图（Word 的 _WwG / Excel 的 EXCEL7）。开着导航窗格、编辑器等任务窗格时，
        // 标签栏只能铺在正文区域上，不能盖住任务窗格的标题和关闭按钮。
        public static bool GetDocViewRect(IntPtr frame, out RECT view)
        {
            view = new RECT();
            if (frame == IntPtr.Zero) return false;
            long best = 0;
            RECT bestRect = new RECT();
            EnumWindowsProc cb = delegate(IntPtr h, IntPtr l)
            {
                if (!IsWindowVisible(h)) return true;
                string cls = GetClassNameOf(h);
                if (cls != "_WwG" && cls != "EXCEL7") return true;
                RECT r;
                if (!GetWindowRect(h, out r)) return true;
                long area = (long)(r.Right - r.Left) * (r.Bottom - r.Top);
                if (area > best) { best = area; bestRect = r; }
                return true;
            };
            EnumChildWindows(frame, cb, IntPtr.Zero);
            if (best <= 0) return false;
            view = bestRect;
            return true;
        }

        // 是否打开了左侧任务窗格（Word 的导航、编辑器，Excel 的任务窗格等）。
        // Word 把它放在 MsoCommandBarDock 里，窗口文本为 MsoDockLeft。
        public static bool HasLeftDockPane(IntPtr frame)
        {
            if (frame == IntPtr.Zero) return false;
            bool found = false;
            EnumWindowsProc cb = delegate(IntPtr h, IntPtr l)
            {
                if (!IsWindowVisible(h)) return true;
                if (GetClassNameOf(h) != "MsoCommandBarDock") return true;
                if (GetWindowTextOf(h) != "MsoDockLeft") return true;
                RECT r;
                if (!GetWindowRect(h, out r)) return true;
                if (r.Right - r.Left < 80) return true;
                found = true;
                return false;
            };
            EnumChildWindows(frame, cb, IntPtr.Zero);
            return found;
        }

        // Word/Excel 的“文件/导出”全屏界面（Backstage）是主窗口里的 FullpageUIHost，
        // 它会盖住整块客户区；这时标签栏必须让开，不能压在导出界面上。
        public static bool HasFullpageOverlay(IntPtr h)
        {
            if (h == IntPtr.Zero) return false;
            RECT wr;
            if (!GetWindowRect(h, out wr)) return false;
            long frameArea = (long)(wr.Right - wr.Left) * (wr.Bottom - wr.Top);
            if (frameArea <= 0) return false;
            bool found = false;
            EnumWindowsProc cb = delegate(IntPtr c, IntPtr l)
            {
                if (!IsWindowVisible(c)) return true;
                if (GetClassNameOf(c) != "FullpageUIHost") return true;
                RECT r;
                if (!GetWindowRect(c, out r)) return true;
                long area = (long)(r.Right - r.Left) * (r.Bottom - r.Top);
                if (area * 2 >= frameArea) { found = true; return false; }
                return true;
            };
            EnumChildWindows(h, cb, IntPtr.Zero);
            return found;
        }
        // 取屏幕上某一点的像素颜色（屏幕坐标）
        public static Color SampleScreenColor(int x, int y)
        {
            IntPtr dc = GetDC(IntPtr.Zero);
            if (dc == IntPtr.Zero) return Color.Empty;
            try
            {
                uint v = GetPixel(dc, x, y);
                if (v == 0xFFFFFFFF) return Color.Empty;
                return Color.FromArgb((int)(v & 0xFF), (int)((v >> 8) & 0xFF), (int)((v >> 16) & 0xFF));
            }
            catch { return Color.Empty; }
            finally { ReleaseDC(IntPtr.Zero, dc); }
        }
        [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
        [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool attach);
        [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder sb, int max);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder sb, int max);
        [DllImport("user32.dll")] static extern int GetWindowTextLength(IntPtr h);
        [DllImport("user32.dll")] public static extern IntPtr SetWinEventHook(uint min, uint max, IntPtr mod, WinEventProc proc, uint pid, uint tid, uint flags);
        [DllImport("user32.dll")] public static extern bool UnhookWinEvent(IntPtr hook);
        [DllImport("user32.dll")] public static extern uint SendInput(uint count, INPUT[] inputs, int size);
        [DllImport("user32.dll")] public static extern bool DestroyIcon(IntPtr h);
        [DllImport("dwmapi.dll")] static extern int DwmGetWindowAttribute(IntPtr h, int attr, out int value, int size);
        [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr h, int attr, ref int value, int size);

        // 让普通窗口呈现“真正最大化”的外观：无圆角（DWMWA_WINDOW_CORNER_PREFERENCE）、无边框（DWMWA_BORDER_COLOR）
        // 关闭 DWM 的窗口过渡动画，切标签时不做淡入/放大
        public static void DisableTransitions(IntPtr h)
        {
            try { int on = 1; DwmSetWindowAttribute(h, DWMWA_TRANSITIONS_FORCEDISABLED, ref on, 4); } catch { }
        }

        // 在窗口还没有真正到前台时，先把它按“已激活”的样子画好、提到顶层但保持不激活。
        // 随后 SetForegroundWindow 只是改激活状态，Word 的标题栏/文件名不再整条重绘。
        public static void PrepareActivation(IntPtr h)
        {
            DisableTransitions(h);
            IntPtr result;
            SendMessageTimeout(h, WM_NCACTIVATE, new IntPtr(1), IntPtr.Zero, SMTO_ABORTIFHUNG, 40, out result);
            RedrawWindow(h, IntPtr.Zero, IntPtr.Zero, RDW_FRAME | RDW_UPDATENOW | RDW_NOCHILDREN);
        }

        public static void SetMaximizedLook(IntPtr h, bool on)
        {
            if (h == IntPtr.Zero || !IsWindow(h)) return;
            try
            {
                int round = on ? 1 /* DONOTROUND */ : 0 /* DEFAULT */;
                DwmSetWindowAttribute(h, 33, ref round, 4);
                int border = on ? unchecked((int)0xFFFFFFFE) /* NONE */ : unchecked((int)0xFFFFFFFF) /* DEFAULT */;
                DwmSetWindowAttribute(h, 34, ref border, 4);
            }
            catch { }
        }
        [DllImport("dwmapi.dll")] static extern int DwmGetWindowAttribute(IntPtr h, int attr, out RECT value, int size);
        [DllImport("kernel32.dll")] static extern IntPtr OpenProcess(uint access, bool inherit, uint pid);
        [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern bool QueryFullProcessImageName(IntPtr h, int flags, StringBuilder sb, ref int size);

        public static string GetClassNameOf(IntPtr h)
        {
            var sb = new StringBuilder(64);
            GetClassName(h, sb, sb.Capacity);
            return sb.ToString();
        }

        public static string GetWindowTextOf(IntPtr h)
        {
            int len = GetWindowTextLength(h);
            if (len <= 0) return "";
            var sb = new StringBuilder(len + 1);
            GetWindowText(h, sb, sb.Capacity);
            return sb.ToString();
        }

        public static bool IsCloaked(IntPtr h)
        {
            int v;
            return DwmGetWindowAttribute(h, 14 /* DWMWA_CLOAKED */, out v, 4) == 0 && v != 0;
        }

        // 窗口可见边框（不含 Win10/11 的隐形缩放边框）
        public static RECT FrameBounds(IntPtr h, RECT fallback)
        {
            RECT r;
            if (DwmGetWindowAttribute(h, 9 /* DWMWA_EXTENDED_FRAME_BOUNDS */, out r, Marshal.SizeOf(typeof(RECT))) == 0 && r.Right > r.Left)
                return r;
            return fallback;
        }


        // 最小化/最大化/关闭按钮占用的区域；标签栏要在这里之前收边
        public static bool GetCaptionButtons(IntPtr h, out RECT r)
        {
            r = new RECT();
            return h != IntPtr.Zero &&
                   DwmGetWindowAttribute(h, 5 /* DWMWA_CAPTION_BUTTON_BOUNDS */, out r, Marshal.SizeOf(typeof(RECT))) == 0 &&
                   r.Right > r.Left;
        }

        // 取屏幕上单个像素的颜色：用来跟随 Word 标题栏的浅色/深色主题
        public static Color ScreenPixel(int x, int y)
        {
            IntPtr dc = GetDC(IntPtr.Zero);
            try
            {
                uint v = GetPixel(dc, x, y);
                if (v == 0xFFFFFFFF) return Color.Empty;
                return Color.FromArgb((int)(v & 0xFF), (int)((v >> 8) & 0xFF), (int)((v >> 16) & 0xFF));
            }
            finally { ReleaseDC(IntPtr.Zero, dc); }
        }

        // 找到 Word / Excel 顶部功能区的底边；失败时退回正文视图（Word 的 _WwG / Excel 的 EXCEL7）的顶边
        public static int BelowRibbonTop(IntPtr h, int fallback)
        {
            int bottom = 0, docTop = 0;
            long docArea = 0;
            FindRibbon(h, ref bottom, ref docTop, ref docArea);
            // Excel 的快速访问工具栏和功能区是两层不同的窗口；只认 MsoDockTop 会停在标题栏下方。
            // Word 维持原来的 MsoDockTop 优先逻辑，避免改变现有定位。
            if (GetClassNameOf(h) == "XLMAIN")
            {
                int top = Math.Max(bottom, docTop);
                if (top > 0) return top;
            }
            else
            {
                if (bottom > 0) return bottom;
                if (docTop > 0) return docTop;
            }
            return fallback;
        }

        static void FindRibbon(IntPtr parent, ref int bottom, ref int docTop, ref long docArea)
        {
            IntPtr c = GetWindow(parent, GW_CHILD);
            while (c != IntPtr.Zero)
            {
                string cls = GetClassNameOf(c);
                if (IsWindowVisible(c))
                {
                    if (cls == "MsoCommandBarDock" && GetWindowTextOf(c) == "MsoDockTop")
                    {
                        RECT r;
                        if (GetWindowRect(c, out r)) bottom = Math.Max(bottom, r.Bottom);
                    }
                    else if (cls == "_WwG" || cls == "EXCEL7")
                    {
                        RECT r;
                        if (GetWindowRect(c, out r))
                        {
                            long area = (long)(r.Right - r.Left) * (r.Bottom - r.Top);
                            if (area > docArea)
                            {
                                docArea = area;
                                docTop = r.Top;
                            }
                        }
                    }
                }
                FindRibbon(c, ref bottom, ref docTop, ref docArea);
                c = GetWindow(c, GW_HWNDNEXT);
            }
        }
        public static WINDOWPLACEMENT GetPlacement(IntPtr h)
        {
            var wp = new WINDOWPLACEMENT();
            wp.length = Marshal.SizeOf(typeof(WINDOWPLACEMENT));
            GetWindowPlacement(h, ref wp);
            return wp;
        }

        public static string GetProcessPath(uint pid)
        {
            IntPtr hp = OpenProcess(0x1000 /* PROCESS_QUERY_LIMITED_INFORMATION */, false, pid);
            if (hp == IntPtr.Zero) return null;
            try
            {
                var sb = new StringBuilder(1024);
                int size = sb.Capacity;
                return QueryFullProcessImageName(hp, 0, sb, ref size) ? sb.ToString() : null;
            }
            finally { CloseHandle(hp); }
        }
    }
}
