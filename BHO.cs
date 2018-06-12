using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

using SHDocVw;
using mshtml;
using System.IO;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices.Expando;
using System.Reflection;

namespace IEMonitor
{
    /* define the IObjectWithSite interface which the BHO class will implement.
     * The IObjectWithSite interface provides simple objects with a lightweight siting mechanism (lighter than IOleObject).
     * Often, an object must communicate directly with a container site that is managing the object. 
     * Outside of IOleObject::SetClientSite, there is no generic means through which an object becomes aware of its site. 
     * The IObjectWithSite interface provides a siting mechanism. This interface should only be used when IOleObject is not already in use.
     * By using IObjectWithSite, a container can pass the IUnknown pointer of its site to the object through SetSite. 
     * Callers can also get the latest site passed to SetSite by using GetSite.
     */
    [
        ComVisible(true),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
        Guid("FC4801A3-2BA9-11CF-A229-00AA003D7352")
        // Never EVER change this UUID!! It allows this BHO to find IE and attach to it
    ]
    public interface IObjectWithSite
    {
        [PreserveSig]
        int SetSite([MarshalAs(UnmanagedType.IUnknown)]object site);
        [PreserveSig]
        int GetSite(ref Guid guid, out IntPtr ppvSite);
    }

    [
        ComVisible(true),
        Guid("4C1D2E51-018B-4A7C-8A07-618452573E42"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IExtension
    {
        [DispId(1)]
        void sendMsgToQt(string s);
    }

    /* The BHO site is the COM interface used to establish a communication.
     * Define a GUID attribute for your BHO as it will be used later on during registration / installation
     */
    [
            ComVisible(true),
            Guid("2788CB25-EF9A-54C1-B43C-E30D1A4A1992"),
            ClassInterface(ClassInterfaceType.None), ProgId("toQt"),
            ComDefaultInterface(typeof(IExtension))
    ]
    public class BHO : IObjectWithSite, IExtension
    {
        private WebBrowser webBrowser;

        private string cUrl = null;

        // function list id
        private ArrayList funcList = new ArrayList();

        public const string BHO_REGISTRY_KEY_NAME =
               "Software\\Microsoft\\Windows\\" +
               "CurrentVersion\\Explorer\\Browser Helper Objects";

        //消息标识

        private const int WM_COPYDATA = 0x004A;

        //消息数据类型(typeFlag以上二进制，typeFlag以下字符)

        private const uint typeFlag = 0x8000;


        /* The SetSite() method is where the BHO is initialized and where you would perform all the tasks that happen only once.
         * When you navigate to a URL with Internet Explorer, you should wait for a couple of events to make sure the required document
         * has been completely downloaded and then initialized. Only at this point can you safely access its content through the exposed
         * object model, if any. This means you need to acquire a couple of pointers. The first one is the pointer to IWebBrowser2, 
         * the interface that renders the WebBrowser object. The second pointer relates to events.
         * This module must register as an event listener with the browser in order to receive the notification of downloads
         * and document-specific events.
         */
        public int SetSite(object site)
        {
            if (site != null)
            {
                webBrowser = (WebBrowser)site;
                webBrowser.DocumentComplete +=
                  new DWebBrowserEvents2_DocumentCompleteEventHandler(
                  this.OnDocumentComplete);
            }
            else
            {
                webBrowser.DocumentComplete -=
                  new DWebBrowserEvents2_DocumentCompleteEventHandler(
                  this.OnDocumentComplete);
                webBrowser = null;
            }

            return 0;
        }

        public int GetSite(ref Guid guid, out IntPtr ppvSite)
        {
            IntPtr punk = Marshal.GetIUnknownForObject(webBrowser);
            int hr = Marshal.QueryInterface(punk, ref guid, out ppvSite);
            Marshal.Release(punk);
            return hr;
        }

        public void WriteLog(string documentName, string msg)
        {
            //string errorLogFilePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log");
            // TODO - path to app dir
            //errorLogFilePath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "/bholog/";


            string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string errorLogFilePath = Path.GetDirectoryName(dllPath);
            errorLogFilePath = Path.Combine(errorLogFilePath, "bholog");
            if (!System.IO.Directory.Exists(errorLogFilePath))
            {
                System.IO.Directory.CreateDirectory(errorLogFilePath);
            }
            string logFile = System.IO.Path.Combine(errorLogFilePath, documentName + "@" + DateTime.Today.ToString("yyyy-MM-dd") + ".txt");
            bool writeBaseInfo = System.IO.File.Exists(logFile);
            StreamWriter swLogFile = new StreamWriter(logFile, true, Encoding.Unicode);
            swLogFile.WriteLine(DateTime.Now.ToString("HH:mm:ss") + "\t" + msg);
            swLogFile.Close();
            swLogFile.Dispose();
        }

        private string GetLocalIp()
        {
            IPHostEntry ipe = Dns.GetHostEntry(Dns.GetHostName());

            IPAddress[] ipas = ipe.AddressList;
            for (int n = 0; n < ipas.Length; n++)
            {
                string ipS = ipas[n].ToString();
                if (ipS.Contains("."))
                {
                    return ipS;
                }
            }

            return "127.0.0.1";
        }

        public void OnDocumentComplete(object pDisp, ref object URL)
        {

            HTMLDocument document = (HTMLDocument)webBrowser.Document;

            dynamic window = document.parentWindow;
            IExpando ScriptObject = (IExpando)window;
            PropertyInfo btnEvent = ScriptObject.GetProperty("toQt", BindingFlags.Default);
            if (btnEvent == null) btnEvent = ScriptObject.AddProperty("toQt");
            btnEvent.SetValue(ScriptObject, this, null);

            IHTMLElement head = (IHTMLElement)((IHTMLElementCollection)
                                    document.all.tags("head")).item(null, 0);

            IHTMLScriptElement scriptObject =
                (IHTMLScriptElement)document.createElement("script");
            scriptObject.type = @"text/javascript";
            scriptObject.text = "var myEles = document.getElementById('su'); if(myEles != undefined && myEles != null) {myEles.onmouseup=aaa;function aaa(){window.toQt.sendMsgToQt('hahaha')}}";
            ((HTMLHeadElement)head).appendChild((IHTMLDOMNode)scriptObject);

        }

        public void sendMsgToQt(string msg)
        {
            WriteLog("toQt", msg);
            string strDlgTitle = "个人工作集成平台助手";

            //did work
            SendString(strDlgTitle,0,"this is my msg");

            //todo-test
            IntPtr hwndRecvWindow = ImportFromDLL.FindWindow(null, strDlgTitle);
            if (hwndRecvWindow == IntPtr.Zero)
            {
                WriteLog("error", "请先启动接收消息程序");
                return;
            }
            else
            {
                WriteLog("log", "found recv message window...");
            }


            IntPtr hwndSendWindow = ImportFromDLL.GetConsoleWindow();
            if (hwndSendWindow == IntPtr.Zero)
            {
                WriteLog("error", "获取自己的窗口句柄失败，请重试");
                return;
            }

            for (int i = 0; i < 10; i++)
            {
                string strText = DateTime.Now.ToString();
                ImportFromDLL.COPYDATASTRUCT copydata = new ImportFromDLL.COPYDATASTRUCT();
                copydata.cbData = Encoding.Default.GetBytes(strText).Length;
                copydata.lpData = strText;

                ImportFromDLL.SendMessage(hwndRecvWindow, ImportFromDLL.WM_COPYDATA, hwndSendWindow, ref copydata);

                WriteLog("success", strText);
            }

            
        }


        public struct COPYDATASTRUCT
        {

            public IntPtr dwData;

            public int cbData;

            public IntPtr lpData;

        }

        

        [DllImport("User32.dll", EntryPoint="SendMessage")]
        
        private static extern int SendMessage(

            int hWnd,                                 // handle to destination window

            int Msg,                              // message

            int wParam,                               // first message parameter

            ref COPYDATASTRUCT lParam    // second message parameter

            );


        [DllImport("User32.dll", EntryPoint = "FindWindow")]

        private static extern int FindWindow(string lpClassName, string lpWindowName);



        //发送数据委托与事件定义

        public delegate void SendStringEvent(object sender, uint flag, string str);

        public delegate void SendBytesEvent(object sender, uint flag, byte[] bt);

        public event SendStringEvent OnSendString;

        public event SendBytesEvent OnSendBytes;


        /// 发送字符串格式数据

        /// </summary>

        /// <param name="destWindow">目标窗口标题</param>

        /// <param name="flag">数据标志</param>

        /// <param name="str">数据</param>

        /// <returns></returns>

        public bool SendString(string destWindow, uint flag, string str)
        {

            if (flag > typeFlag)
            {

                WriteLog("error","要发送的数据不是字符格式");

                return false;

            }

            int WINDOW_HANDLER = FindWindow(null, @destWindow);

            if (WINDOW_HANDLER == 0)
            {
                WriteLog("error", "not found window");
                return false;
            }
            try
            {

                byte[] sarr = System.Text.Encoding.Default.GetBytes(str);

                COPYDATASTRUCT cds;

                cds.dwData = (IntPtr)flag;

                cds.cbData = sarr.Length;

                cds.lpData = Marshal.AllocHGlobal(sarr.Length);

                Marshal.Copy(sarr, 0, cds.lpData, sarr.Length);

                SendMessage(WINDOW_HANDLER, WM_COPYDATA, 0, ref cds);

                if (OnSendString != null)
                {

                    OnSendString(this, flag, str);

                }

                return true;

            }

            catch (Exception e)
            {

                WriteLog("error", e.Message);

                return false;

            }

        }

        /// <summary>

        /// 发送二进制格式数据

        /// </summary>

        /// <param name="destWindow">目标窗口</param>

        /// <param name="flag">数据标志</param>

        /// <param name="data">数据</param>

        /// <returns></returns>

        public bool SendBytes(string destWindow, uint flag, byte[] data)
        {

            if (flag <= typeFlag)
            {

                WriteLog("error", "要发送的数据不是二进制格式");

                return false;

            }

            int WINDOW_HANDLER = FindWindow(null, @destWindow);

            if (WINDOW_HANDLER == 0) return false;

            try
            {

                COPYDATASTRUCT cds;

                cds.dwData = (IntPtr)flag;

                cds.cbData = data.Length;

                cds.lpData = Marshal.AllocHGlobal(data.Length);

                Marshal.Copy(data, 0, cds.lpData, data.Length);

                SendMessage(WINDOW_HANDLER, WM_COPYDATA, 0, ref cds);

                if (OnSendBytes != null)
                {

                    OnSendBytes(this, flag, data);

                }

                return true;

            }

            catch (Exception e)
            {

                WriteLog("error", e.Message);

                return false;

            }

        }



        





        /* The Register method simply tells IE which is the GUID of your extension so that it could be loaded.
         * The "No Explorer" value simply says that we don't want to be loaded by Windows Explorer.
         */
        [ComRegisterFunction]
        public static void RegisterBHO(Type type)
        {
            RegistryKey registryKey =
              Registry.LocalMachine.OpenSubKey(BHO_REGISTRY_KEY_NAME, true);

            if (registryKey == null)
                registryKey = Registry.LocalMachine.CreateSubKey(
                                        BHO_REGISTRY_KEY_NAME);

            string guid = type.GUID.ToString("B");
            RegistryKey ourKey = registryKey.OpenSubKey(guid);

            if (ourKey == null)
            {
                ourKey = registryKey.CreateSubKey(guid);
            }

            ourKey.SetValue("NoExplorer", 1, RegistryValueKind.DWord);

            registryKey.Close();
            ourKey.Close();
        }

        [ComUnregisterFunction]
        public static void UnregisterBHO(Type type)
        {
            RegistryKey registryKey =
              Registry.LocalMachine.OpenSubKey(BHO_REGISTRY_KEY_NAME, true);
            string guid = type.GUID.ToString("B");

            if (registryKey != null)
                registryKey.DeleteSubKey(guid, false);
        }

    }

    public class ImportFromDLL
    {
        public const int WM_COPYDATA = 0x004A;

        //启用非托管代码  
        [StructLayout(LayoutKind.Sequential)]
        public struct COPYDATASTRUCT
        {
            public int dwData;    //not used  
            public int cbData;    //长度  
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpData;
        }

        [DllImport("User32.dll")]
        public static extern int SendMessage(
            IntPtr hWnd,　　　  // handle to destination window   
            int Msg,　　　      // message  
            IntPtr wParam,　   // first message parameter   
            ref COPYDATASTRUCT pcd // second message parameter   
        );

        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("Kernel32.dll", EntryPoint = "GetConsoleWindow")]
        public static extern IntPtr GetConsoleWindow();

    }
}
