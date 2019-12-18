using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Threading;

namespace BethVideo
{
    //对外消息委托代理
    public delegate void RecvTempMsgHandler(string Msg);

    [GuidAttribute("1A585C4D-3371-48dc-AF8A-AFFECC1B0968")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ControlEvents
    {
        [DispIdAttribute(0x001)]
        void OnRecvTempMsg(string Msg);
    }

    [Guid("984770F0-A5B6-441f-A01A-58B6F1E31E3F")]
    [ComSourceInterfacesAttribute(typeof(ControlEvents))]
    public partial class ucVideo : UserControl, IObjectSafety
    {
        /*******************************************************/
        #region IObjectSafety 接口成员实现（直接拷贝即可）

        private const string _IID_IDispatch = "{00020400-0000-0000-C000-000000000046}";
        private const string _IID_IDispatchEx = "{a6ef9860-c720-11d0-9337-00a0c90dcaa9}";
        private const string _IID_IPersistStorage = "{0000010A-0000-0000-C000-000000000046}";
        private const string _IID_IPersistStream = "{00000109-0000-0000-C000-000000000046}";
        private const string _IID_IPersistPropertyBag = "{37D84F60-42CB-11CE-8135-00AA004BB851}";

        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x00000001;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x00000002;
        private const int S_OK = 0;
        private const int E_FAIL = unchecked((int)0x80004005);
        private const int E_NOINTERFACE = unchecked((int)0x80004002);

        private bool _fSafeForScripting = true;
        private bool _fSafeForInitializing = true;

        public int GetInterfaceSafetyOptions(ref Guid riid, ref int pdwSupportedOptions, ref int pdwEnabledOptions)
        {
            int Rslt = E_FAIL;

            string strGUID = riid.ToString("B");
            pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            switch (strGUID)
            {
                case _IID_IDispatch:
                case _IID_IDispatchEx:
                    Rslt = S_OK;
                    pdwEnabledOptions = 0;
                    if (_fSafeForScripting == true)
                        pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER;
                    break;
                case _IID_IPersistStorage:
                case _IID_IPersistStream:
                case _IID_IPersistPropertyBag:
                    Rslt = S_OK;
                    pdwEnabledOptions = 0;
                    if (_fSafeForInitializing == true)
                        pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_DATA;
                    break;
                default:
                    Rslt = E_NOINTERFACE;
                    break;
            }

            return Rslt;
        }

        public int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions)
        {
            int Rslt = E_FAIL;
            string strGUID = riid.ToString("B");
            switch (strGUID)
            {
                case _IID_IDispatch:
                case _IID_IDispatchEx:
                    if (((dwEnabledOptions & dwOptionSetMask) == INTERFACESAFE_FOR_UNTRUSTED_CALLER) && (_fSafeForScripting == true))
                        Rslt = S_OK;
                    break;
                case _IID_IPersistStorage:
                case _IID_IPersistStream:
                case _IID_IPersistPropertyBag:
                    if (((dwEnabledOptions & dwOptionSetMask) == INTERFACESAFE_FOR_UNTRUSTED_DATA) && (_fSafeForInitializing == true))
                        Rslt = S_OK;
                    break;
                default:
                    Rslt = E_NOINTERFACE;
                    break;
            }

            return Rslt;
        }
        #endregion
        /*******************************************************/
        //定义事件
        public event RecvTempMsgHandler OnRecvTempMsg;

        //定义url类型 1 为单个对象 2为数组
        public const int UrlType = 1;
        public const int ArrType = 2;

        //用来处理数组信息，防止程序卡死
        public Thread m_DealArrData;

        public string m_sJsonData; //存储h5请求数据
        private bool m_bIsStartThread;//是否进行处理

        public ucVideo()
        {
            InitializeComponent();
            m_DealArrData = new Thread(DealArrInfo);
            m_DealArrData.Start();
            m_bIsStartThread = false;
        }
        #region [使用的私有函数]
        /// <summary>
        /// 下载图片
        /// </summary>
        /// <param name="picUrl">图片Http地址</param>
        /// <param name="savePath">保存路径</param>
        /// <param name="timeOut">Request最大请求时间，如果为-1则无限制</param>
        /// <returns></returns>
        private bool DownloadPicture(string picUrl, string savePath, int timeOut)
        {
            bool value = false;
            WebResponse response = null;
            Stream stream = null;
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(picUrl);
                if (timeOut != -1) request.Timeout = timeOut;
                response = request.GetResponse();
                stream = response.GetResponseStream();
                if (!response.ContentType.ToLower().StartsWith("text/"))
                    value = SaveBinaryFile(response, savePath);
            }
            finally
            {
                if (stream != null) stream.Close();
                if (response != null) response.Close();
            }
            return value;
        }
        /// <summary>
        /// 保存二进制文件
        /// </summary>
        /// <param name="response"></param>
        /// <param name="savePath"></param>
        /// <returns></returns>
        private bool SaveBinaryFile(WebResponse response, string savePath)
        {
            bool value = false;
            byte[] buffer = new byte[1024];
            Stream outStream = null;
            Stream inStream = null;
            try
            {
                if (File.Exists(savePath)) File.Delete(savePath);
                outStream = System.IO.File.Create(savePath);
                inStream = response.GetResponseStream();
                int l;
                do
                {
                    l = inStream.Read(buffer, 0, buffer.Length);
                    if (l > 0) outStream.Write(buffer, 0, l);
                } while (l > 0);
                value = true;
            }
            finally
            {
                if (outStream != null) outStream.Close();
                if (inStream != null) inStream.Close();
            }
            return value;
        }
        /// <summary>
        /// 开启子线程进行图片保存
        /// </summary>
        /// <param name="_sJson">数组</param>
        public void DealArrInfo()
        {
            while (true)
            {
                Thread.Sleep(2000);
                if (m_bIsStartThread)
                {
                    string sUrl = "";
                    string sPath = "";
                    string sName = "";
                    var varJsonInfo = JArray.Parse(m_sJsonData);
                    for (int i = 0; i < varJsonInfo.Count; i++)
                    {
                        sUrl = varJsonInfo[i]["Url"].ToString();
                        sPath = varJsonInfo[i]["Path"].ToString();
                        sName = varJsonInfo[i]["Name"].ToString();
                        sPath = sPath + "//" + sName;
                        this.DownloadPicture(sUrl, sPath, -1);
                    }
                    m_bIsStartThread = false;
                }
            }
            
        }
        /// <summary>
        /// 将得到的数据转换为base64
        /// </summary>
        /// <param name="response"></param>
        private string DealPicToBase64(WebResponse response)
        {
            byte[] buffer = new byte[1024];
            Stream inStream = null;
            inStream = response.GetResponseStream();
            inStream.Read(buffer, 0, buffer.Length);
            string sBase64 = Convert.ToBase64String(buffer);
            return sBase64;
        }
        #endregion

        #region [对外接口]
        /// <summary>
        /// 根据rul下载图片
        /// </summary>
        /// <param name="_sJson"></param>
        /// <param name="_sType"></param>
        public void DownPicWithUrl(string _sJson, string _sType)
        {
            m_sJsonData = _sJson;
            string sUrl = "";
            string sPath = "";
            string sName = "";
            int iType = Convert.ToInt32(_sType);
            switch (iType)
            {
                case UrlType://单个对象类型
                    var JsonInfo = JObject.Parse(_sJson);
                    sUrl = JsonInfo["Url"].ToString();
                    sPath = JsonInfo["Path"].ToString();
                    sName = JsonInfo["Name"].ToString();
                    sPath = sPath + "//" + sName;
                    this.DownloadPicture(sUrl, sPath, -1);
                    break;
                case ArrType://数组类型
                    m_bIsStartThread = true;
                    break;
            }
        }
        /// <summary>
        /// url转base64
        /// </summary>
        /// <param name="_sUrl">图片url</param>
        public void UrlToBase64(string _sUrl)
        {
            WebResponse response = null;
            Stream stream = null;
            string sBase64 = null;
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_sUrl);
                response = request.GetResponse();
                stream = response.GetResponseStream();
                if (!response.ContentType.ToLower().StartsWith("text/"))
                    sBase64 = DealPicToBase64(response);
                //回调给前端
                if (OnRecvTempMsg != null)
                {
                    string Json = "{\"Type\":\"Base64\",\"data\"," + sBase64 + " }";
                    OnRecvTempMsg(Json);
                }
            }
            finally
            {
                if (stream != null) stream.Close();
                if (response != null) response.Close();
            }
        }
        /// <summary>
        /// 释放控件内存 关闭线程等
        /// </summary>
        public void DelMemory()
        {
            this.m_DealArrData.Join();
        }
        #endregion

    }
}
