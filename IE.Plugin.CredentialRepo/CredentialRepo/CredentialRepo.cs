using IE.Plugin.CredentialRepo.Common;
using Microsoft.Win32;
using mshtml;
using Newtonsoft.Json;
using SHDocVw;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using winform=System.Windows.Forms;

namespace IE.Plugin.CredentialRepo
{
    [
    ComVisible(true),
    Guid("bdd4c782-81f3-45d4-a9ee-9ddd8Bf08c7c"),
    ClassInterface(ClassInterfaceType.None)
    ]
    public class RememberMe : IObjectWithSite
    {
        WebBrowser webBrowser;
        HTMLDocument document;
        Credential credential;
        List<Credential> credentials;
        string autoLoginUrl = null;
        bool allowAutoLogin = true;
        FileWriter writer = new FileWriter(string.Format(@"{0}\log_credentialRepo_{1}.txt", 
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), DateTime.Now.ToString("dd_MMM_yyyy")));

        bool crdentialsLoaded = false;

        string mClickControlId, mClickControlUniqueId, mClickControlTagName;

        #region event handlers
        public void OnDocumentComplete(object pDisp, ref object URL)
        {
            if (URL.Equals("about:blank")) return;

            document = (HTMLDocument)webBrowser.Document;

            credential = new Credential();

            if(!crdentialsLoaded)
                loadCredentials();

            writer.WriteLine(string.Format("OnDocumentComplete :: Doc URL: {0} | param URL: {1}", document.url, URL));
            //mClickControlId = mClickControlUniqueId = mClickControlTagName = null;
            wrapEvents((IHTMLDocument2)document);

            getInputElements(document.all);

            System.Threading.Timer timer = null;
            timer = new System.Threading.Timer((obj) =>
            {
                IHTMLDocument2 docItem = (IHTMLDocument2)obj;
                try
                {
                    timer.Dispose();
                    fillCredentials();
                }
                catch(Exception ex)
                {
                    string msg = ex.Message;
                }
            }, document, 1000, System.Threading.Timeout.Infinite);
        }

        private void wrapEvents(IHTMLDocument2 docItem)
        {
            mshtml.HTMLDocumentEvents2_Event iEvent=null;

            //if (typeof(mshtml.HTMLDocumentEvents2_Event).IsAssignableFrom(typeof(HTMLDocument)))
            iEvent = (mshtml.HTMLDocumentEvents2_Event)docItem;

            if (iEvent != null)
            {
                try
                {
                    iEvent.onclick -= new mshtml.HTMLDocumentEvents2_onclickEventHandler(OnBodyItemClick);
                }
                catch (Exception ex)
                {
                    string errorMsg = ex.Message;
                }
                iEvent.onclick += new mshtml.HTMLDocumentEvents2_onclickEventHandler(OnBodyItemClick);
            }
        }

        private bool OnBodyItemClick(IHTMLEventObj pEvtObj)
        {
            if (pEvtObj.srcElement != null &&
                (pEvtObj.srcElement.tagName.ToLower().Equals("button")
                || (pEvtObj.srcElement.tagName.ToLower().Equals("input") 
                    && ((pEvtObj.srcElement as mshtml.DispHTMLInputElement).type == "button" || (pEvtObj.srcElement as mshtml.DispHTMLInputElement).type == "submit"))
                || pEvtObj.srcElement.tagName.ToLower().Equals("a")
                || (!string.IsNullOrEmpty(pEvtObj.srcElement.className) && pEvtObj.srcElement.className.ToLower().StartsWith("button"))))
            {
                //if (credential == null) credential = new Credential();

                mClickControlId = pEvtObj.srcElement.id ?? pEvtObj.srcElement.tagName.ToLower();
                mClickControlUniqueId = (pEvtObj.srcElement as mshtml.DispHTMLInputElement).className;
                mClickControlTagName = pEvtObj.srcElement.tagName.ToLower();

                if (document!=null && !string.IsNullOrEmpty(document.url))
                    captureCredentials(document.url);
            }
            //else
            //{
            //    if (credential == null) credential = new Credential();

            //    //if (credential.ClickControlId != (pEvtObj.srcElement.id ?? pEvtObj.srcElement.tagName.ToLower()))
            //    {
            //        credential.ClickControlId = pEvtObj.srcElement.id ?? (pEvtObj.srcElement as mshtml.DispHTMLInputElement).name.ToLower();
            //        credential.ClickControlUniqueId = (pEvtObj.srcElement as mshtml.DispHTMLInputElement).className;
            //        credential.ClickControlTagName = pEvtObj.srcElement.tagName.ToLower();
            //    }
            //}

            return true;
        }

        private void fillCredentials(int mode=0, mshtml.HTMLDocument docItem=null)
        {
            bool uidFound, pwdFound;
            uidFound = pwdFound = false;
            string url = null;

            if (docItem != null)
                document = docItem;

            url = webBrowser.LocationURL;

            if (url.IndexOf(';') > -1)
                url = url.Substring(0, url.IndexOf(';'));
            else if (url.IndexOf('?') > -1)
                url = url.Substring(0, url.IndexOf('?'));
            else if (url.IndexOf('#') > -1)
                url = url.Substring(0, url.IndexOf('#'));

            writer.WriteLine(string.Format("{0}-\nHTML:{1}", url, document.body.innerHTML));

            if (credentials != null && credentials.Count > 0)
            {
                credential = credentials.Find(cred => cred.URL.Equals(url, StringComparison.InvariantCultureIgnoreCase));
                if (credential != null)
                {

                    if (document.getElementsByTagName("input").length > 0)
                    {
                        foreach (HTMLInputElement tempElement in document.getElementsByTagName("input"))
                        {
                            if ((tempElement as mshtml.DispHTMLInputElement).id == credential.UIDControlId
                                || (tempElement as mshtml.DispHTMLInputElement).name == credential.UIDControlId) //user id
                            {
                                //(tempElement as mshtml.DispHTMLInputElement).setAttribute("value", credential.UID);
                                //(tempElement as mshtml.DispHTMLInputElement).setAttribute("text", credential.UID);
                                tempElement.focus();
                                //tempElement.value = credential.UID;
                                (tempElement as mshtml.DispHTMLInputElement).setAttribute("value", credential.UID);
                                (tempElement as mshtml.DispHTMLInputElement).setAttribute("text", credential.UID);

                                tempElement.focus();
                                winform.SendKeys.SendWait(" ");
                                winform.SendKeys.SendWait("{BACKSPACE}");
                                System.Threading.Thread.Sleep(400);
                                (tempElement as IHTMLElement3).FireEvent("onchange", null);
                                //(tempElement as IHTMLElement3).FireEvent("onkeypress", null); 

                                uidFound = true;
                            }
                            else if ((tempElement as mshtml.DispHTMLInputElement).id == credential.PWDControlId
                                || (tempElement as mshtml.DispHTMLInputElement).name == credential.PWDControlId) //password
                            {
                                //(tempElement as mshtml.DispHTMLInputElement).setAttribute("value", credential.PWD);
                                //(tempElement as mshtml.DispHTMLInputElement).setAttribute("text", credential.PWD);
                                tempElement.focus();
                                if (!string.IsNullOrEmpty(credential.PWD))
                                {
                                    //tempElement.value = Utility.Decrypt(credential.PWD, true);
                                    (tempElement as mshtml.DispHTMLInputElement).setAttribute("value", Utility.Decrypt(credential.PWD, true));
                                    (tempElement as mshtml.DispHTMLInputElement).setAttribute("text", Utility.Decrypt(credential.PWD, true));
                                }
                                else
                                {
                                    //tempElement.value = credential.PWD;
                                    (tempElement as mshtml.DispHTMLInputElement).setAttribute("value", credential.PWD);
                                    (tempElement as mshtml.DispHTMLInputElement).setAttribute("text", credential.PWD);
                                }

                                tempElement.focus();
                                winform.SendKeys.SendWait(" ");
                                winform.SendKeys.SendWait("{BACKSPACE}");
                                System.Threading.Thread.Sleep(400);
                                (tempElement as IHTMLElement3).FireEvent("onchange", null);
                                //(tempElement as IHTMLElement3).FireEvent("onkeypress", null);

                                pwdFound = true;
                            }

                            //if (uidFound & pwdFound) break;
                        }

                        if (!string.IsNullOrEmpty(credential.ClickControlTagName) && mode == 0 && allowAutoLogin)
                        {
                            if (autoLoginUrl == url) return;

                            foreach (mshtml.DispHTMLInputElement tempElement in document.getElementsByTagName(credential.ClickControlTagName))
                            {
                                if ((tempElement as mshtml.DispHTMLInputElement).id == credential.ClickControlId
                                    || (tempElement as mshtml.DispHTMLInputElement).className == credential.ClickControlUniqueId) //Next page control id
                                {
                                    tempElement.click();
                                    autoLoginUrl = url;
                                    break;
                                }
                            }
                        }

                        //document.getElementById(credential.UIDControlId).setAttribute("value", credential.UID);
                        //document.getElementById(credential.PWDControlId).setAttribute("value", credential.PWD);
                    }
                    else if(document.frames.length>0)
                    {
                        try
                        {
                            mshtml.FramesCollection frames = document.frames;
                            //document = (HTMLDocument)((mshtml.IHTMLFrameBase2)frames.item(0)).contentWindow.document;
                            document = (mshtml.HTMLDocument)((mshtml.HTMLWindow2)frames.item(0)).document;

                            /*((mshtml.DispHTMLDocument)document).onreadystatechange += new mshtml.HTMLDocumentEvents_onreadystatechangeEventHandler(() =>
                            {
                                if (((mshtml.DispHTMLDocument)document).readyState == "complete" || ((mshtml.DispHTMLDocument)document).readyState == "loaded")
                                {
                                    string status = "I have been loaded";
                                }
                            });*/

                            System.Threading.Timer timer = null;
                            timer = new System.Threading.Timer((obj) =>
                            {
                                try
                                {
                                    mshtml.HTMLDocument doc = (mshtml.HTMLDocument)obj;
                                    timer.Dispose();

                                    //Himadri:: ideally it should be called on dom complete but event not getting fired on readystatechanged. Will look into it later
                                    if ((((mshtml.HTMLDocument)doc).readyState == "complete" || ((mshtml.HTMLDocument)doc).readyState == "loaded"))
                                    {
                                        string innerHTML = doc.body.innerHTML;
                                        fillCredentials(mode, doc);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string msg = ex.Message;
                                }
                            }, document, 2000, System.Threading.Timeout.Infinite);
                        }
                        catch(Exception ex1)
                        {
                            string errorMsg = ex1.Message;
                        }
                    }
                }
            }
        }

        public void OnNavigateComplete2(object pDisp, ref object URL)
        {
            document = (HTMLDocument)webBrowser.Document;
            if (autoLoginUrl!=document.url) autoLoginUrl = null;

            credential = new Credential();
            loadCredentials();
            fillCredentials(1);
            //System.Threading.Timer timer = null;
            //timer = new System.Threading.Timer((obj) =>
            //{
            //    fillCredentials();
            //    timer.Dispose();
            //}, null, 1000, System.Threading.Timeout.Infinite);
        }

        private void loadCredentials()
        {
            string path = Environment.CurrentDirectory + @"\credentialStore.json";
            StreamReader reader = null;
            if(File.Exists(path))
            {
                try
                {
                    using (reader = new StreamReader(path))
                    {
                        credentials = JsonConvert.DeserializeObject<List<Credential>>(reader.ReadToEnd());
                    }

                }
                catch (Exception) { }
                finally
                {
                    if (reader != null)
                        reader.Close();

                    crdentialsLoaded = true;
                }
            }
        }

        private void showCredentials(Credential credential)
        {
            if(credential!=null)
                System.Windows.Forms.MessageBox.Show(credential.ToString());
        }

        public void OnBeforeNavigate2(object pDisp, ref object URL, ref object Flags, ref object TargetFrameName, ref object PostData, ref object Headers, ref bool Cancel)
        {
            try
            {
                document = (HTMLDocument)webBrowser.Document;

                captureCredentials(webBrowser.LocationURL);
                //if (credential != null && !string.IsNullOrEmpty(credential.URL))
                //    showCredentials(credential);
            }
            catch(Exception ex)
            {
                string msg = ex.Message;
                //log it here
            }
        }

        private bool HtmlOnChangeEvent_onchange(mshtml.IHTMLEventObj pEvtObj)
        {
            string value = null;
            string pwd, uid;
            string uidControlId, pwdControlId;

            mshtml.HTMLInputTextElement element = pEvtObj.srcElement as mshtml.HTMLInputTextElement;
            //IHTMLInputElement_value = "123qwe"
            if (element.type.ToLower() == "password")//this is for password
            {
                pwd = element.value;
                pwdControlId = element.id;
            }
            else //assume that this is for user id
            {
                uid = element.value;
                uidControlId = element.id;
            }

            return true;
        }
        #endregion

        private List<HTMLInputElement> getInputElements(IHTMLElementCollection elements)
        {
            List<HTMLInputElement> inputElements = new List<HTMLInputElement>();
            foreach (IHTMLElement element in elements)
            {
                if (element.tagName.ToLower()=="input" && ((mshtml.HTMLInputElement)element).type.ToLower()!="hidden")
                    inputElements.Add((HTMLInputElement)element);
                else if(element.tagName.ToLower() == "iframe")
                {
                    try
                    {
                        //document = (mshtml.HTMLDocument)((mshtml.HTMLWindow2)frames.item(0)).document;
                        //((mshtml.HTMLBodyClass)((mshtml.HTMLDocumentClass)((mshtml.HTMLWindow2Class)((mshtml.HTMLIFrameClass)element).contentWindow).document).body).IHTMLElement_innerHTML
                        IHTMLDocument2 doc = ((mshtml.HTMLIFrame)element).contentWindow.document;
                        wrapEvents(doc);
                        //var doc = (mshtml.HTMLDocument)element.document;
                        if (doc.readyState == "complete" || doc.readyState == "loaded")
                        {
                            inputElements.AddRange(getInputElements(doc.all));
                        }
                    }
                    catch(Exception ex)
                    {
                        string errorMsg = ex.Message;
                    }
                }

            }
            return inputElements;
        }

        private void captureCredentials(string url)
        {
            string uid, pwd;
            string uidControlId, pwdControlId;

            //var elements = (webBrowser.Document as HTMLDocument).getElementsByTagName("input");
            //var elements = document.getElementsByTagName("input");

            var elements = getInputElements(document.all);

            Uri uri = null;
            if (url.IndexOf(';') > -1)
                url = url.Substring(0, url.IndexOf(';'));
            else if (url.IndexOf('?') > -1)
                url = url.Substring(0, url.IndexOf('?'));
            else if (url.IndexOf('#') > -1)
                url = url.Substring(0, url.IndexOf('#'));

            Uri.TryCreate(url, UriKind.RelativeOrAbsolute, out uri);

            //if (uri != browser1.Url) return;

            //if(credential==null)
            credential = new Credential();

            credential.ClickControlId = mClickControlId;
            credential.ClickControlTagName = mClickControlTagName;
            credential.ClickControlUniqueId=mClickControlUniqueId;

            int idx = -1;
            int elementIdx = -1;
            HTMLInputElement inputElement = null;
            HTMLInputElement lastInputElement = null;
            HTMLInputElement prevInputElement = null;

            foreach (HTMLInputElement element in elements)
            {
                idx++;
                prevInputElement = inputElement;
                inputElement = element; //((element as HtmlElement).DomElement as mshtml.HTMLInputElement);
                if (inputElement.type.ToLower().Equals("password"))
                {
                    pwd = inputElement.value;
                    pwdControlId = inputElement.id ?? inputElement.name;
                    credential.URL = url;
                    if(pwd!=null)
                        credential.PWD = Utility.Encrypt(pwd, true);
                    credential.PWDControlId = pwdControlId;
                    credential.PWDControlUniqueId = (inputElement as DispHTMLInputElement).className;

                    //HTMLInputTextElementEvents2_Event htmlOnChangeEvent = (inputElement as DispHTMLInputElement) as HTMLInputTextElementEvents2_Event;
                    //htmlOnChangeEvent.onchange += HtmlOnChangeEvent_onchange;

                    elementIdx = idx;
                    lastInputElement = prevInputElement;
                    if (!string.IsNullOrEmpty(pwd)) //issues here
                        break;
                }
            }

            if (elementIdx > -1) //we have some password entered
            {
                inputElement = lastInputElement; //((elements[elementIdx - 1] as HtmlElement).DomElement as mshtml.HTMLInputElement);

                if (inputElement != null)
                {
                    uid = inputElement.value;
                    uidControlId = inputElement.id ?? inputElement.name;
                    credential.UID = uid;
                    credential.UIDControlId = uidControlId;
                    credential.UIDControlUniqueId = (inputElement as mshtml.DispHTMLInputElement).className;
                    //mshtml.HTMLInputTextElementEvents2_Event htmlOnChangeEvent = (inputElement as mshtml.DispHTMLInputElement) as mshtml.HTMLInputTextElementEvents2_Event;
                    //htmlOnChangeEvent.onchange += HtmlOnChangeEvent_onchange;
                }
            }

            try
            {
                if (credential != null && !string.IsNullOrEmpty(credential.URL) 
                    && (!string.IsNullOrEmpty(credential.UID)
                    || !string.IsNullOrEmpty(credential.PWD)))
                {
                    if (credentials == null) credentials = new List<Credential>();

                    Credential existingCredential = credentials.Find(cred => cred.URL.Equals(credential.URL));
                    if (existingCredential==null)
                    {
                        credentials.Add(credential);
                    }
                    else
                    {
                        existingCredential.UID = credential.UID?? existingCredential.UID;
                        existingCredential.PWD = credential.PWD?? existingCredential.PWD;
                        existingCredential.PWDControlId = existingCredential.PWDControlId ?? credential.PWDControlId;
                        existingCredential.UIDControlId = existingCredential.UIDControlId ?? credential.UIDControlId;
                        existingCredential.PWDControlUniqueId = existingCredential.PWDControlUniqueId ?? credential.PWDControlUniqueId;
                        existingCredential.UIDControlUniqueId = existingCredential.UIDControlUniqueId ?? credential.UIDControlUniqueId;
                        existingCredential.ClickControlId = existingCredential.ClickControlId ?? credential.ClickControlId;
                        existingCredential.ClickControlTagName= existingCredential.ClickControlTagName ?? credential.ClickControlTagName;
                        existingCredential.ClickControlUniqueId= existingCredential.ClickControlUniqueId ?? credential.ClickControlUniqueId;
                    }
                    //Always update the list to file
                    FileWriter writer = new FileWriter();
                    writer.Write<List<Credential>>(credentials, FormatType.JSON);
                    //credential = new Credential();
                }
            }
            catch(Exception)
            {
                //log here
            }
        }

        #region Internal helper functions
        public static string BHOKEYNAME = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";

        [ComRegisterFunction]
        public static void RegisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BHOKEYNAME, true);

            if (registryKey == null)
                registryKey = Registry.LocalMachine.CreateSubKey(BHOKEYNAME);

            string guid = type.GUID.ToString("B");
            RegistryKey ourKey = registryKey.OpenSubKey(guid);

            if (ourKey == null)
                ourKey = registryKey.CreateSubKey(guid);

            ourKey.SetValue("Alright", 1);
            registryKey.Close();
            ourKey.Close();
        }

        [ComUnregisterFunction]
        public static void UnregisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BHOKEYNAME, true);
            string guid = type.GUID.ToString("B");

            if (registryKey != null)
                registryKey.DeleteSubKey(guid, false);
        }

        public int SetSite(object site)
        {

            if (site != null)
            {
                try
                {
                    webBrowser = (WebBrowser)site;
                    webBrowser.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
                    webBrowser.BeforeNavigate2 += new DWebBrowserEvents2_BeforeNavigate2EventHandler(this.OnBeforeNavigate2);
                    webBrowser.NavigateComplete2 += new DWebBrowserEvents2_NavigateComplete2EventHandler(this.OnNavigateComplete2);
                    //webBrowser.WindowStateChanged += new DWebBrowserEvents2_WindowStateChangedEventHandler(this.OnWindowStateChanged);
                }
                finally { }
            }
            else
            {
                try
                {
                    webBrowser.DocumentComplete -= new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
                    webBrowser.BeforeNavigate2 -= new DWebBrowserEvents2_BeforeNavigate2EventHandler(this.OnBeforeNavigate2);
                    webBrowser.NavigateComplete2 -= new DWebBrowserEvents2_NavigateComplete2EventHandler(this.OnNavigateComplete2);
                    //webBrowser.WindowStateChanged -= new DWebBrowserEvents2_WindowStateChangedEventHandler(this.OnWindowStateChanged);

                    webBrowser = null;
                }
                finally { }
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
        #endregion
    }
}
