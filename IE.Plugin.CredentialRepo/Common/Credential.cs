using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IE.Plugin.CredentialRepo.Common
{
    public class Credential
    {
        public string URL { get; set; }
        public string UID { get; set; }
        public string PWD { get; set; }
        public string UIDControlId { get; set; }
        public string UIDControlUniqueId { get; set; }
        public string PWDControlId { get; set; }
        public string PWDControlUniqueId { get; set; }
        public string ClickControlId { get; set; }
        public string ClickControlUniqueId { get; set; }
        public string ClickControlTagName { get; set; }

        public Credential()
        {
            this.URL = string.Empty;
            this.UID = string.Empty;
            this.UIDControlId = string.Empty;
            this.PWD = string.Empty;
            this.PWDControlId = string.Empty;

            this.UIDControlUniqueId = this.PWDControlUniqueId = this.ClickControlUniqueId = this.ClickControlTagName = string.Empty;
        }

        public override string ToString()
        {
            StringBuilder strReturnMessage = new StringBuilder();

            strReturnMessage.AppendLine(string.Format("File Path : {0}", Environment.CurrentDirectory + @"\credentialStore.json"));
            strReturnMessage.AppendLine(string.Format("URL : {0}", this.URL));
            strReturnMessage.AppendLine(string.Format("UID : {0} | UIDControlId : {1}", this.UID, this.UIDControlId));
            strReturnMessage.AppendLine(string.Format("PWD : {0} | PWDControlId : {1}", this.PWD, this.PWDControlId));
            strReturnMessage.AppendLine(string.Format("Click Control Id : {0}", this.ClickControlId));

            return strReturnMessage.ToString();
        }
    }
}
