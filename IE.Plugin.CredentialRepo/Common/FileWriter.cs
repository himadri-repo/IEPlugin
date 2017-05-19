using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IE.Plugin.CredentialRepo.Common
{
    public enum FormatType
    {
        CSV=0,
        JSON=1
    }

    public class FileWriter : IWriter
    {
        private string FileName = null;
        public FileWriter()
        {
            
        }

        public FileWriter(string fileName)
        {
            this.FileName = fileName;
        }

        public bool WriteLine(string data)
        {
            bool bflag = true;
            StreamWriter streamWriter = null;

            try
            {
                using (streamWriter = new StreamWriter(FileName, true, ASCIIEncoding.UTF8))
                {
                    streamWriter.WriteLine(data);
                    streamWriter.Close();
                }

                bflag = true;
            }
            catch (Exception ex)
            {
                //log error
                string error = ex.Message;
            }
            finally
            {
                //log finalization
            }

            return bflag;
        }
        public bool Write<T>(T entity, FormatType formatType) where T : List<Credential>,new()
        {
            bool bflag = false;
            string path = Application.ExecutablePath;

            if (formatType == FormatType.CSV)
                FileName = Environment.CurrentDirectory + @"\credentialStore.csv";
            else if (formatType == FormatType.JSON)
                FileName = Environment.CurrentDirectory + @"\credentialStore.json";

            StreamWriter streamWriter = null;

            try
            {
                using (streamWriter = new StreamWriter(FileName, false, ASCIIEncoding.UTF8))
                {
                    if (formatType == FormatType.CSV)
                        ToCSV(entity, streamWriter);
                    else if (formatType == FormatType.JSON)
                        ToJSON(entity, streamWriter);
                    streamWriter.Close();
                }

                bflag = true;
            }
            catch(Exception ex)
            {
                //log error
                string error = ex.Message;
            }
            finally
            {
                //log finalization
            }

            return bflag;
        }

        public bool ToCSV(List<Credential> entities, StreamWriter writer)
        {
            bool bflag = false;
            StringBuilder builder = new StringBuilder();

            try
            {
                foreach (Credential credential in entities)
                {
                    builder.AppendLine(string.Format("{0},{1},{2},{3},{4}", credential.URL, credential.UID, credential.UIDControlId, credential.PWD, credential.PWDControlId));
                    writer.Write(builder.ToString());
                }
                writer.Flush();
                bflag = true;
            }
            catch (Exception)
            {
                //log exception here
            }
            finally
            {
                //log status here
            }

            return bflag;
        }

        public bool ToJSON(List<Credential> entities, StreamWriter writer)
        {
            bool bflag = false;
            StringBuilder builder = new StringBuilder();

            try
            {
                writer.Write(JsonConvert.SerializeObject(entities, Formatting.Indented));
                bflag = true;
            }
            catch (Exception)
            {
                //log exception here
            }
            finally
            {
                //log status here
            }

            return bflag;
        }
    }
}
