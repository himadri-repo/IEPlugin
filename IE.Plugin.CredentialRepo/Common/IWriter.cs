using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IE.Plugin.CredentialRepo.Common
{
    public interface IWriter
    {
        bool Write<T>(T entity, FormatType formatType) where T : List<Credential>, new();
    }

    public interface IEntity
    {
        bool ToCSV(StreamWriter writer);
        bool ToJSON(StreamWriter writer);
    }
}
