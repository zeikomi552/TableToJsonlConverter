using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableToJsonlConverter.Interface
{
    internal interface IZkTableToJsonl
    {
        void Read();
        void Write(string path);

        string JsonLines { get; }
    }
}
