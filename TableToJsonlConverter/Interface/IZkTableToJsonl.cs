using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableToJsonlConverter.Interface
{
    internal interface IZkTableToJsonl
    {
        void Input();
        void Output();

        string JsonLines { get; }
    }
}
