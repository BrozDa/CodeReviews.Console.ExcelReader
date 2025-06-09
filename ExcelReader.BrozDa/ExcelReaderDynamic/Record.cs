using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderDynamic
{
    internal class Record
    {
        public int Id { get; set; }
        public Dictionary<string, string> Data { get; set; } = new();
    }
}
