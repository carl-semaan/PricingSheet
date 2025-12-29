using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using PricingSheetCore.Models;

namespace PricingSheetCore.Readers
{
    public class Reader
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }

        public Reader() { }

        public Reader(string filePath, string fileName)
        {
            this.FilePath = filePath;
            this.FileName = fileName;
        }
    }
}
