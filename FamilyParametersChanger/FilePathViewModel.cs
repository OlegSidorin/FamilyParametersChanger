using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace FamilyParametersChanger
{
    public class FilePathViewModel
    {
        public string PathString { get; set; }
        public string Name { get; set; }
        public string ImgSource { get; set; }
        public Thickness LeftMargin { get; set; }
        public FileType FileType { get; set; }
    }
}
