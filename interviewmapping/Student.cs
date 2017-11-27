using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace interviewmapping
{
    class Student
    {
        public string Name { get; set; } = "";

        public string Class { get; set; } = "";

        public Dictionary<int,string> MappingCompany = new Dictionary<int, string>();
    }
}
