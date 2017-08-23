using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSC_3REmailAutomation
{
    class Employee
    {
        public string Name { get; set; }
        public string Number { get; set; }
        public string Employee_ID { get; set; }
        public string Email_ID { get; set; }
    }
    class EmpConstants
    {
        private const string DOMAIN_NAME = "xyz.com";
    }
}
