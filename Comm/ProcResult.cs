using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoborSmartReport.Comm
{
    public class ProcResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }

        public static ProcResult Ok(string message = "")
        {
            return new ProcResult
            {
                Success = true,
                Message = message
            };
        }

        public static ProcResult Fail(string message)
        {
            return new ProcResult
            {
                Success = false,
                Message = message
            };
        }
    }
}
