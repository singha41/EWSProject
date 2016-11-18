using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Exchange.Samples.EWS
{
    class Tracing
    {
        private static TextWriter logFileWriter = null;

        public static void OpenLog(string fileName)
        {
            logFileWriter = new StreamWriter(fileName);
        }

        public static void Write(string format, params object[] args)
        {
            Console.Write(format, args);
            if (logFileWriter != null)
            {
                logFileWriter.Write(format, args);
            }
        }

        public static void WriteLine(string format, params object[] args)
        {
            Console.WriteLine(format, args);
            if (logFileWriter != null)
            {
                logFileWriter.WriteLine(format, args);
            }
        }

        public static void CloseLog()
        {
            logFileWriter.Flush();
            logFileWriter.Close();
            //Features Branch 1.1 Chnage 
        }
    }
}
