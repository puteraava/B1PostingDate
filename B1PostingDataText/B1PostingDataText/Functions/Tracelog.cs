using System;
using System.IO;
using System.Windows.Forms;

namespace B1PostingDataText.Functions
{
    public class Tracelog
    {
        static public void TransWriteLine(string value)
        {
            try
            {
                string log = Application.StartupPath + @"\TransServices.log";
                using (var writer = File.AppendText(log))
                {
                    writer.WriteLine($"[{DateTime.Now:dd-MM-yyyy HH:mm:ss}] "+value);
                }
            }
            catch (Exception e)
            {
                e.ToString();
            }
        }
    }
}