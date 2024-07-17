using System;

namespace B1PostingDataText.B1Connection
{
    public class CleanUp
    {
        public static void CleanUpServices(object obj)
        {
            if (obj != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
        }

        public static void CleanUpGCCollect()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
