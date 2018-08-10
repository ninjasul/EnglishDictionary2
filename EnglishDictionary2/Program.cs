namespace EnglishDictionary2
{
    using System;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    internal static class Program
    {
        // 콘솔 디버깅을 위한 코드
        /*
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();
        */

        [STAThread]
        private static void Main()
        {
            // 콘솔창 생성
            //AllocConsole();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new EnglishWordBook());
        }
    }
}

