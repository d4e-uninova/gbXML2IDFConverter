using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace gbXML2IDFConverter
{
    //Design Builder 4.5.0.148
    public class gbXML2IDF
    {
        private static Logger logger = new Logger("gbXML2IDF");

        [DllImport("FindAddress.dll")]
        private static extern int FindAddress(int pid, string data, IntPtr startAddress, IntPtr endAddress, out IntPtr ptr);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool WriteProcessMemory(
            IntPtr hProcess,
            IntPtr lpBaseAddress,
            byte[] lpBuffer,
            int nSize,
            out IntPtr lpNumberOfBytesWritten);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool CloseHandle(IntPtr hHandle);
        [DllImport("kernel32.dll")]
        static extern IntPtr OpenProcess(
             ProcessAccessFlags processAccess,
             bool bInheritHandle,
             int processId
        );


        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsWindowEnabled(IntPtr hWnd);

        [DllImport("USER32.DLL", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("USER32.DLL")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, uint wParam, uint lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr PostMessage(IntPtr hWnd, uint msg, uint wParam, uint lParam);

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr i);

        private delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);

        [Flags]
        public enum ProcessAccessFlags : uint
        {
            All = 0x001F0FFF,
            Terminate = 0x00000001,
            CreateThread = 0x00000002,
            VirtualMemoryOperation = 0x00000008,
            VirtualMemoryRead = 0x00000010,
            VirtualMemoryWrite = 0x00000020,
            DuplicateHandle = 0x00000040,
            CreateProcess = 0x000000080,
            SetQuota = 0x00000100,
            SetInformation = 0x00000200,
            QueryInformation = 0x00000400,
            QueryLimitedInformation = 0x00001000,
            Synchronize = 0x00100000
        }

       
       
        /// <summary>
        /// Callback method to be used when enumerating windows.
        /// </summary>
        /// <param name="handle">Handle of the next window</param>
        /// <param name="pointer">Pointer to a GCHandle that holds a reference to the list to fill</param>
        /// <returns>True to continue the enumeration, false to bail</returns>
        private static bool EnumWindow(IntPtr handle, IntPtr pointer)
        {
            GCHandle gch = GCHandle.FromIntPtr(pointer);
            List<IntPtr> list = gch.Target as List<IntPtr>;
            if (list == null)
            {
                throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
            }
            list.Add(handle);
            //  You can modify this to check to see if you want to cancel the operation, then return a null here
            return true;
        }

        /// <summary>
        /// Returns a list of child windows
        /// </summary>
        /// <param name="parent">Parent of the windows to return</param>
        /// <returns>List of child windows</returns>
        private static List<IntPtr> GetChildWindows(IntPtr parent)
        {
            List<IntPtr> result = new List<IntPtr>();
            GCHandle listHandle = GCHandle.Alloc(result);
            try
            {
                EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle));
            }
            finally
            {
                if (listHandle.IsAllocated)
                    listHandle.Free();
            }
            return result;
        }

        private const int WM_COMMAND = 0x111;
        private const int WM_LBUTTONDOWN = 0x201;
        private const int WM_LBUTTONUP = 0x202;
        private const int HIDE = 0x5;

        private string designBuilderPath;

        public gbXML2IDF(string designBuilderPath)
        {
            this.designBuilderPath = designBuilderPath;
        }
        public static void WriteMem(Process p, IntPtr address, string data)
        {
            var hProc = OpenProcess(ProcessAccessFlags.All, false, (int)p.Id);
            byte[] dataByte = Encoding.ASCII.GetBytes(data);
            IntPtr nl = IntPtr.Zero;
            WriteProcessMemory(hProc, address, dataByte, dataByte.Length, out nl);
            CloseHandle(hProc);
        }


        private void tryToKill(Process process)
        {
            try
            {
                logger.log("Closing process '" + process.ProcessName + "' with id " + process.Id);
                process.Kill();
                process.WaitForExit();
            }
            catch(Exception ex)
            {
                logger.log(ex.Message,Logger.LogType.ERROR);
            }
        }

        public bool Convert(string inFile,string outFile)
        {

            Process[] processes;
            try
            {
                bool isDirectory = Directory.Exists(outFile);

                processes = Process.GetProcessesByName("DesignBuilder");

                for (var i = 0; i < processes.Length; i++)
                {
                    tryToKill(processes[i]);
                }

                if (isDirectory)
                {
                    outFile = outFile + "\\" + Path.ChangeExtension(Path.GetFileName(inFile), "idf");
                }


                string tmpIdfFile = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\DesignBuilder\EnergyPlus\in.idf";
                if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\DesignBuilder\"))
                {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\DesignBuilder\");
                }
                if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\DesignBuilder\EnergyPlus"))
                {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\DesignBuilder\EnergyPlus\");
                }
                Process process;
                IntPtr mainWindowHandle = IntPtr.Zero;
                IntPtr hWnd = IntPtr.Zero;
                IntPtr stringPtr = IntPtr.Zero;
                IntPtr childHWnd = IntPtr.Zero;
                IntPtr childCompare = IntPtr.Zero;
                double deltaT;
                int timeOut = 240;
                List<IntPtr> childs = new List<IntPtr>();
                string tmpDir = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.System)) + "dtmp\\";
                string tmpXmlFile = tmpDir + "t.xml";
                int n = 0;
                IntPtr arrHWnd;
                DateTime t1;
                IntPtr[] ptrArr;
                int arrLen;
                if (File.Exists(outFile))
                {
                    File.Delete(outFile);
                }
                process = new Process();
                process.StartInfo.FileName = this.designBuilderPath + @"\DesignBuilder.exe";
                process.Start();
                logger.log("Process '" + process.ProcessName + "' with id " + process.Id + " started");
                do
                {
                    Thread.Sleep(1);
                    hWnd = FindWindow("ThunderRT6FormDC", "");
                } while (hWnd == IntPtr.Zero || hWnd != IntPtr.Zero && GetChildWindows(hWnd).Count() == 0);

                while (GetChildWindows(hWnd).Count() > 0)
                {
                    Thread.Sleep(1);
                    ShowWindow(hWnd, HIDE);
                }

                do
                {
                    Thread.Sleep(1);
                    mainWindowHandle = FindWindow("ThunderRT6FormDC", "DesignBuilder");
                } while (mainWindowHandle == IntPtr.Zero);
                do
                {
                    Thread.Sleep(1);
                } while (!ShowWindow(mainWindowHandle, HIDE));
                hWnd = IntPtr.Zero;
                Thread.Sleep(1000);
                PostMessage(mainWindowHandle, WM_COMMAND, 0x7300004, (uint)GetChildWindows(mainWindowHandle)[40]);
                do
                {
                    Thread.Sleep(1);
                    hWnd = FindWindow("ThunderRT6FormDC", "Please wait");
                } while (hWnd == IntPtr.Zero);


                do
                {
                    ShowWindow(hWnd, HIDE);
                    Thread.Sleep(1);
                    hWnd = FindWindow("ThunderRT6FormDC", "Please wait");
                } while (hWnd != IntPtr.Zero);

                do
                {
                    Thread.Sleep(1);
                    hWnd = FindWindow("ThunderRT6FormDC", "Please wait");
                } while (hWnd == IntPtr.Zero);

                do
                {
                    ShowWindow(hWnd, HIDE);
                    Thread.Sleep(1);
                    hWnd = FindWindow("ThunderRT6FormDC", "Please wait");
                } while (hWnd != IntPtr.Zero);
                PostMessage(mainWindowHandle, WM_COMMAND, 0x7300003, (uint)GetChildWindows(mainWindowHandle)[40]);
                hWnd = IntPtr.Zero;
                do
                {
                    Thread.Sleep(1);
                    hWnd = FindWindow("#32770", "DesignBuilder");
                } while (hWnd == IntPtr.Zero);
                ShowWindow(hWnd, HIDE);
                PostMessage(hWnd, WM_COMMAND, 0x6, 0x20061C);

                hWnd = IntPtr.Zero;

                do
                {
                    Thread.Sleep(1);
                    hWnd = FindWindow("ThunderRT6FormDC", "Please wait");
                } while (hWnd == IntPtr.Zero);
                do
                {
                    Thread.Sleep(1);
                    hWnd = FindWindow("ThunderRT6FormDC", "Please wait");
                    ShowWindow(hWnd, HIDE);
                } while (hWnd != IntPtr.Zero);

                do
                {
                    Thread.Sleep(1);
                    hWnd = FindWindow("ThunderRT6FormDC", "Import BIM Model");
                } while (hWnd == IntPtr.Zero);
                ShowWindow(hWnd, HIDE);
                arrLen = FindAddress(process.Id, "<Select file>", (IntPtr)0x1B000000, (IntPtr)0x20000000, out arrHWnd);
                ptrArr = new IntPtr[arrLen];
                Marshal.Copy(arrHWnd, ptrArr, 0, arrLen);
                if (!Directory.Exists(tmpDir))
                {
                    Directory.CreateDirectory(tmpDir);
                }
                if (File.Exists(tmpXmlFile))
                {
                    File.Delete(tmpXmlFile);
                }
                File.Copy(inFile, tmpXmlFile);
                for (var i = 0; i < ptrArr.Length; i++)
                {
                    WriteMem(process, ptrArr[i], tmpXmlFile);
                }
                PostMessage(GetChildWindows(hWnd)[50], WM_COMMAND, 0x12, (uint)GetChildWindows(hWnd)[52]);
                t1 = DateTime.Now;
                logger.log("Importing gbXML file...");
                do
                {
                    ShowWindow(hWnd, HIDE);
                    Thread.Sleep(1);
                    deltaT = (DateTime.Now - t1).TotalSeconds;
                    if (deltaT > timeOut || FindWindow("#32770", "DesignBuilder") != IntPtr.Zero)
                    {
                        logger.log("Error importing gbXML file", Logger.LogType.ERROR);
                        tryToKill(process);
                        return false;
                    }
                } while (!IsWindowEnabled(GetChildWindows(hWnd)[54]));
                PostMessage(GetChildWindows(hWnd)[50], WM_COMMAND, 0x12, (uint)GetChildWindows(hWnd)[54]);
                hWnd = IntPtr.Zero;
                do
                {
                    Thread.Sleep(1);
                    hWnd = FindWindow("ThunderRT6FormDC", "Please wait");
                } while (hWnd == IntPtr.Zero);
                do
                {
                    Thread.Sleep(1);
                    hWnd = FindWindow("ThunderRT6FormDC", "Please wait");
                    ShowWindow(hWnd, HIDE);
                } while (hWnd != IntPtr.Zero);

                hWnd = IntPtr.Zero;
                t1 = DateTime.Now;
                do
                {
                    do
                    {
                        Thread.Sleep(1);
                        hWnd = FindWindow("ThunderRT6FormDC", "Please wait");
                    } while (hWnd == IntPtr.Zero && (DateTime.Now - t1).TotalSeconds < 2);
                    if ((DateTime.Now - t1).Seconds < 2)
                    {
                        do
                        {
                            Thread.Sleep(1);
                            t1 = DateTime.Now;
                            hWnd = FindWindow("ThunderRT6FormDC", "Please wait");
                            ShowWindow(hWnd, HIDE);
                        } while (hWnd != IntPtr.Zero);
                    }
                } while ((DateTime.Now - t1).Seconds < 2);
                hWnd = IntPtr.Zero;
                do
                {
                    hWnd = FindWindow(null, "Import Messages");
                    Thread.Sleep(1);
                } while (hWnd == IntPtr.Zero);
                SendMessage(GetChildWindows(hWnd)[0], WM_LBUTTONDOWN, 0x1, 0x100030);
                SendMessage(GetChildWindows(hWnd)[0], WM_LBUTTONUP, 0x0, 0x100030);
                Thread.Sleep(1000);
                if (File.Exists(tmpIdfFile))
                {
                    File.Delete(tmpIdfFile);
                }
                childs = GetChildWindows(mainWindowHandle);
                for (var i = 0; i < childs.Count; i++)
                {
                    if (GetChildWindows(childs[i]).Count() == 1)
                    {
                        PostMessage(GetChildWindows(childs[i])[0], WM_LBUTTONDOWN, 0x1, 0x9013B);
                        PostMessage(GetChildWindows(childs[i])[0], WM_LBUTTONUP, 0x1, 0x9013B);
                    }
                }
                logger.log("Generating IDF file...");
                hWnd = IntPtr.Zero;
                t1 = DateTime.Now;
                do
                {
                    hWnd = FindWindow("ThunderRT6FormDC", "Please wait");
                    if (hWnd != IntPtr.Zero)
                    {
                        ShowWindow(hWnd, HIDE);
                    }
                    Thread.Sleep(1);
                    deltaT = (DateTime.Now - t1).TotalSeconds;
                    if (deltaT > timeOut || FindWindow("#32770", "DesignBuilder") != IntPtr.Zero)
                    {
                        logger.log("Error generating IDF file", Logger.LogType.ERROR);
                        tryToKill(process);
                        return false;
                    }
                } while (!File.Exists(tmpIdfFile));
                n = 0;
                do
                {
                    hWnd = FindWindow("ThunderRT6FormDC", "Please wait");
                    if (hWnd != IntPtr.Zero)
                    {
                        ShowWindow(hWnd, HIDE);
                    }
                    n++;
                    Thread.Sleep(10);
                } while (n < 200);
                tryToKill(process);
                processes = Process.GetProcessesByName("RunEPDLL");
                for (var i = 0; i < processes.Length; i++)
                {
                    tryToKill(processes[i]);
                }

                processes = Process.GetProcessesByName("DesignBuilderEnergyPlus");
                for (var i = 0; i < processes.Length; i++)
                {
                    tryToKill(processes[i]);
                }

                File.Copy(tmpIdfFile, outFile);
                if (File.Exists(tmpXmlFile))
                {
                    File.Delete(tmpXmlFile);
                }
                if (Directory.Exists(tmpDir))
                {
                    Directory.Delete(tmpDir, true);
                }
                return true;
            }
            catch(Exception e)
            {
                
                logger.log(e.Message, Logger.LogType.ERROR);
                processes = Process.GetProcessesByName("DesignBuilder");
                for (var i = 0; i < processes.Length; i++)
                {
                    tryToKill(processes[i]);
                }
                processes = Process.GetProcessesByName("RunEPDLL");
                for (var i = 0; i < processes.Length; i++)
                {
                    tryToKill(processes[i]);
                }

                processes = Process.GetProcessesByName("DesignBuilderEnergyPlus");
                for (var i = 0; i < processes.Length; i++)
                {
                    tryToKill(processes[i]);
                }
                return false;
            }
            
        }

    }
}
