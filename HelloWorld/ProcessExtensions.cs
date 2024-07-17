using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

public static class ProcessExtensions
{
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool TerminateProcess(IntPtr hProcess, uint uExitCode);

    public static void KillTree(this Process process)
    {
        foreach (Process childProcess in GetChildProcesses(process.Id))
        {
            TerminateProcess(childProcess.Handle, 1);
        }
        TerminateProcess(process.Handle, 1);
    }

    private static Process[] GetChildProcesses(int parentId)
    {
        return Process.GetProcesses()
                      .Where(p => p.Id != parentId && GetParentProcessId(p.Id) == parentId)
                      .ToArray();
    }

    private static int GetParentProcessId(int processId)
    {
        IntPtr handle = IntPtr.Zero;
        try
        {
            handle = OpenProcess(ProcessAccessFlags.QueryInformation, false, processId);
            if (handle != IntPtr.Zero)
            {
                PROCESS_BASIC_INFORMATION pbi;
                int status = NtQueryInformationProcess(handle, PROCESSINFOCLASS.ProcessBasicInformation, out pbi, Marshal.SizeOf<PROCESS_BASIC_INFORMATION>(), out _);
                if (status == 0)
                {
                    return pbi.InheritedFromUniqueProcessId.ToInt32();
                }
            }
        }
        finally
        {
            if (handle != IntPtr.Zero)
            {
                CloseHandle(handle);
            }
        }
        return -1;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct PROCESS_BASIC_INFORMATION
    {
        public IntPtr Reserved1;
        public IntPtr PebBaseAddress;
        public IntPtr Reserved2_0;
        public IntPtr Reserved2_1;
        public IntPtr UniqueProcessId;
        public IntPtr InheritedFromUniqueProcessId;
    }

    private enum PROCESSINFOCLASS
    {
        ProcessBasicInformation = 0
    }

    [DllImport("ntdll.dll")]
    private static extern int NtQueryInformationProcess(IntPtr processHandle, PROCESSINFOCLASS processInformationClass, out PROCESS_BASIC_INFORMATION processInformation, int processInformationLength, out int returnLength);

    [DllImport("kernel32.dll")]
    private static extern IntPtr OpenProcess(ProcessAccessFlags desiredAccess, bool inheritHandle, int processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr hObject);

    [Flags]
    private enum ProcessAccessFlags : uint
    {
        QueryInformation = 0x400,
    }
}