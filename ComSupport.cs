using System;
using System.Runtime.InteropServices;

public static class ComSupport
{
    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

    public static object GetActiveObject(string progId)
    {
        Guid clsid;
        int hr = CLSIDFromProgID(progId, out clsid);
        if (hr < 0) Marshal.ThrowExceptionForHR(hr);

        object obj;
        GetActiveObject(ref clsid, IntPtr.Zero, out obj);
        return obj;
    }
}