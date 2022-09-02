using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

class CSburnCD
{
    [DllImport("shfolder.dll")]
    static extern int SHGetFolderPath(IntPtr hwndOwner, int nFolder,
                                      IntPtr hToken, int dwFlags,
                                      StringBuilder pszPath);

    const int CSIDL_CDBURN_AREA = 0x3B;

    const int SHGFP_TYPE_CURRENT = 0;

    public static void CSharpBurnCD()
    {
        StringBuilder szPath = new StringBuilder(1024);

        //szPath.Append( dataPath);

        if (SHGetFolderPath((IntPtr)0, CSIDL_CDBURN_AREA, (IntPtr)0,
            SHGFP_TYPE_CURRENT, szPath) != 0)
            Console.WriteLine("SHGetFolderPath() failure");
        else
            Console.WriteLine("SHGetFolderPath return value = " + szPath);

        Guid CLSID_CDBurn = new Guid("fbeb8a05-beee-4442-804e-409d6c4515e9");

        Type t = Type.GetTypeFromCLSID(CLSID_CDBurn);
        if (t == null)
        {
            Console.WriteLine("ICDBurn not supported by OS");
            return;
        }

        ICDBurn iface = (ICDBurn)Activator.CreateInstance(t);
        if (iface == null)
        {
            Console.WriteLine("Unable to obtain interface");
            return;
        }

        bool hasRecorder = false;
        iface.HasRecordableDrive(ref hasRecorder);
        Console.WriteLine("HasRecordableDrive return value = " + hasRecorder);

        if (hasRecorder)
        {
            StringBuilder driveLetter = new StringBuilder(4);
            iface.GetRecorderDriveLetter(driveLetter, 4);
            Console.WriteLine("GetRecorderDriveLetter return value = " +
                              driveLetter);
            iface.Burn((IntPtr)0);

            
        }
    }
}

[ComImport]
[Guid("3d73a659-e5d0-4d42-afc0-5121ba425c8d")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ICDBurn
{
    void GetRecorderDriveLetter([MarshalAs(UnmanagedType.LPWStr)]
                                StringBuilder pszDrive, uint cch);
    void Burn(IntPtr hwnd);
    void HasRecordableDrive(ref bool HasRecorder);
}