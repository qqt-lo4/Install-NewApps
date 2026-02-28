function Set-WindowIcon {
    <#
    .SYNOPSIS
        Sets the icon for a WPF window (title bar and taskbar)

    .DESCRIPTION
        Sets the window icon in both the title bar and taskbar by:
        1. Setting the WPF Icon property for the title bar
        2. Using IPropertyStore to set a unique AppUserModelID to separate from the host process
        3. Using WM_SETICON to set the taskbar icon

    .PARAMETER Window
        The WPF Window object to set the icon for

    .PARAMETER IconPath
        The absolute path to the icon file (.ico)

    .PARAMETER AppId
        Optional AppUserModelID to identify the window separately from the host process.
        If not specified, a unique GUID is generated automatically.

    .EXAMPLE
        Set-WindowIcon -Window $window -IconPath "C:\Icons\app.ico"

    .EXAMPLE
        Set-WindowIcon -Window $window -IconPath "C:\Icons\app.ico" -AppId "MyCompany.MyApp.MainWindow"

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Window]$Window,

        [Parameter(Mandatory = $true)]
        [string]$IconPath,

        [Parameter(Mandatory = $false)]
        [string]$AppId = ""
    )

    # Load required assemblies
    Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
    Add-Type -AssemblyName PresentationCore -ErrorAction SilentlyContinue
    Add-Type -AssemblyName WindowsBase -ErrorAction SilentlyContinue

    if (-not $IconPath -or -not (Test-Path $IconPath)) { return }

    # Generate unique AppId if not provided
    if (-not $AppId) {
        $AppId = [System.Guid]::NewGuid().ToString()
    }

    try {
        # Set WPF icon for title bar
        $uri = [System.Uri]::new($IconPath, [System.UriKind]::Absolute)
        $bitmap = [System.Windows.Media.Imaging.BitmapFrame]::Create($uri)
        $Window.Icon = $bitmap

        # Get window handle and set taskbar icon via IPropertyStore
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        $hwnd = $helper.EnsureHandle()

        # Define COM interfaces and P/Invoke for setting AppUserModelID on window
        # Check if type already exists to avoid conflicts
        if (-not ([System.Management.Automation.PSTypeName]'WindowIconAppId').Type) {
            Add-Type -TypeDefinition @"
                using System;
                using System.Runtime.InteropServices;

                public class WindowIconAppId {
                    [DllImport("shell32.dll")]
                    private static extern int SHGetPropertyStoreForWindow(
                        IntPtr hwnd,
                        ref Guid iid,
                        [MarshalAs(UnmanagedType.Interface)] out IPropertyStore propertyStore);

                    [ComImport]
                    [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
                    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
                    private interface IPropertyStore {
                        int GetCount(out uint cProps);
                        int GetAt(uint iProp, out PROPERTYKEY pkey);
                        int GetValue(ref PROPERTYKEY key, out PropVariant pv);
                        int SetValue(ref PROPERTYKEY key, ref PropVariant pv);
                        int Commit();
                    }

                    [StructLayout(LayoutKind.Sequential, Pack = 4)]
                    private struct PROPERTYKEY {
                        public Guid fmtid;
                        public uint pid;
                    }

                    [StructLayout(LayoutKind.Explicit)]
                    private struct PropVariant {
                        [FieldOffset(0)] public ushort vt;
                        [FieldOffset(8)] public IntPtr pwszVal;

                        public static PropVariant FromString(string value) {
                            var pv = new PropVariant();
                            pv.vt = 31; // VT_LPWSTR
                            pv.pwszVal = Marshal.StringToCoTaskMemUni(value);
                            return pv;
                        }
                    }

                    private static readonly Guid IPropertyStoreGuid = new Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99");

                    private static readonly PROPERTYKEY PKEY_AppUserModel_ID = new PROPERTYKEY {
                        fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"),
                        pid = 5
                    };

                    public static void SetAppUserModelId(IntPtr hwnd, string appId) {
                        IPropertyStore propStore;
                        Guid guid = IPropertyStoreGuid;
                        int hr = SHGetPropertyStoreForWindow(hwnd, ref guid, out propStore);
                        if (hr == 0 && propStore != null) {
                            var pv = PropVariant.FromString(appId);
                            var key = PKEY_AppUserModel_ID;
                            propStore.SetValue(ref key, ref pv);
                            propStore.Commit();
                            Marshal.ReleaseComObject(propStore);
                        }
                    }
                }
"@
        }

        [WindowIconAppId]::SetAppUserModelId($hwnd, $AppId)

        # Also set the icon via WM_SETICON
        if (-not ([System.Management.Automation.PSTypeName]'WindowIconSetter').Type) {
            Add-Type -TypeDefinition @"
                using System;
                using System.Runtime.InteropServices;

                public class WindowIconSetter {
                    [DllImport("user32.dll", CharSet = CharSet.Auto)]
                    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

                    [DllImport("user32.dll", CharSet = CharSet.Auto)]
                    public static extern IntPtr LoadImage(IntPtr hInstance, string lpIconName, uint uType, int cxDesired, int cyDesired, uint fuLoad);

                    public const uint WM_SETICON = 0x0080;
                    public const uint ICON_SMALL = 0;
                    public const uint ICON_BIG = 1;
                    public const uint IMAGE_ICON = 1;
                    public const uint LR_LOADFROMFILE = 0x0010;

                    public static void SetIcon(IntPtr hwnd, string iconPath) {
                        IntPtr hIconSmall = LoadImage(IntPtr.Zero, iconPath, IMAGE_ICON, 16, 16, LR_LOADFROMFILE);
                        IntPtr hIconBig = LoadImage(IntPtr.Zero, iconPath, IMAGE_ICON, 32, 32, LR_LOADFROMFILE);
                        if (hIconSmall != IntPtr.Zero) SendMessage(hwnd, WM_SETICON, (IntPtr)ICON_SMALL, hIconSmall);
                        if (hIconBig != IntPtr.Zero) SendMessage(hwnd, WM_SETICON, (IntPtr)ICON_BIG, hIconBig);
                    }
                }
"@
        }

        [WindowIconSetter]::SetIcon($hwnd, $IconPath)
    }
    catch {
        Write-Verbose "Failed to set window icon: $_"
    }
}
