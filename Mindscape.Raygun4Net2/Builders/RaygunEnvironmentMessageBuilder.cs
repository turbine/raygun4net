using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Mindscape.Raygun4Net.Messages;

namespace Mindscape.Raygun4Net.Builders
{
  public class RaygunEnvironmentMessageBuilder
  {
    public static RaygunEnvironmentMessage Build()
    {
      RaygunEnvironmentMessage message = new RaygunEnvironmentMessage();

      // Different environments can fail to load the environment details.
      // For now if they fail to load for whatever reason then just
      // swallow the exception. A good addition would be to handle
      // these cases and load them correctly depending on where its running.
      // see http://raygun.io/forums/thread/3655

      try
      {
        IntPtr hWnd = GetActiveWindow();
        RECT rect;
        GetWindowRect(hWnd, out rect);
        message.WindowBoundsWidth = rect.Right - rect.Left;
        message.WindowBoundsHeight = rect.Bottom - rect.Top;
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine(string.Format("Error retrieving window dimensions: {0}", ex.Message));
      }

      try
      {
        DateTime now = DateTime.Now;
        message.UtcOffset = TimeZone.CurrentTimeZone.GetUtcOffset(now).TotalHours;
        message.Locale = CultureInfo.CurrentCulture.DisplayName;
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine(string.Format("Error retrieving time and locale: {0}", ex.Message));
      }

      try
      {
        message.ProcessorCount = Environment.ProcessorCount;
        message.Architecture = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE");
        message.OSVersion = Environment.OSVersion.VersionString;
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine(string.Format("Error retrieving processor info: {0}", ex.Message));
      }

      try
      {
        message.DiskSpaceFree = GetDiskSpace();
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine(string.Format("Error retrieving disk info: {0}", ex.Message));
      }

      // The WMI stuff isn't present in linux mono, and these values seem silly to me anyway
      //try
      //{
      //  string Query = "SELECT TotalVisibleMemorySize, FreePhysicalMemory, MaxProcessMemorySize FROM Win32_OperatingSystem";
      //  var searcher = new System.Management.ManagementObjectSearcher(Query);

      //  message.TotalPhysicalMemory = 0;
      //  message.AvailablePhysicalMemory = 0;
      //  message.TotalVirtualMemory = 0;
      //  foreach (ManagementObject WniPART in searcher.Get())
      //  {
      //    message.TotalPhysicalMemory += Convert.ToUInt64(WniPART.Properties["TotalVisibleMemorySize"].Value);
      //    message.AvailablePhysicalMemory += Convert.ToUInt64(WniPART.Properties["FreePhysicalMemory"].Value);
      //    message.TotalVirtualMemory += Convert.ToUInt64(WniPART.Properties["MaxProcessMemorySize"].Value);

      //    // I can't find a query that returns this value for 32-bit apps. Since we don't know if /3GB is  
      //    // on or not, so assume 2G for a 32bit exe on 32bit os and 4G for 32bit exe on 64bit OS
      //    if (IntPtr.Size == 4)
      //    {
      //      if (message.TotalVirtualMemory > 4096 * 1024)
      //        message.TotalVirtualMemory = 4096 * 1024;
      //      else if (message.TotalVirtualMemory > 2047 * 1024)
      //        message.TotalVirtualMemory = 2047 * 1024;
      //    }
      //  }
      //}
      //catch (Exception ex)
      //{
      //  System.Diagnostics.Debug.WriteLine(string.Format("Error retrieving os memory info: {0}", ex.Message));
      //}

      //if (message.TotalVirtualMemory > 0)
      //{
      //  try
      //  {
      //    var process = System.Diagnostics.Process.GetCurrentProcess();
      //    message.AvailableVirtualMemory = message.TotalVirtualMemory - (ulong)(process.VirtualMemorySize64 / 1024);
      //  }
      //  catch (Exception ex)
      //  {
      //    System.Diagnostics.Debug.WriteLine(string.Format("Error retrieving os memory info: {0}", ex.Message));
      //  }
      //}

      // In MB
      //message.TotalPhysicalMemory /= 1024;
      //message.AvailablePhysicalMemory /= 1024;
      //message.TotalVirtualMemory /= 1024;
      //message.AvailableVirtualMemory /= 1024;

      try
      {
        var process = System.Diagnostics.Process.GetCurrentProcess();
        message.TotalVirtualMemory = (ulong)(process.VirtualMemorySize64/(1024*1024));
        message.TotalPhysicalMemory = (ulong)(process.PrivateMemorySize64/(1024*1024));
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine(string.Format("Error retrieving process memory info: {0}", ex.Message));
      }

      return message;
    }

    private static List<double> GetDiskSpace()
    {
      List<double> diskSpaceFree = new List<double>();
      foreach (DriveInfo drive in DriveInfo.GetDrives())
      {
        if (drive.IsReady)
        {
          diskSpaceFree.Add((double)drive.AvailableFreeSpace / 0x40000000); // in GB
        }
      }
      return diskSpaceFree;
    }

    [DllImport("user32.dll")]
    static extern IntPtr GetForegroundWindow();

    private static IntPtr GetActiveWindow()
    {
      IntPtr handle = IntPtr.Zero;
      return GetForegroundWindow();
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
      public int Left;
      public int Top;
      public int Right;
      public int Bottom;
    }
  }
}
