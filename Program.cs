using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using SolidEdgeFramework;

namespace SENomexLayers
{
    public class MarshalHelper
    {
        public static object? GetActiveObject(string progId, bool throwOnError = false)
        {
            if (progId == null)
            {
                throw new ArgumentNullException(nameof(progId));
            }

            Guid clsid;
            int hr = CLSIDFromProgIDEx(progId, out clsid);

            if (hr < 0)
            {
                if (throwOnError)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }
                return null;
            }

            object? obj = null;
            hr = GetActiveObject(ref clsid, IntPtr.Zero, out obj);

            if (hr < 0)
            {
                if (throwOnError)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }
                return null;
            }

            return obj;
        }

        [DllImport("ole32")]
        private static extern int CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid lpclsid);

        [DllImport("oleaut32")]
        private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
    }

    public class OleMessageFilter : IOleMessageFilter
    {
        public static void Register()
        {
            IOleMessageFilter newFilter = new OleMessageFilter();
            IOleMessageFilter? oldFilter = null;
            int iRetVal;

            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            {
                iRetVal = CoRegisterMessageFilter(newFilter, out oldFilter);
            }
            else
            {
                throw new COMException("Unable to register message filter because the current thread apartment state is not STA.");
            }
        }

        public static void Revoke()
        {
            IOleMessageFilter? oldFilter = null;
            CoRegisterMessageFilter(null, out oldFilter);
        }

        public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return (int)SERVERCALL.SERVERCALL_ISHANDLED;
        }

        public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            if (dwRejectType == (int)SERVERCALL.SERVERCALL_RETRYLATER)
            {
                return 99;
            }

            return -1;
        }

        public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            return (int)PENDINGMSG.PENDINGMSG_WAITDEFPROCESS;
        }

        [DllImport("Ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
    }

    public enum SERVERCALL
    {
        SERVERCALL_ISHANDLED = 0,
        SERVERCALL_REJECTED = 1,
        SERVERCALL_RETRYLATER = 2
    }

    public enum PENDINGMSG
    {
        PENDINGMSG_CANCELCALL = 0,
        PENDINGMSG_WAITNOPROCESS = 1,
        PENDINGMSG_WAITDEFPROCESS = 2
    }

    [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

        [PreserveSig]
        int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

    public class SolidEdgeConnector
    {
        public SolidEdgeDocument ConnectToSolidEdge()
        {
            Application? objApplication = null;
            SolidEdgeDocument? objDocument = null;

            try
            {
                OleMessageFilter.Register();

                // Connect to a running instance of Solid Edge
                objApplication = (Application?)MarshalHelper.GetActiveObject("SolidEdge.Application");

                // Check if objApplication is not null before accessing ActiveDocument
                if (objApplication != null)
                {
                    // Get the active document object
                    objDocument = objApplication.ActiveDocument;
                }
                else
                {
                    throw new InvalidOperationException("Failed to get the SolidEdge.Application object.");
                }
            }
            finally
            {
                OleMessageFilter.Revoke();
            }

            return objDocument ?? throw new InvalidOperationException("Failed to get the active document.");
        }

        public List<string> GetCustomProperties(SolidEdgeDocument document)
        {
            var customPropertiesList = new List<string>();
            try
            {
                var objPropSets = document.Properties;
                foreach (Properties objProps in objPropSets)
                {
                    if (objProps.Name == "Custom")
                    {
                        foreach (Property objProp in objProps)
                        {
                            customPropertiesList.Add($"{objProp.Name}: {objProp.get_Value()}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to retrieve custom properties: {ex.Message}");
            }

            return customPropertiesList;
        }

        public List<string> GetNomexLayers(List<string> uniqueCodes, string searchFolder)
        {
            var nomexLayers = new List<string>();

            foreach (var uniqueCode in uniqueCodes)
            {
                var filePath = Path.Combine(searchFolder, $"{uniqueCode}.psm");
                if (File.Exists(filePath))
                {
                    SolidEdgeDocument? objDocument = null;
                    try
                    {
                        var objApplication = (Application?)MarshalHelper.GetActiveObject("SolidEdge.Application");
                        if (objApplication != null)
                        {
                            objDocument = (SolidEdgeDocument?)objApplication.Documents.Open(filePath);
                            if (objDocument != null)
                            {
                                var customProperties = GetCustomProperties(objDocument);
                                foreach (var prop in customProperties)
                                {
                                    if (prop.StartsWith("NOMEX_LAYERS:"))
                                    {
                                        nomexLayers.Add(prop);
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine($"Failed to open document: {filePath}");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Failed to get the SolidEdge.Application object.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to retrieve SINK CONFIGURATION for {uniqueCode}: {ex.Message}");
                    }
                    finally
                    {
                        objDocument?.Close(false);
                    }
                }
                else
                {
                    Console.WriteLine($"File not found: {filePath}");
                }
            }

            return nomexLayers;
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        try
        {
            if (args.Length < 2)
            {
                throw new ArgumentException("Insufficient arguments provided.");
            }

            var uniqueCode = args[0];
            var searchFolder = args[1];

            var solidEdgeConnector = new SENomexLayers.SolidEdgeConnector();
            var uniqueCodes = new List<string> { uniqueCode };

            var sinkConfigurationValues = solidEdgeConnector.GetNomexLayers(uniqueCodes, searchFolder);
            foreach (var value in sinkConfigurationValues)
            {
                Console.WriteLine(value);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
