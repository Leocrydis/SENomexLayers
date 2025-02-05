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
        public SolidEdgePart.SheetMetalDocument ConnectToSolidEdge(string filePath)
        {
            Application? objApplication = null;
            SolidEdgePart.SheetMetalDocument? objDocument = null;

            try
            {
                OleMessageFilter.Register();

                // Try to connect to a running instance of Solid Edge
                objApplication = (Application?)MarshalHelper.GetActiveObject("SolidEdge.Application");

                if (objApplication != null)
                {
                    // Ensure Solid Edge is in visible mode and display alerts are disabled
                    objApplication.Visible = true;
                    objApplication.DisplayAlerts = false;

                    // Open the document
                    objDocument = (SolidEdgePart.SheetMetalDocument)objApplication.Documents.Open(filePath);
                }
                else
                {
                    // If no running instance, create a new instance of Solid Edge
                    objApplication = (Application)Activator.CreateInstance(Type.GetTypeFromProgID("SolidEdge.Application")!);

                    if (objApplication != null)
                    {
                        // Ensure Solid Edge is in non-visible mode and display alerts are disabled
                        objApplication.Visible = false;
                        objApplication.DisplayAlerts = false;

                        // Open the document
                        objDocument = (SolidEdgePart.SheetMetalDocument)objApplication.Documents.Open(filePath);
                    }
                    else
                    {
                        throw new InvalidOperationException("Failed to create a new instance of SolidEdge.Application.");
                    }
                }
            }
            finally
            {
                OleMessageFilter.Revoke();
            }

            return objDocument ?? throw new InvalidOperationException("Failed to open the document.");
        }

        //First try to open the custom section of the file properties without opening the document
        public List<string> GetCustomProperties(string filePath)
        {
            var customPropertiesList = new List<string>();
            try
            {
                var propertySets = new SolidEdgeFileProperties.PropertySets();
                propertySets.Open(filePath, true);

                foreach (SolidEdgeFileProperties.Properties properties in propertySets)
                {
                    if (properties.Name == "Custom")
                    {
                        foreach (SolidEdgeFileProperties.Property property in properties)
                        {
                            customPropertiesList.Add($"{property.Name}: {property.Value}");
                        }
                    }
                }
            }
            // If the the custom sections happens to not be there its likely because the document is open in Solid Edge
            catch (Exception ex)
            {
                //Most likely means someone has it open in Solid Edge?
                Console.WriteLine($"Failed to retrieve file properties without opening the document: {ex.Message}"); 

                // If failed, open the document in Solid Edge using SolidEdgeConnector Function
                SolidEdgePart.SheetMetalDocument? objDocument = null;
                try
                {
                    objDocument = ConnectToSolidEdge(filePath);
                    if (objDocument != null)
                    {
                        var objPropSets = objDocument.Properties;
                        foreach (Properties objProps in objPropSets)
                        {
                            if (objProps.Name == "Custom")
                            {
                                for (int i = 1; i <= objProps.Count; i++)
                                {
                                    var objProp = objProps.Item(i) as PropertyEx;
                                    if (objProp != null)
                                    {
                                        var value = ((dynamic)objProp).Value;
                                        customPropertiesList.Add($"{objProp.Name}: {value}");
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Failed to open document: {filePath}");
                    }
                }
                catch (Exception innerEx)
                {
                    Console.WriteLine($"Failed to retrieve custom properties by opening the document: {innerEx.Message}");
                }
                finally
                {
                    objDocument?.Close(false);
                }
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
                    try
                    {
                        // Attempt to get properties
                        var customProperties = GetCustomProperties(filePath);

                        if (customProperties != null)
                        {
                            foreach (var prop in customProperties)
                            {
                                if (prop.StartsWith("NOMEX_LAYERS"))
                                {
                                    nomexLayers.Add($"{uniqueCode}: {prop}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to retrieve NOMEX_LAYERS for {uniqueCode}: {ex.Message}");
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
    [STAThread]
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Insufficient arguments provided. Please provide the search folder followed by a comma-separated list of unique codes.");
            return;
        }

        var searchFolder = args[0];
        var uniqueCodes = new List<string>(args[1].Split(','));

        var thread = new Thread(() =>
        {
            try
            {
                var solidEdgeConnector = new SENomexLayers.SolidEdgeConnector();
                var nomexLayerValues = solidEdgeConnector.GetNomexLayers(uniqueCodes, searchFolder);
                foreach (var value in nomexLayerValues)
                {
                    Console.WriteLine(value);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
    }
}