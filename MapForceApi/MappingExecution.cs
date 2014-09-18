//#define DEBUG_TRANSFORMATION

using System;
using System.Management.Instrumentation;
using System.Runtime.InteropServices;
using System.Threading;
using MapForceLib;
using Component = MapForceLib.Component;

namespace MapForceApi
{
    public class MappingExecution : IDisposable
    {
        [DllImport("User32.dll")]
        private static extern IntPtr GetLastActivePopup(IntPtr hWnd);
        [DllImport("User32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, IntPtr lParam);
        [DllImport("kernel32.dll")]
        // ReSharper disable once UnusedMember.Local
        private static extern Int32 GetLastError();
        const int WmSyscommand = 0x0112;
        const int ScClose = 0xF060;

        private readonly Application mapforce;

        public MappingExecution()
        {
            // Open MapForce or access currently running instance
            // Make MapForce UI visible. This is an API requirement for output generation.
            mapforce = new Application { Visible = true };
        }

        // ---- MAIN -------------------------------------
        public void ExecuteMap(string mapForceDataMappingFile, string inputFile, string outputFile)
        {
            try
            {
                // Open an existing mapping.
                Document doc = null;

#if DEBUG_TRANSFORMATION
                doc = mapforce.OpenDocument(mapForceDataMappingFile);
#else
                ThreadStart threadStart = () =>
                {
                    try
                    {
                        doc = mapforce.OpenDocument(mapForceDataMappingFile);
                    }
                    catch (Exception ex)
                    {
                        Error(ex.Message, ex);
                    }
                };

                RunCommandAndShutDownPopups(threadStart);
#endif
                if (null == doc) return;

                // Find existing components by name in the main mapping.
                // The names of components may not be unique as a schema component's name is derived from its schema file name.
                //var sourceComponent = FindComponent(doc.MainMapping, "Employees");
                //var targetComponent = FindComponent(doc.MainMapping, "PersonList");

                // If you do not know the names of the components for some reason, you could
                // use the following functions instead of FindComponent.
                var sourceComponent = GetFirstSourceComponent(doc.MainMapping);
                var targetComponent = GetFirstTargetComponent(doc.MainMapping);

                // Specify the desired input and output files.
                if (!string.IsNullOrEmpty(inputFile))
                {
                    sourceComponent.InputInstanceFile = inputFile;
                }

                if (!string.IsNullOrEmpty(outputFile))
                {
                    targetComponent.OutputInstanceFile = outputFile;
                }

                // Perform the transformation.
                // You can use doc.GenerateOutput() if you do not need result messages.
                // If you have a mapping with more than one target component and you want
                // to execute the transformation only for one specific target component,
                // call targetComponent.GenerateOutput() instead.
                AppOutputLines resultMessages = null;
#if DEBUG_TRANSFORMATION
                resultMessages = doc.GenerateOutputEx();
#else
                threadStart = () => { resultMessages = doc.GenerateOutputEx(); };
                RunCommandAndShutDownPopups(threadStart);
#endif
                string summaryInfo = string.Format("Transformation result: {0}", GetResultMessagesString(resultMessages));

                Console.WriteLine(summaryInfo);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Error("Failure", ex);
            }
            catch (Exception err)
            {
                Error("Failure", err);
            }

        }

        private void RunCommandAndShutDownPopups(ThreadStart threadStart)
        {
            var newThread = new Thread(threadStart);
            newThread.Start();

            // Let the mapping run for a second before we start looking for popups
            Thread.Sleep(1000);

            // Wait for the window to show a success or failure popup and shut it down
            while (newThread.IsAlive)
            {
                IntPtr popupHandle = GetLastActivePopup((IntPtr) mapforce.WindowHandle);

                // Do not close the main window
                if (popupHandle != (IntPtr) mapforce.WindowHandle)
                {
                    SendMessage(popupHandle, WmSyscommand, ScClose, IntPtr.Zero);
                }
            }

            // Wait for the thread to terminate 
            newThread.Join();
        }

        // ---- general helpers ------------------------------
        private void Error(string message, Exception err)
        {
            if (err != null)
            {
                Console.WriteLine("ERROR: ({0}) {1} - {2}", err.HResult & 0xffff, err, message);
            }
            else
            {
                Console.WriteLine("ERROR: " + message);
            }
        }

        // ---- MapForce helpers -----------------------

        // Searches in the specified mapping for a component by name and returns it.
        // If not found, throws an error.
        // ReSharper disable once UnusedMember.Local
        private Component FindComponent(Mapping mapping, string componentName)
        {
            foreach (Component component in mapping.Components)
            {
                if (component.Name == componentName)
                {
                    return component;
                }
            }

            throw new InstanceNotFoundException("Cannot find component with name " + componentName);
        }

        // Browses components in a mapping and returns the first one found acting as
        // source component (i.e. having connections on its right side).
        private Component GetFirstSourceComponent(Mapping mapping)
        {
            foreach (Component component in mapping.Components)
            {
                if (component.UsageKind == ENUMComponentUsageKind.eComponentUsageKind_Instance &&
                    component.HasOutgoingConnections)
                {
                    return component;
                }
            }

            throw new InstanceNotFoundException("Cannot find a source component");
        }

        // Browses components in a mapping and returns the first one found acting as
        // target component (i.e. having connections on its left side).
        private Component GetFirstTargetComponent(Mapping mapping)
        {
            foreach (Component component in mapping.Components)
            {
                if (component.UsageKind == ENUMComponentUsageKind.eComponentUsageKind_Instance &&
                    component.HasIncomingConnections)
                {
                    return component;
                }
            }

            throw new InstanceNotFoundException("Cannot find a target component");
        }

        private string IndentTextLines(string s)
        {
            return "\t" + s.Replace(@"\n", @"\n\t");
        }

        private string GetAppoutputLineFullText(AppOutputLine oAppoutputLine)
        {
            string s = oAppoutputLine.GetLineText();

            foreach (AppOutputLine oAppoutputChildLine in oAppoutputLine.ChildLines)
            {
                string sChilds = GetAppoutputLineFullText(oAppoutputChildLine);
                s += "\n" + IndentTextLines(sChilds);
            }

            return s;
        }

        // Create a nicely formatted string from AppOutputLines
        private string GetResultMessagesString(AppOutputLines oAppoutputLines)
        {
            var s1 = "Transformation result messages:\n";
            foreach (AppOutputLine oAppoutputLine in oAppoutputLines)
            {
                s1 += GetAppoutputLineFullText(oAppoutputLine);
                s1 += "\n";
            }

            return s1;
        }

        public void Dispose()
        {
            SendMessage((IntPtr)mapforce.WindowHandle, WmSyscommand, ScClose, IntPtr.Zero);
        }
    }
}