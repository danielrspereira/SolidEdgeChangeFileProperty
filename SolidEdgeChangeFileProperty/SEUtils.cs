using System;
using System.Runtime.InteropServices;

namespace SolidEdgeChangeFileProperty
{
    class SEUtils
    {
        //[DllImport("ole32.dll")]
        //static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        //[DllImport("ole32.dll")]
        //static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        //const int MK_E_UNAVAILABLE = (int)(0x800401E3 - 0x100000000);
        const int MK_E_UNAVAILABLE = unchecked((int)0x800401E3);


        //Connect Method 1: Connects to a running instance of Solid Edge and return an object of type SolidEdgeFramework.Application
        //Used if no parameter is inserted and calls Connect Method 2
        public static SolidEdgeFramework.Application Connect()
        {
            return Connect(startIfNotRunning: false);
        }

        //Connect Method 2: Connects to or starts a new instance of Solid Edge. Parameter "true", return object of type SolidEdgeFramework.Application
        public static SolidEdgeFramework.Application Connect(bool startIfNotRunning)
        {
            try
            {
                // Attempt to connect to a running instance of Solid Edge.
                //return (SolidEdgeFramework.Application)Marshal.GetActiveObject(progID: SolidEdgeSDK.PROGID.SolidEdge_Application);
                return (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                switch (ex.ErrorCode)
                {
                    // Solid Edge is not running.
                    case MK_E_UNAVAILABLE:
                        if (startIfNotRunning)
                        {
                            // Start Solid Edge.
                            return Start();
                        }
                        else
                        {
                            // Rethrow exception.
                            throw;
                        }
                    default:
                        // Rethrow exception.
                        throw;
                }
            }
            catch
            {
                // Rethrow exception.
                throw;
            }
        }

        // Connects to or starts a new instance of Solid Edge. Parameters startFfNotRunning and ensureVisible. Returns an object of type SolidEdgeFramework.Application
        public static SolidEdgeFramework.Application Connect(bool startIfNotRunning, bool ensureVisible)
        {
            SolidEdgeFramework.Application application = null;

            try
            {
                // Attempt to connect to a running instance of Solid Edge.
                //application = (SolidEdgeFramework.Application)Marshal.GetActiveObject(progID: SolidEdgeSDK.PROGID.SolidEdge_Application);
                application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                switch (ex.ErrorCode)
                {
                    // Solid Edge is not running.
                    case MK_E_UNAVAILABLE:
                        if (startIfNotRunning)
                        {
                            // Start Solid Edge.
                            application = Start();
                            break;
                        }
                        else
                        {
                            // Rethrow exception.
                            throw;
                        }
                    default:
                        // Rethrow exception.
                        throw;
                }
            }
            catch
            {
                // Rethrow exception.
                throw;
            }

            if ((application != null) && (ensureVisible))
            {
                application.Visible = true;
            }

            return application;
        }

        public static SolidEdgeFramework.SolidEdgeDocument ConnectToActiveDocument(SolidEdgeFramework.Application app)
        {
            return (SolidEdgeFramework.SolidEdgeDocument)app.ActiveDocument;
        }


        //Creates and returns a new instance of Solid Edge.
        public static SolidEdgeFramework.Application Start()
        {
            // On a system where Solid Edge is installed, the COM ProgID will be
            // defined in registry: HKEY_CLASSES_ROOT\SolidEdge.Application
            Type t = Type.GetTypeFromProgID("SolidEdge.Application", throwOnError: true);

            // Using the discovered Type, create and return a new instance of Solid Edge.
            return (SolidEdgeFramework.Application)Activator.CreateInstance(type: t);
        }

        //Return the document type
        // Get a reference to the active document.
        public static string GetActiveDocType()
        {
            SolidEdgeFramework.Application application = Connect();
            SolidEdgeFramework.SolidEdgeDocument document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
            string docType = null;

            // Using Type property, determine document type.
            switch (document.Type)
            {
                case SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument:
                    docType = "Assembly Document";
                    break;
                case SolidEdgeFramework.DocumentTypeConstants.igDraftDocument:
                    docType = "Draft Document";
                    break;
                case SolidEdgeFramework.DocumentTypeConstants.igPartDocument:
                    docType = "Part Document";
                    break;
                case SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument:
                    docType = "SheetMetal Document";
                    break;
                case SolidEdgeFramework.DocumentTypeConstants.igUnknownDocument:
                    docType = "Unknown Document";
                    break;
                case SolidEdgeFramework.DocumentTypeConstants.igWeldmentAssemblyDocument:
                    docType = "Weldment Assembly Document";
                    break;
                case SolidEdgeFramework.DocumentTypeConstants.igWeldmentDocument:
                    docType = "Weldment Document";
                    break;
            }
            return docType;

        }

        public static void OpenSolidEdgeFile()
        {
            SolidEdgeFramework.Application application = Connect(true, true);
            SolidEdgeFramework.Documents documents = application.Documents;
            documents.OpenWithFileOpenDialog();
        }

        public static void OpenSolidEdgeFile(string file)
        {
            SolidEdgeFramework.Application application = Connect(true, true);
            SolidEdgeFramework.Documents documents = application.Documents;
            documents.Open(file);
        }

        public static void SaveActiveDocument()
        {
            SolidEdgeFramework.Application application = Connect();
            SolidEdgeFramework.SolidEdgeDocument document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
            document.Save();
        }

        public static void SaveAllOpenDocuments()
        {
            SolidEdgeFramework.Application application = Connect();
            SolidEdgeFramework.Documents documents = application.Documents;
            foreach (SolidEdgeFramework.SolidEdgeDocument document in documents)
            {
                document.Save();
            }
        }

        public static void CloseSolidEdgeFile(string file, bool saveChanges)
        {
            SolidEdgeFramework.Application application = Connect();
            SolidEdgeFramework.Documents documents = application.Documents;
            documents.CloseDocument(file, saveChanges);
        }

        public static void DoIdle()
        {
            SolidEdgeFramework.Application application = Connect();
            application.DoIdle();
        }
    }
}
