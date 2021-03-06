using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using VSInterop = Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using VSLangProj;


namespace VSAutoFormatter
{
    public class Addin : IDTExtensibility2
    {
        private DTE2 _applicationObject;
        private AddIn _addInInstance;
        private DocumentListener _listener;


        public const string FORMAT_DOCUMENT_CMD_NAME = "Edit.FormatDocument";


        public Addin()
        {
        }


        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _applicationObject = (DTE2)application;
            _addInInstance = (AddIn)addInInst;

            IServiceProvider serviceProvider = new ServiceProvider(_applicationObject as VSInterop.IServiceProvider);

            _listener = new DocumentListener(serviceProvider);
            _listener.BeforeSave += DocumentListener_BeforeSave;
            _listener.Start();
        }


        private int DocumentListener_BeforeSave(uint docCookie)
        {
            Document activeDocument = _applicationObject.ActiveDocument;
            Document document = FindDocument(docCookie);

            if (!CanInvokeFormat(document))
                return VSConstants.S_OK;

            document.Activate();
            _applicationObject.ExecuteCommand(FORMAT_DOCUMENT_CMD_NAME, String.Empty);
            activeDocument.Activate();

            return VSConstants.S_OK;
        }

        private bool CanInvokeFormat(Document document)
        {
            if (document == null)
                return false;

            if (document.ProjectItem.FileCodeModel == null)
                return false;

            string projectKind = document.ProjectItem.ContainingProject.Kind;
            
            if (projectKind != PrjKind.prjKindCSharpProject)
                return false;

            if (!document.Name.EndsWith(".cs"))
                return false;

            return true;
        }


        private Document FindDocument(uint docCookie)
        {
            Document result = null;
            RunningDocumentInfo rdi = _listener.RunningDocumentTable.GetDocumentInfo(docCookie);
            string documentPath = rdi.Moniker;

            foreach (Document doc in _applicationObject.Documents)
            {
                if (doc.FullName == documentPath)
                {
                    result = doc;
                    break;
                }
            }

            return result;
        }


        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            if (_listener == null)
                return;

            _listener.BeforeSave -= DocumentListener_BeforeSave;
            _listener.Dispose();
            _listener = null;
        }


        public void OnAddInsUpdate(ref Array custom)
        {
        }


        public void OnStartupComplete(ref Array custom)
        {
        }


        public void OnBeginShutdown(ref Array custom)
        {
        }
    }
}