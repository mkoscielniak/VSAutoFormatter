using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio;
using VSInterop = Microsoft.VisualStudio.OLE.Interop;

namespace VSAutoFormatter
{
    public class DocumentListener : IDisposable, IVsRunningDocTableEvents3
    {
        private uint? _cookie;
        private static volatile object _mutex;


        public DocumentListener(IServiceProvider serviceProvider)
        {
            _mutex = new object();
            RunningDocumentTable = new RunningDocumentTable(serviceProvider);
        }


        public RunningDocumentTable RunningDocumentTable { get; protected set; }


        #region Events


        public event OnAfterAttributeChangeEventHandler AfterAttributeChange;
        public event OnAfterDocumentWindowHideEventHandler AfterDocumentWindowHide;
        public event OnAfterFirstDocumentLockEventHandler AfterFirstDocumentLock;
        public event OnAfterSaveEventHandler AfterSave;
        public event OnBeforeDocumentWindowShowEventHandler BeforeDocumentWindowShow;
        public event OnBeforeLastDocumentUnlockEventHandler BeforeLastDocumentUnlock;
        public event OnAfterAttributeChangeExHandler AfterAttributeChangeEx;
        public event OnBeforeSaveHandler BeforeSave;


        #endregion


        public void Start()
        {
            _cookie = RunningDocumentTable.Advise(this);
        }


        #region IDisposable Members


        public void Dispose()
        {
            lock (_mutex)
            {
                if (RunningDocumentTable == null)
                    return;

                if (!_cookie.HasValue)
                    return;

                RunningDocumentTable.Unadvise(_cookie.Value);
                _cookie = null;
                RunningDocumentTable = null;
            }
        }


        #endregion


        #region IVsRunningDocTableEvents3 Members


        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
        {
            if (AfterAttributeChange != null)
                return AfterAttributeChange(docCookie, grfAttribs);
            return VSConstants.S_OK;
        }


        public int OnAfterAttributeChangeEx(uint docCookie, uint grfAttribs, IVsHierarchy pHierOld, uint itemidOld, string pszMkDocumentOld, IVsHierarchy pHierNew, uint itemidNew, string pszMkDocumentNew)
        {
            if (AfterAttributeChangeEx != null)
                return AfterAttributeChangeEx(docCookie, grfAttribs, pHierOld, itemidOld, pszMkDocumentOld, pHierNew, itemidNew, pszMkDocumentNew);
            return VSConstants.S_OK;
        }


        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame)
        {
            if (AfterDocumentWindowHide != null)
                return AfterDocumentWindowHide(docCookie, pFrame);
            return VSConstants.S_OK;
        }


        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            if (AfterFirstDocumentLock != null)
                return AfterFirstDocumentLock(docCookie, dwRDTLockType, dwReadLocksRemaining, dwEditLocksRemaining);
            return VSConstants.S_OK;
        }


        public int OnAfterSave(uint docCookie)
        {
            if (AfterSave != null)
                return AfterSave(docCookie);
            return VSConstants.S_OK;
        }


        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
        {
            if (BeforeDocumentWindowShow != null)
                return BeforeDocumentWindowShow(docCookie, fFirstShow, pFrame);
            return VSConstants.S_OK;
        }


        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            if (BeforeLastDocumentUnlock != null)
                return BeforeLastDocumentUnlock(docCookie, dwRDTLockType, dwReadLocksRemaining, dwEditLocksRemaining);
            return VSConstants.S_OK;
        }


        public int OnBeforeSave(uint docCookie)
        {
            if (BeforeSave != null)
                return BeforeSave(docCookie);
            return VSConstants.S_OK;
        }


        #endregion


        #region Delegates

        public delegate int OnAfterAttributeChangeEventHandler(uint docCookie, uint grfAttribs);

        public delegate int OnAfterDocumentWindowHideEventHandler(uint docCookie, IVsWindowFrame pFrame);

        public delegate int OnAfterFirstDocumentLockEventHandler(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining);

        public delegate int OnAfterSaveEventHandler(uint docCookie);

        public delegate int OnBeforeDocumentWindowShowEventHandler(uint docCookie, int fFirstShow, IVsWindowFrame pFrame);

        public delegate int OnBeforeLastDocumentUnlockEventHandler(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining);

        public delegate int OnAfterAttributeChangeExHandler(uint docCookie, uint grfAttribs, IVsHierarchy pHierOld, uint itemidOld, string pszMkDocumentOld, IVsHierarchy pHierNew, uint itemidNew, string pszMkDocumentNew);

        public delegate int OnBeforeSaveHandler(uint docCookie);

        #endregion
    }
}
