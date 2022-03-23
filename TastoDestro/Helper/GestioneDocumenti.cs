using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell.Interop;
    using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.ComponentModelHost;

namespace TastoDestro.Helper
{
    class GestioneDocumenti: IVsRunningDocTableEvents3
    {
        #region Members

        private RunningDocumentTable mRunningDocumentTable;
        private IVsEditorAdaptersFactoryService editorAdaptersFactory;

        private DTE mDte;

        public delegate void OnBeforeSaveHandler(object sender, Document document);
        public event OnBeforeSaveHandler BeforeSave;

        #endregion

        #region Constructor

        public  GestioneDocumenti(Package aPackage)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            mDte = (DTE)Package.GetGlobalService(typeof(DTE));
            IComponentModel componentModel = Package.GetGlobalService(typeof(SComponentModel)) as IComponentModel;
            editorAdaptersFactory = componentModel.GetService<IVsEditorAdaptersFactoryService>();
            mRunningDocumentTable = new RunningDocumentTable(aPackage);
            mRunningDocumentTable.Advise(this);
        }

        #endregion

        #region IVsRunningDocTableEvents3 implementation

        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterAttributeChangeEx(uint docCookie, uint grfAttribs, IVsHierarchy pHierOld, uint itemidOld, string pszMkDocumentOld, IVsHierarchy pHierNew, uint itemidNew, string pszMkDocumentNew)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame)
        {
   
            return VSConstants.S_OK;

        }

        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            
            return VSConstants.S_OK;
        }

        public int OnAfterSave(uint docCookie)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            if (((dwRDTLockType == 2)||(dwRDTLockType==258)) && (dwReadLocksRemaining == 2) && (dwEditLocksRemaining == 0))
            {
                if (null == BeforeSave)
                    return VSConstants.S_OK;

                var document = FindDocumentByCookie(docCookie);
                if (null == document)
                    return VSConstants.S_OK;

                //BeforeSave(this, FindDocumentByCookie(docCookie));
                BeforeSave(this, document);
                return VSConstants.S_OK;
            }
            return VSConstants.S_OK;
        }
    

        public int OnBeforeSave(uint docCookie)
        {
        
            return VSConstants.S_OK;
        }




        #endregion

        #region Private Methods

        private Document FindDocumentByCookie(uint docCookie)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var documentInfo = mRunningDocumentTable.GetDocumentInfo(docCookie);
            var textLines = documentInfo.DocData as IVsTextLines;

            var textBuffer=editorAdaptersFactory.GetDataBuffer(textLines);
            if (textBuffer.CurrentSnapshot.Length == 0) { return null; }
            else
            {
                //textBufferProvider.GetTextBuffer(out ciao);

                //if (documentInfo.=null) {
                return mDte.Documents.Cast<Document>().FirstOrDefault(doc => doc.FullName == documentInfo.Moniker);
                //}
                //else
                //{
                //    return null;
                //}
            }
        }

        #endregion
    }
}
