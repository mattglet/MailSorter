using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace MailSorter
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;
        private Microsoft.Office.Interop.Outlook.Folder _destinationFolder = null;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string returnText = string.Empty;

            switch (ribbonID)
            {
                case "Microsoft.Outlook.Explorer":
                    returnText = GetResourceText("MailSorter.Ribbon.xml");
                    break;
                default:
                    returnText = string.Empty;
                    break;
            }

            return returnText;
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void FileMessage_OnClick(Microsoft.Office.Core.IRibbonControl control)
        {
            var app = new Microsoft.Office.Interop.Outlook.Application();
            var explorer = app.ActiveExplorer();
            Microsoft.Office.Interop.Outlook.Selection selections = explorer.Selection;
            Microsoft.Office.Interop.Outlook.Folder inbox = app.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox) as Microsoft.Office.Interop.Outlook.Folder;

            foreach (var selection in selections)
            {
                var personName = string.Empty;

                if (selection is Microsoft.Office.Interop.Outlook.MailItem)
                {
                    personName = ((Microsoft.Office.Interop.Outlook.MailItem)selection).SenderName;
                    _destinationFolder = GetMatchingFolder(inbox, personName);

                    MoveMailItem((Microsoft.Office.Interop.Outlook.MailItem)selection, personName);
                }
                else if (selection is Microsoft.Office.Interop.Outlook.MeetingItem)
                {
                    personName = ((Microsoft.Office.Interop.Outlook.MeetingItem)selection).SenderName;
                    _destinationFolder = GetMatchingFolder(inbox, personName);

                    MoveMeetingItem((Microsoft.Office.Interop.Outlook.MeetingItem)selection, personName);
                }
                else
                {
                    MessageBox.Show("Unknown message type, can't move it.");
                    break;
                }

                _destinationFolder = null;
            }

            _destinationFolder = null;
        }

        #endregion

        #region Helpers
        private void MoveMailItem(Microsoft.Office.Interop.Outlook.MailItem item, string personName)
        {
            if (_destinationFolder != null)
            {
                item.UnRead = false;
                item.Move(_destinationFolder);
            }
            else
            {
                MessageBox.Show("Folder not found:" + personName);
            }
        }

        private void MoveMeetingItem(Microsoft.Office.Interop.Outlook.MeetingItem item, string personName)
        {
            if (_destinationFolder != null)
            {
                item.UnRead = false;
                item.Move(_destinationFolder);
            }
            else
            {
                MessageBox.Show("Folder not found:" + personName);
            }
        }

        private Microsoft.Office.Interop.Outlook.Folder GetMatchingFolder(Microsoft.Office.Interop.Outlook.Folder folder, string personName)
        {
            if (folder.Name != personName && _destinationFolder == null)
            {
                foreach (Microsoft.Office.Interop.Outlook.Folder subFolder in folder.Folders)
                {
                    if (subFolder.Name != personName && _destinationFolder == null)
                    {
                        _destinationFolder = GetMatchingFolder(subFolder, personName);
                    }
                    else
                    {
                        if (_destinationFolder == null)
                        {
                            _destinationFolder = subFolder;
                        }
                    }
                }
            }
            else
            {
                _destinationFolder = folder;
            }

            return _destinationFolder;
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
