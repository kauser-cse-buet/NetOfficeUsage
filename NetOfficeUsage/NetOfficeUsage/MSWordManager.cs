using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using Office = NetOffice.OfficeApi;
using NetOffice.WordApi;
using System.Globalization;

namespace NetOfficeUsage
{
    class MSWordManager
    {
        public void CreateNewDoc(string docName)
        {
            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // add a new document
            Word.Document newDocument = wordApplication.Documents.Add();

            // insert some text
            wordApplication.Selection.TypeText("This text is written by NetOffice");

            wordApplication.Selection.HomeKey(WdUnits.wdLine, WdMovementType.wdExtend);
            wordApplication.Selection.Font.Color = WdColor.wdColorSeaGreen;
            wordApplication.Selection.Font.Bold = 1;
            wordApplication.Selection.Font.Size = 18;

            string applicationPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            // save the document
            string fileExtension = GetDefaultExtension(wordApplication);
            object documentFile =
                   string.Format("{0}\\" + docName + "{1}", applicationPath, fileExtension);
            newDocument.SaveAs(documentFile);

            // close word and dispose reference
            wordApplication.Quit();
            wordApplication.Dispose();

            Console.WriteLine("Document saved.");

        }

        #region Helper

        /// <summary>
        /// returns the valid file extension for the instance. for example ".doc" or ".docx"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(Word.Application application)
        {
            double version = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture);
            if (version >= 12.00)
                return ".docx";
            else
                return ".doc";
        }

        #endregion
    }
}
