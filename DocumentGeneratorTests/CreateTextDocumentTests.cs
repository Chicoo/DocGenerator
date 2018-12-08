using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DocumentGenerator.WordDocuments;
using System.IO;

namespace DocumentGeneratorTests
{
    [TestClass]
    public class CreateTextDocumentTests
    {
        #region Constants
        private const string TEMPLATE_NAME = "Default";
        private const string TEMPLATE_PATH = "Resources";
        private const string TEMPLATE_EXTENSION = "dotx";
        private const string DOCUMENT_NAME = "Document";
        private const string DOCUMENT_PATH = "Resources";
        private const string DOCUMENT_EXTENSION = "dotx";
        #endregion Constants

        #region Fields
        private string _filename;
        #endregion Fields

        #region Initialize
        [TestInitialize]
        public void TestStartUp()
        {
            //Clean up the resources directory in the bin directory.
            if (!Directory.Exists(TEMPLATE_PATH)) Directory.CreateDirectory(TEMPLATE_PATH);
            foreach (var file in Directory.GetFiles(TEMPLATE_PATH))
            {
                File.Delete(file);
            }

            //Write the document and template file to the resource directory
            var templateLocation = string.Format("{0}\\{1}.{2}", TEMPLATE_PATH, TEMPLATE_NAME, TEMPLATE_EXTENSION);
            _filename = string.Format("{0}\\{1}.{2}", DOCUMENT_PATH, DOCUMENT_NAME, DOCUMENT_EXTENSION);
            File.WriteAllBytes(templateLocation, Properties.Resources.Default);
            File.WriteAllBytes(_filename, Properties.Resources.Document);
        } 
        #endregion Initialize

        #region Create new documents.
        [TestMethod]
        public void CreateEmptyDocument()
        {
            TextDocument doc = TextDocument.Create();
            Assert.IsNotNull(doc, "Text Document not created");
            if (doc != null)
            {
                doc.Filename = "test.docx";
                Assert.AreEqual(doc.Filename, "test.docx", "Filename not set correct");
            }
        }

        [TestMethod]
        public void CreateNamedDocument()
        {
            TextDocument doc = TextDocument.Create("test.docx", true);
            Assert.IsNotNull(doc, "Text Document not created");
            if (doc != null)
            {
                Assert.AreEqual(doc.Filename, "test.docx", "Filename not set correct");
            }
        }

        [TestMethod]
        public void CreateDocumentFromTemplate()
        {
            var templateLocation = string.Format("{0}\\{1}.{2}", TEMPLATE_PATH, TEMPLATE_NAME, TEMPLATE_EXTENSION);
            TextDocument doc = TextDocument.Create("test.docx", templateLocation, true);
            Assert.IsNotNull(doc, "Text Document not created");
            if (doc != null)
            {
                Assert.AreEqual(doc.Filename, "test.docx", "Filename not set correct");
            }
        }

        [TestMethod]
        public void CreateDocumentFromTemplateStream()
        {
            using (var stream = new MemoryStream(Properties.Resources.Default))
            {
                using (var reader = new StreamReader(stream))
                {
                    var doc = TextDocument.Create(stream);

                    Assert.IsNotNull(doc, "Text Document not created");
                    if (doc != null)
                    {
                        doc.Filename = "test.docx";
                        Assert.AreEqual(doc.Filename, "test.docx", "Filename not set correct");
                    }
                }
            }
        }

        [TestMethod]
        [ExpectedException(typeof(FileNotFoundException), "File not found exception was not thrown")]
        public void CreateNamedExistingDocument()
        {
            var filename = string.Format("{0}\\{1}.{2}", DOCUMENT_PATH, DOCUMENT_NAME, DOCUMENT_EXTENSION);
            TextDocument doc = TextDocument.Create(Path.GetFullPath(filename), false);
        }

        [TestMethod]
        [ExpectedException(typeof(FileNotFoundException), "File not found exception was not thrown")]
        public void CreateExistingDocumentFromTemplate()
        {
            var templateLocation = string.Format("{0}\\{1}.{2}", TEMPLATE_PATH, TEMPLATE_NAME, TEMPLATE_EXTENSION);
            var filename = string.Format("{0}\\{1}.{2}", DOCUMENT_PATH, DOCUMENT_NAME, DOCUMENT_EXTENSION);
            TextDocument doc = TextDocument.Create(Path.GetFullPath(filename), Path.GetFullPath(templateLocation), false);
        }
        #endregion Create new documents.

        #region Open existing document
        [TestMethod]
        public void OpenDocument()
        {
            var doc = TextDocument.Open(_filename);
            Assert.IsNotNull(doc, "Text Document not created");
            if (doc != null)
            {
                Assert.AreEqual(doc.Filename, _filename, "Filename not set correct");
            }
        }

        [TestMethod]
        [ExpectedException(typeof(FileNotFoundException), "File not found exception was not thrown")]
        public void OpenNotExistingDocument()
        {
            var filename = "void";
            TextDocument doc = TextDocument.Open(filename);
        }
        #endregion  Open existing document
    }
}
