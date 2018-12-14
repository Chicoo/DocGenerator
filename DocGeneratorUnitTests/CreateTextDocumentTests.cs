using Xunit;
using DocumentGenerator.WordDocuments;
using System.IO;
using DocGenerator.UnitTests.Fixtures;

namespace DocGenerator.UnitTests
{
    [Collection("Shakespeare collection")]
    public class CreateTextDocumentTests
    {
        private const string TEMPLATE_NAME = "Default";
        private const string TEMPLATE_PATH = "TestData";
        private const string TEMPLATE_EXTENSION = "dotx";
        private const string DOCUMENT_NAME = "Document";
        private const string DOCUMENT_PATH = "TestData";
        private const string DOCUMENT_EXTENSION = "dotx";

        private string _filename;

        ShakespeareFixtures _shakespeareFixtures;

        public CreateTextDocumentTests(ShakespeareFixtures shakespeareFixtures)
        {
            _shakespeareFixtures = shakespeareFixtures;

            _filename = string.Format("{0}\\{1}.{2}", DOCUMENT_PATH, DOCUMENT_NAME, DOCUMENT_EXTENSION);
        } 

        [Fact]
        public void CreateEmptyDocument()
        {
            TextDocument doc = TextDocument.Create();
            Assert.NotNull(doc);
            if (doc != null)
            {
                doc.Filename = "test.docx";
                Assert.Equal("test.docx", doc.Filename);
            }
        }

        [Fact]
        public void CreateNamedDocument()
        {
            TextDocument doc = TextDocument.Create("test.docx", true);
            Assert.NotNull(doc);
            if (doc != null)
            {
                Assert.Equal("test.docx", doc.Filename);
            }
        }

        [Fact]
        public void CreateDocumentFromTemplate()
        {
            var templateLocation = string.Format("{0}\\{1}.{2}", TEMPLATE_PATH, TEMPLATE_NAME, TEMPLATE_EXTENSION);
            TextDocument doc = TextDocument.Create("test.docx", templateLocation, true);
            Assert.NotNull(doc);
            if (doc != null)
            {
                Assert.Equal("test.docx", doc.Filename);
            }
        }

        [Fact]
        public void CreateDocumentFromTemplateStream()
        {
            using (var stream = new MemoryStream(_shakespeareFixtures.Default))
            {
                using (var reader = new StreamReader(stream))
                {
                    var doc = TextDocument.Create(stream);

                    Assert.NotNull(doc);
                    if (doc != null)
                    {
                        doc.Filename = "test.docx";
                        Assert.Equal("test.docx", doc.Filename);
                    }
                }
            }
        }

        [Fact]
        public void CreateNamedExistingDocument()
        {
            var filename = string.Format("{0}\\{1}.{2}", DOCUMENT_PATH, DOCUMENT_NAME, DOCUMENT_EXTENSION);
            Assert.Throws<FileNotFoundException>(() => TextDocument.Create(Path.GetFullPath(filename), false));
        }

        [Fact]
        public void CreateExistingDocumentFromTemplate()
        {
            var templateLocation = string.Format("{0}\\{1}.{2}", TEMPLATE_PATH, TEMPLATE_NAME, TEMPLATE_EXTENSION);
            var filename = string.Format("{0}\\{1}.{2}", DOCUMENT_PATH, DOCUMENT_NAME, DOCUMENT_EXTENSION);
            Assert.Throws<FileNotFoundException>(() => TextDocument.Create(Path.GetFullPath(filename), Path.GetFullPath(templateLocation), false));
        }

        [Fact]
        public void OpenDocument()
        {
            var doc = TextDocument.Open(_filename);
            Assert.NotNull(doc);
            if (doc != null)
            {
                Assert.Equal(_filename, doc.Filename);
            }
        }

        [Fact]
        public void OpenNotExistingDocument()
        {
            var filename = "void";
            Assert.Throws<FileNotFoundException>(() => TextDocument.Open(filename));
        }
    }
}
