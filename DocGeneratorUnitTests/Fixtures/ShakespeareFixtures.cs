using System;
using System.Drawing;
using System.IO;
using Xunit;

namespace DocGenerator.UnitTests.Fixtures
{
    public class ShakespeareFixtures
    {
        /// <summary>
        /// Gest an <see cref="Image"/> of Shakespeare;
        /// </summary>
        public Image Shakespeare { get; }

        /// <summary>
        /// Gets an <see cref="Image"/> of the Sonnets title page
        /// </summary>
        public Image Sonnets1609titlepage { get; }

        /// <summary>
        /// Gets an <see cref="byte[]"/> of the default Word template (.dotx)
        /// </summary>
        public byte[] Default { get; }

        /// <summary>
        /// Gets an <see cref="byte[]"/> of a Word document (.docx)
        /// </summary>
        public byte[] Document { get; }

        /// <summary>
        /// Gets an <see cref="byte[]"/> of the Word template (.dotx) for the Sonnets.
        /// </summary>
        public byte[] The_Sonnets_Template { get; }

        private const string WHALE_IMAGE_NAME = "Humpback_Whale";
        private const string WHALE_IMAGE_EXTENSION = "jpg";
        private const string TURTLE_IMAGE_NAME = "Green_Sea_Turtle";
        private const string TURTLE_IMAGE_EXTENSION = "jpg";
        private const string IMAGE_PATH = "TestData";

        private const string TEMPLATE_NAME = "Default";
        private const string TEMPLATE_PATH = "TestData";
        private const string TEMPLATE_EXTENSION = "dotx";
        private const string DOCUMENT_NAME = "Document";
        private const string DOCUMENT_PATH = "TestData";
        private const string DOCUMENT_EXTENSION = "dotx";
        private const string TITLE_IMAGE_NAME = "TitlePage";
        private const string PORTRAIT_IMAGE_NAME = "Shakespeare";
        private const string IMAGE_EXTENSION = "jpg";

        public ShakespeareFixtures()
        {
            Shakespeare = Image.FromFile("Resources\\Shakespeare.jpg");
            Sonnets1609titlepage = Image.FromFile("Resources\\Sonnets1609titlepage.jpg");

            Default = File.ReadAllBytes("Resources\\Default.dotx");
            Document = File.ReadAllBytes("Resources\\Document.docx");
            The_Sonnets_Template = File.ReadAllBytes("Resources\\The Sonnets Template.dotx");

            Initialize();
        }

        private void Initialize()
        {
            //Clean up the resources directory in the bin directory.
            if (!Directory.Exists(IMAGE_PATH)) Directory.CreateDirectory(IMAGE_PATH);
            foreach (var file in Directory.GetFiles(IMAGE_PATH))
            {
                File.Delete(file);
            }

            //Delete the temp dirtectory for odf images.
            string dir = string.Format("{0}\\DocumentGenerator", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData));
            if (Directory.Exists(dir))
            {
                foreach (var file in Directory.GetFiles(dir))
                {
                    File.Delete(file);
                }
                Directory.Delete(dir);
            }

            //Write the document and template file to the resource directory
            var whale_Location = string.Format("{0}\\{1}.{2}", IMAGE_PATH, WHALE_IMAGE_NAME, WHALE_IMAGE_EXTENSION);
            var turtle_location = string.Format("{0}\\{1}.{2}", IMAGE_PATH, TURTLE_IMAGE_NAME, TURTLE_IMAGE_EXTENSION);
            var whale = Sonnets1609titlepage;
            var turtle = Shakespeare;

            whale.Save(whale_Location);
            turtle.Save(turtle_location);

            //Write images to the resource directory
            var title_Location = string.Format("{0}\\{1}.{2}", IMAGE_PATH, TITLE_IMAGE_NAME, IMAGE_EXTENSION);
            var portrait_location = string.Format("{0}\\{1}.{2}", IMAGE_PATH, PORTRAIT_IMAGE_NAME, IMAGE_EXTENSION);
            var title = Sonnets1609titlepage;
            var portrait = Shakespeare;

            title.Save(title_Location);
            portrait.Save(portrait_location);

            //Get the path for the document and template file.
            var templateLocation = string.Format("{0}\\{1}.{2}", TEMPLATE_PATH, TEMPLATE_NAME, TEMPLATE_EXTENSION);
            var filename = string.Format("{0}\\{1}.{2}", DOCUMENT_PATH, DOCUMENT_NAME, DOCUMENT_EXTENSION);

            //Remove the files if they exist
            if (File.Exists(templateLocation)) File.Delete(templateLocation);
            if (File.Exists(filename)) File.Delete(filename);

            //Write the document and template file to the resource directory
            File.WriteAllBytes(templateLocation, Default);
            File.WriteAllBytes(filename, Document);
        }
    }

    [CollectionDefinition("Shakespeare collection")]
    public class ShakespeareCollection : ICollectionFixture<ShakespeareFixtures>
    {
        // This class has no code, and is never created. Its purpose is simply
        // to be the place to apply [CollectionDefinition] and all the
        // ICollectionFixture<> interfaces.
    }

}
