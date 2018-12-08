using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DocumentGenerator.WordDocuments;
using System.Collections.Generic;
using System.IO;
using DocumentGenerator.Common;
using System.Text;

namespace DocumentGeneratorTests
{
    [TestClass]
    public class SaveDocuments
    {
        #region Fields
        #endregion Fields

        #region Constants
        private const string FILE_PATH = "Resources";
        private const string FILE_NAME_WORD = "Test.docx";
        private const string FILE_NAME_WORD_FULL_DOCUMENT = "Full Text.docx";
        private const string FILE_NAME_ODF = "Test.odt";
        private const string FILE_NAME_ODF_FULL_DOCUMENT = "Full Text.odt";
        private const string TEMPLATE_NAME = "Default";
        private const string TEMPLATE_PATH = "Resources";
        private const string TEMPLATE_EXTENSION = "dotx";
        private const string DOCUMENT_NAME = "Document";
        private const string DOCUMENT_PATH = "Resources";
        private const string DOCUMENT_EXTENSION = "dotx";
        private const string TITLE_IMAGE_NAME = "TitlePage";
        private const string PORTRAIT_IMAGE_NAME = "Shakespeare";
        private const string IMAGE_EXTENSION = "jpg";
        private const string IMAGE_PATH = "Resources";
        #endregion  Constants

        #region Initialize
        [ClassInitialize]
        public static void InitializeClass(TestContext context)
        {
            //Clean up the resources directory in the bin directory.
            if (!Directory.Exists(FILE_PATH)) Directory.CreateDirectory(FILE_PATH);
            foreach (var file in Directory.GetFiles(FILE_PATH))
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

            //Write images to the resource directory
            var title_Location = string.Format("{0}\\{1}.{2}", IMAGE_PATH, TITLE_IMAGE_NAME, IMAGE_EXTENSION);
            var portrait_location = string.Format("{0}\\{1}.{2}", IMAGE_PATH, PORTRAIT_IMAGE_NAME, IMAGE_EXTENSION);
            var title = Properties.Resources.Sonnets1609titlepage;
            var portrait = Properties.Resources.Shakespeare;

            title.Save(title_Location);
            portrait.Save(portrait_location);

            title.Dispose();
            portrait.Dispose();
        }

        [TestInitialize]
        public void Initialize()
        {
            //Get the path for the document and template file.
            var templateLocation = string.Format("{0}\\{1}.{2}", TEMPLATE_PATH, TEMPLATE_NAME, TEMPLATE_EXTENSION);
            var filename = string.Format("{0}\\{1}.{2}", DOCUMENT_PATH, DOCUMENT_NAME, DOCUMENT_EXTENSION);

            //Remove the files if they exist
            if (File.Exists(templateLocation)) File.Delete(templateLocation);
            if (File.Exists(filename)) File.Delete(filename);

            //Write the document and template file to the resource directory
            File.WriteAllBytes(templateLocation, Properties.Resources.Default);
            File.WriteAllBytes(filename, Properties.Resources.Document);
        } 
        #endregion Initialize

        #region Document Content
        private void TwoTablesDocument(TextDocument document)
        {
            AddLinearTableToDocument(document, "Table 1");
            AddTableWithDifferentNumberOfColumnsPerRowToDocument(document, "Table 2");
        }

        private void AddAllContentToDocument(TextDocument document)
        {
            AddPictureFromFile(document, "The Title Page");

            //Create a paragraph with a header
            var p = document.AddNamedParagraph("Introduction", SonnetIntroduction(), 1, "1");
            p.SetParagraphType("Introdcution");

            //Create a paragraph with a header
            var p1 = document.AddNamedParagraph("Sonnet 1", Sonnet1(), 1, "1");
            p1.SetParagraphType("Sonnet");

            //Create a paragraph with a header on the level 0.
            var p0 = document.AddParagraph("Other sonnets", "Other sonnets from Shakespeare", 0);

            //Now we create a second pargraph as a sub paragraph of the previous one.
            //As the paragraphs are added sequentialy this paragraph will be directly added after p1 and because the paragraph level is 2
            //it will become a subparagraph.
            var subOfP1 = document.AddNamedParagraph("Sonnet II", Sonnet2(), 2, "2");
            subOfP1.SetParagraphType("Sonnet");

            //Now we add only a header
            var header = document.AddParagraph("Sonnet III", string.Empty, 1);

            //Now add a paragraph with only text
            var text = document.AddParagraph("The text for sonnet III", 1);

            //You can also change the text of the paragraph.
            text.Text = Sonnet3();

            //Some basic text styles can also be performed.
            var styledParagraph = document.AddNamedParagraph("Commentary", CommentarySonnet3(), 1, "4");
            styledParagraph.SetParagraphType("Commentary");

            AddLinearTableToDocument(document, "Table 1");

            AddTableWithDifferentNumberOfColumnsPerRowToDocument(document, "Table 2");

            AddTableWithNoColumnDefinitionToDocument(document, "Table 3");

            //Add a numbered list
            AddNumberedList(document, "Name");

            //Add the same list but then as a bullet list
            AddBulletList(document, "Name");

            //Add an image
            AddPictureFromStream(document, "Shakespeare");

            //Add an image without a title
            var picture3 = document.AddPicture(Properties.Resources.Shakespeare, string.Empty);
        }

        private Table AddLinearTableToDocument(TextDocument document, string tableName)
        {
            //Lets add a table
            var table = document.AddTable(tableName);
            //With two columns
            table.AddColumns(new List<string>() { "Title", "Author" });
            //And two rows
            table.AddRow(new List<string>() { "Sonnet I", "Shakespeare" });
            table.AddRow(new List<string>() { "Sonnet II", "Shakespeare" });

            return table;
        }

        private Table AddTableWithDifferentNumberOfColumnsPerRowToDocument(TextDocument document, string tableName)
        {
            //Lets add another table with different number of columns per row.
            var table = document.AddTable(tableName);
            //With two columns
            table.AddColumns(new List<string>() { "Title", "Author" });
            //And two rows
            table.AddRow(new List<string>() { "Sonnet I", "Shakespeare" });
            table.AddRow(new List<string>() { "Sonnet II", "Shakespeare", "William" });

            return table;
        }

        private Table AddTableWithNoColumnDefinitionToDocument(TextDocument document, string tableName)
        {
            //Lets add another table with no column definitions.
            var table = document.AddTable(tableName);
            //And two rows
            table.AddRow(new List<string>() { "Sonnet I", "Shakespeare" });
            table.AddRow(new List<string>() { "Sonnet II", "Shakespeare" });

            return table;
        }

        private NumberedList AddNumberedList(TextDocument document, string name)
        {
            //Add a numbered list
            var numberedList = document.AddNumberedList();
            numberedList.AddItem(0, "Characters (Numbered List):");
            numberedList.AddItem(1, "Fair Youth");
            numberedList.AddItem(2, "For a closer look at the negative aspects of the poet's relationship with the young man and his mistress, please see Sonnet 75 and Sonnet 147.");
            numberedList.AddItem(2, "For a celebration of the love between the young man and the poet, see Sonnet 18 and Sonnet 29.");
            numberedList.AddItem(2, "For the poet's views on the mortality of the young man, see Sonnet 73.");
            numberedList.AddItem(1, "The Dark Lady");
            numberedList.AddItem(2, "For a closer look at the negative aspects of the poet's relationship with the young man and his mistress, please see Sonnet 75 and Sonnet 147.");
            numberedList.AddItem(2, "For the poet’s description of his mistress, see Sonnet 130.");
            numberedList.AddItem(1, "The Rival Poet");
            numberedList.AddItem(2, "For a celebration of the love between the young man and the poet, see Sonnet 18 and Sonnet 29.");

            return numberedList;
        }

        private NumberedList AddNumberedListFromLevel1(TextDocument document, string name)
        {
            //Add a numbered list
            var numberedList = document.AddNumberedList();
            numberedList.AddItem(1, "Characters");
            numberedList.AddItem(2, "Fair Youth");
            numberedList.AddItem(3, "For a closer look at the negative aspects of the poet's relationship with the young man and his mistress, please see Sonnet 75 and Sonnet 147.");
            numberedList.AddItem(3, "For a celebration of the love between the young man and the poet, see Sonnet 18 and Sonnet 29.");
            numberedList.AddItem(3, "For the poet's views on the mortality of the young man, see Sonnet 73.");
            numberedList.AddItem(2, "The Dark Lady");
            numberedList.AddItem(3, "For a closer look at the negative aspects of the poet's relationship with the young man and his mistress, please see Sonnet 75 and Sonnet 147.");
            numberedList.AddItem(3, "For the poet’s description of his mistress, see Sonnet 130.");
            numberedList.AddItem(2, "The Rival Poet");
            numberedList.AddItem(3, "For a celebration of the love between the young man and the poet, see Sonnet 18 and Sonnet 29.");

            return numberedList;
        }

        private BulletList AddBulletList(TextDocument document, string name)
        {
            //Add the same list but then as a bullet list
            var bulletList = document.AddBulletList();
            bulletList.AddItem(0, "Most popular sonnets (bullets):");
            bulletList.AddItem(1, "Sonnet 126");
            bulletList.AddItem(2, "O thou my lovely boy");
            bulletList.AddItem(1, "Sonnet 130");
            bulletList.AddItem(2, "My Mistress' eyes");
            bulletList.AddItem(1, "Sonnet 029");
            bulletList.AddItem(2, "When in disgrace with fortune");
            bulletList.AddItem(1, "Sonnet 116");
            bulletList.AddItem(2, "Let me not to the marriage of true minds");
            bulletList.AddItem(1, "Sonnet 18");
            bulletList.AddItem(2, "Shall I compare thee to a Summer's day? ");

            return bulletList;
        }

        private BulletList AddBulletListFromLevel1(TextDocument document, string name)
        {
            //Add the same list but then as a bullet list
            var bulletList = document.AddBulletList();
            bulletList.AddItem(1, "Most popular sonnets:");
            bulletList.AddItem(2, "Sonnet 126 O thou my lovely boy");
            bulletList.AddItem(3, "O thou my lovely boy");
            bulletList.AddItem(2, "Sonnet 130");
            bulletList.AddItem(3, "My Mistress' eyes");
            bulletList.AddItem(2, "Sonnet 029");
            bulletList.AddItem(3, "When in disgrace with fortune");
            bulletList.AddItem(2, "Sonnet 116");
            bulletList.AddItem(3, "Let me not to the marriage of true minds");
            bulletList.AddItem(2, "Sonnet 18");
            bulletList.AddItem(3, "Shall I compare thee to a Summer's day? ");

            return bulletList;
        }

        private Picture AddPictureFromStream(TextDocument document, string title)
        {
            return document.AddPicture(Properties.Resources.Shakespeare, title);
        }

        private Picture AddPictureFromFile(TextDocument document, string title)
        {
            //Add an image with filepath
            var title_location = string.Format("{0}\\{1}.{2}", IMAGE_PATH, TITLE_IMAGE_NAME, IMAGE_EXTENSION);
            return document.AddPicture(title_location, title);
        }

        private string SonnetIntroduction()
        {
            var text = new StringBuilder();
            text.Append("Shakespeare's sonnets are a collection of 154 sonnets, ");
            text.Append("dealing with themes such as the passage of time, love,");
            text.Append(" beauty and mortality, first published in a 1609 quarto entitled <i>SHAKE-SPEARES SONNETS</i>.:");
            text.Append(" <i>Never before imprinted</i>. (although sonnets 138 and 144 had previously been published in the 1599 miscellany <i>The Passionate Pilgrim</i>).");
            text.Append(" The quarto ends with \"A Lover's Complaint\" , a narrative poem of 47 seven-line stanzas written in rhyme royal.");

            text.Append("The first 17 poems, traditionally called the procreation sonnets, ");
            text.Append("are addressed to a young man urging him to marry and have children in order to immortalize his beauty by passing it to the next generation.");
            text.Append(" [1] Other sonnets express the speaker's love for a young man; brood upon loneliness, death, ");
            text.Append("and the transience of life; seem to criticise the young man for preferring a rival poet; ");
            text.Append("express ambiguous feelings for the speaker's mistress; and pun on the poet's name. ");
            text.Append("The final two sonnets are allegorical treatments of Greek epigrams referring to the \"little love-god\" Cupid. ");
            text.AppendLine("The publisher, <b>Thomas Thorpe</b>, entered the book in the Stationers' Register on 20 May 1609:");
            text.AppendLine("\tTho. Thorpe. Entred for his copie under the hands of master Wilson and master Lownes Wardenes a booke called Shakespeares sonnettes vjd.");
            text.Append("Whether Thorpe used an authorised manuscript from Shakespeare or an unauthorised copy is unknown. <i>George Eld</i> printed the quarto, ");
            text.Append("and the run was divided between the booksellers <b>William Aspley</b> and <b>John Wright</b>.");

            return text.ToString();
        }

        private string Sonnet1()
        {
            var text = new StringBuilder();
            text.AppendLine("From fairest creatures we desire increase");
            text.AppendLine("That thereby beauty's rose might never die,");
            text.AppendLine("But as the riper should by time decease,");
            text.AppendLine("His tender heir might bear his memory:");
            text.AppendLine("But thou contracted to thine own bright eyes,");
            text.AppendLine("Feed'st thy light's flame with self-substantial fuel,");
            text.AppendLine("Making a famine where abundance lies,");
            text.AppendLine("Thy self thy foe, to thy sweet self too cruel:");
            text.AppendLine("Thou that art now the world's fresh ornament,");
            text.AppendLine("And only herald to the gaudy spring,");
            text.AppendLine("Within thine own bud buriest thy content,");
            text.AppendLine("And, tender churl, mak'st waste in niggarding:");
            text.AppendLine("Pity the world, or else this glutton be,");
            text.AppendLine("To eat the world's due, by the grave and thee.");

            return text.ToString();
        }

        private string Sonnet2()
        {
            var text = new StringBuilder();
            text.AppendLine("When forty winters shall besiege thy brow,");
            text.AppendLine("And dig deep trenches in thy beauty's field,");
            text.AppendLine("Thy youth's proud livery so gazed on now,");
            text.AppendLine("Will be a totter'd weed of small worth held:");
            text.AppendLine("Then being asked, where all thy beauty lies,");
            text.AppendLine("Where all the treasure of thy lusty days; ");
            text.AppendLine("To say, within thine own deep sunken eyes,");
            text.AppendLine("Were an all-eating shame, and thriftless praise.");
            text.AppendLine("How much more praise deserv'd thy beauty's use,");
            text.AppendLine("If thou couldst answer 'This fair child of mine");
            text.AppendLine("Shall sum my count, and make my old excuse,'");
            text.AppendLine("Proving his beauty by succession thine!");
            text.AppendLine("This were to be new made when thou art old,");
            text.AppendLine("And see thy blood warm when thou feel'st it cold.");

            return text.ToString();
        }

        private string Sonnet3()
        {
            var text = new StringBuilder();
            text.AppendLine("Look in thy glass and tell the face thou viewest");
            text.AppendLine("Now is the time that face should form another;");
            text.AppendLine("Whose fresh repair if now thou not renewest,");
            text.AppendLine("Thou dost beguile the world, unbless some mother.");
            text.AppendLine("For where is she so fair whose uneared womb");
            text.AppendLine("Disdains the tillage of thy husbandry?");
            text.AppendLine("Or who is he so fond will be the tomb");
            text.AppendLine("Of his self-love, to stop posterity? ");
            text.AppendLine("Thou art thy mother's glass and she in thee");
            text.AppendLine("Calls back the lovely April of her prime;");
            text.AppendLine("So thou through windows of thine age shalt see,");
            text.AppendLine("Despite of wrinkles, this thy golden time.");
            text.AppendLine("But if thou live, remembered not to be,");
            text.AppendLine("Die single and thine image dies with thee.");

            return text.ToString();
        }

        private string CommentarySonnet3()
        {
            var text = new StringBuilder();
            text.AppendLine("<b>1.</b> <u>Look in thy glass and tell the face thou viewest</u>");
            text.AppendLine("\t<u>glass</u> =\t<i>mirror</i>; glass in the Sonnets usually means <s>minor</s> mirror.");
            text.AppendLine("\t<u>the face thou viewest</u> =\t<i>your reflection</i>. I.e. speak to yourself and tell yourself that 'Now is the time etc'.");
            return text.ToString();
        }

        #endregion Document Content.

        #region Save document
        [TestMethod]
        public void SaveNewDocumentInOOXML()
        {
            var fileName = string.Format("{0}\\{1}", FILE_PATH, FILE_NAME_WORD_FULL_DOCUMENT);
            var document = TextDocument.Create(fileName, true);
            AddAllContentToDocument(document);
            //TwoTablesDocument(document);
            document.Save(DocumentGenerator.Common.DocumentType.OOXMLTextDocument);

            Assert.IsTrue(File.Exists(fileName), "File not saved");
            
        }

        [TestMethod]
        public void SaveNewDocumentInODF()
        {
            var fileName = string.Format("{0}\\{1}", FILE_PATH, FILE_NAME_ODF_FULL_DOCUMENT);
            var document = TextDocument.Create(fileName, true);
            AddAllContentToDocument(document);

            document.Save(DocumentGenerator.Common.DocumentType.ODFTextDocument);

            Assert.IsTrue(File.Exists(fileName), "File not saved");
        }

        [TestMethod]
        public void SaveNewDocumentInOOXMLWithNoRowsOrCols()
        {
            var fileName = string.Format("{0}\\{1}", FILE_PATH, FILE_NAME_WORD);
            var document = TextDocument.Create(fileName, true);
            AddLinearTableToDocument(document, "Table 1");
            document.Tables[0].RemoveColumn(0);
            document.Tables[0].RemoveColumn(0);
            document.Tables[0].RemoveRow(0);
            document.Tables[0].RemoveRow(0);

            document.Save(DocumentGenerator.Common.DocumentType.OOXMLTextDocument);

            Assert.IsTrue(File.Exists(fileName), "File not saved");

        }

        [TestMethod]
        public void SaveNewDocumentInODFWithNoRowsOrCols()
        {
            var fileName = string.Format("{0}\\{1}", FILE_PATH, FILE_NAME_ODF);
            var document = TextDocument.Create(fileName, true);
            AddLinearTableToDocument(document, "Table 1");
            document.Tables[0].RemoveColumn(0);
            document.Tables[0].RemoveColumn(0);
            document.Tables[0].RemoveRow(0);
            document.Tables[0].RemoveRow(0);

            document.Save(DocumentGenerator.Common.DocumentType.ODFTextDocument);

            Assert.IsTrue(File.Exists(fileName), "File not saved");
        }

        [TestMethod]
        [ExpectedException(typeof(FileLoadException), "AODL did not return the error for unsupported format")]
        public void SaveNewDocumentInODFWithNotSupportedExtension()
        {
            var fileName = string.Format("{0}\\{1}", FILE_PATH, "Test.odf");
            var document = TextDocument.Create(fileName, true);

            document.Save(DocumentGenerator.Common.DocumentType.ODFTextDocument);
        }

        [TestMethod]
        public void SaveDocumentFromTemplateStream()
        {
            using (var stream = new MemoryStream(Properties.Resources.Default))
            {
                using (var reader = new StreamReader(stream))
                {
                    var doc = TextDocument.Create(stream);

                    Assert.IsNotNull(doc, "Text Document not created");
                    if (doc != null)
                    {
                        var fileName = string.Format("{0}\\{1}", FILE_PATH, FILE_NAME_WORD);
                        doc.Filename = fileName;
                        doc.Save(DocumentGenerator.Common.DocumentType.OOXMLTextDocument);
                        Assert.IsTrue(File.Exists(fileName), "File not saved");
                    }
                }
            }
        }

        [TestMethod]
        public void SaveDocumentFromTemplateFile()
        {
            var templateLocation = string.Format("{0}\\{1}.{2}", TEMPLATE_PATH, TEMPLATE_NAME, TEMPLATE_EXTENSION);
            var fileName = string.Format("{0}\\{1}", FILE_PATH, FILE_NAME_WORD);
            TextDocument doc = TextDocument.Create(fileName, templateLocation, true);
            if (doc != null)
            {
                doc.Save(DocumentGenerator.Common.DocumentType.OOXMLTextDocument);
                Assert.IsTrue(File.Exists(fileName), "File not saved");
            }
        }

        [TestMethod]
        [ExpectedException(typeof(FilenameEmptyException), "File saved without a filename")]
        public void SaveDocumentWithoutFilename()
        {
            var document = TextDocument.Create();
            document.Save(DocumentGenerator.Common.DocumentType.OOXMLTextDocument);
        }

        [TestMethod]
        public void SaveDocumentFromTemplateWithCustomTags()
        {
            using (var stream = new MemoryStream(Properties.Resources.The_Sonnets_Template))
            {
                using (var reader = new StreamReader(stream))
                {
                    var doc = TextDocument.Create(stream);
                    doc.Filename = TEMPLATE_PATH +  "\\Shakespeare Sonnets.docx";

                    doc.ReplaceCustomTag("Author", "W. Shakespeare");
                    doc.ReplaceCustomTag("Title", "The Sonnets");
                    doc.ReplaceCustomTag("AuthorImage", Properties.Resources.Shakespeare);

                    doc.AddParagraph("Sonnet I", Sonnet1(), 1);
                    doc.AddParagraph("Sonnet II", Sonnet2(), "normal", "Heading1");

                    AddNumberedList(doc, string.Empty);
                    AddBulletList(doc, string.Empty);

                    doc.Save(DocumentType.OOXMLTextDocument);
                }
            }
        }

        [TestMethod]
        public void SaveDocumentFromTemplateWithOnlyRowsDefinedInTable()
        {
            using (var stream = new MemoryStream(Properties.Resources.Default))
            {
                using (var reader = new StreamReader(stream))
                {
                    var doc = TextDocument.Create(stream);
                    doc.Filename = TEMPLATE_PATH + "\\Test.docx";

                    AddTableWithNoColumnDefinitionToDocument(doc, "Table 1");

                    doc.Save(DocumentType.OOXMLTextDocument);
                }
            }
        }
        #endregion Save document
    }
}
