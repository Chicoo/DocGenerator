using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DocumentGenerator.WordDocuments;
using System.IO;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using DocumentGenerator.Common;
using System.Security.Cryptography;

namespace DocumentGeneratorTests
{
    [TestClass]
    public class EditTextDocumentTests
    {
        #region Constants
        private const string WHALE_IMAGE_NAME = "Humpback_Whale";
        private const string WHALE_IMAGE_EXTENSION = "jpg";
        private const string TURTLE_IMAGE_NAME = "Green_Sea_Turtle";
        private const string TURTLE_IMAGE_EXTENSION = "jpg";
        private const string IMAGE_PATH = "Resources";
        #endregion Constants

        #region Fields
        //The document used in the tests.
        TextDocument _document;

        //The id of the paragraph used in the tests.
        string _paragraphId; 
        #endregion Fields

        #region Intialize
        [TestInitialize]
        public void Initialize()
        {
            _document = TextDocument.Create();
            _paragraphId = Guid.NewGuid().ToString();

            //Clean up the resources directory in the bin directory.
            if (!Directory.Exists(IMAGE_PATH)) Directory.CreateDirectory(IMAGE_PATH);
            foreach (var file in Directory.GetFiles(IMAGE_PATH))
            {
                File.Delete(file);
            }

            //Write the document and template file to the resource directory
            var whale_Location = string.Format("{0}\\{1}.{2}", IMAGE_PATH, WHALE_IMAGE_NAME, WHALE_IMAGE_EXTENSION);
            var turtle_location = string.Format("{0}\\{1}.{2}", IMAGE_PATH, TURTLE_IMAGE_NAME, TURTLE_IMAGE_EXTENSION);
            var whale = Properties.Resources.Sonnets1609titlepage;
            var turtle = Properties.Resources.Shakespeare;

            whale.Save(whale_Location);
            turtle.Save(turtle_location);

            whale.Dispose();
            turtle.Dispose();
        } 
        #endregion Intialize

        #region Add paragraphs
        [TestMethod]
        public void AddParagraphAtTopLevel()
        {
            var paragraph = _document.AddParagraph("This is the paragraph", 0);
            Assert.IsNotNull(paragraph, "Paragraph not created");
            Assert.IsTrue(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.AreSame(paragraph, _document.Paragraphs[0], "Paragraph not added correct to the document");
        }

        [TestMethod]
        public void AddParagraph()
        {
            var paragraph = _document.AddParagraph("This is the paragraph", 1);
            Assert.IsNotNull(paragraph, "Paragraph not created");
            Assert.IsTrue(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.AreSame(paragraph, _document.Paragraphs[0], "Paragraph not added correct to the document");
        }

        [TestMethod]
        public void AddTwoParagraphs()
        {
            var paragraph1 = _document.AddParagraph("This is the first paragraph", 1);
            var paragraph2 = _document.AddParagraph("This is the second paragraph", 1);
            Assert.IsNotNull(paragraph1, "Paragraph 1 not created");
            Assert.IsNotNull(paragraph2, "Paragraph 2 not created");
            Assert.IsTrue(_document.Paragraphs.Count == 2, "Paragraphs not added to the document");
            Assert.AreSame(paragraph1, _document.Paragraphs[0], "Paragraph 1 not added correct to the document");
            Assert.AreSame(paragraph2, _document.Paragraphs[1], "Paragraph 2 not added correct to the document");
        }

        [TestMethod]
        public void AddParagraphWithHeading()
        {
            var paragraph = _document.AddParagraph("This is the paragraph", "Header for Paragraph", 1);
            Assert.IsNotNull(paragraph, "Paragraph not created");
            Assert.IsTrue(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.AreSame(paragraph, _document.Paragraphs[0], "Paragraph not added correct to the document");
        }

        [TestMethod]
        public void AddParagraphWithStyle()
        {
            var paragraph = _document.AddParagraph("This is the paragraph", "Normal");
            Assert.IsNotNull(paragraph, "Paragraph not created");
            Assert.IsTrue(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.AreSame(paragraph, _document.Paragraphs[0], "Paragraph not added correct to the document");
        }

        [TestMethod]
        public void AddParagraphWithHeadingAndStyle()
        {
            var paragraph = _document.AddParagraph("This is the paragraph", "Header for Paragraph", "Normal", "Header1");
            Assert.IsNotNull(paragraph, "Paragraph not created");
            Assert.IsTrue(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.AreSame(paragraph, _document.Paragraphs[0], "Paragraph not added correct to the document");
        }

        [TestMethod]
        public void AddNamedParagraph()
        {
            var paragraph = _document.AddNamedParagraph("This is the paragraph", 1, _paragraphId);
            Assert.IsNotNull(paragraph, "Paragraph not created");
            Assert.IsTrue(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.IsTrue(((NamedParagraph)_document.Paragraphs[0]).Id == _paragraphId, "The id of the paragraph is not set correct");
            Assert.AreSame(paragraph, _document.Paragraphs[0], "Paragraph not added correct to the document");
        }

        [TestMethod]
        public void AddNamedParagraphWithHeading()
        {
            var paragraph = _document.AddNamedParagraph("This is the paragraph", "Header for Paragraph", 1, _paragraphId);
            Assert.IsNotNull(paragraph, "Paragraph not created");
            Assert.IsTrue(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.IsTrue(((NamedParagraph)_document.Paragraphs[0]).Id == _paragraphId, "The id of the paragraph is not set correct");
            Assert.AreSame(paragraph, _document.Paragraphs[0], "Paragraph not added correct to the document");
        }

        [TestMethod]
        public void AddNamedParagraphWithStyle()
        {
            var paragraph = _document.AddNamedParagraph("This is the paragraph", "Normal", _paragraphId);
            Assert.IsNotNull(paragraph, "Paragraph not created");
            Assert.IsTrue(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.IsTrue(((NamedParagraph)_document.Paragraphs[0]).Id == _paragraphId, "The id of the paragraph is not set correct");
            Assert.AreSame(paragraph, _document.Paragraphs[0], "Paragraph not added correct to the document");
        }

        [TestMethod]
        public void AddNamedParagraphWithHeadingAndStyle()
        {
            var paragraph = _document.AddNamedParagraph("This is the paragraph", "Header for Paragraph", "Normal", "Header1", _paragraphId);
            Assert.IsNotNull(paragraph, "Paragraph not created");
            Assert.IsTrue(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.IsTrue(((NamedParagraph)_document.Paragraphs[0]).Id == _paragraphId, "The id of the paragraph is not set correct");
            Assert.AreSame(paragraph, _document.Paragraphs[0], "Paragraph not added correct to the document");
        }
        #endregion Add paragraphs

        #region Edit paragraphs
        [TestMethod]
        public void EditParagraphText()
        {
            //Create a paragraph

            var paragraph = _document.AddParagraph("This is the text", 1);
            if (paragraph == null) return;
            //Change the text
            paragraph.Text = "This is the new text";

            //Check the change
            if (_document.Paragraphs.Count != 1) return;
            var text = _document.Paragraphs[0].Text;
            Assert.IsTrue(text == "This is the new text", "The paragraph text has not been changed");
        }

        [TestMethod]
        public void AddParagraphType()
        {
            //Create a paragraph

            var paragraph = _document.AddNamedParagraph("This is the paragraph", 1, Guid.NewGuid().ToString());
            if (paragraph == null) return;
            //Change the text
            paragraph.SetParagraphType("Text");

            //Check the change
            if (_document.Paragraphs.Count != 1) return;
            var type = ((NamedParagraph)_document.Paragraphs[0]).GetParagraphType();
            Assert.IsTrue(type == "Text", "The paragraph type has not been changed");

            type = ((NamedParagraph)_document.Paragraphs[0]).Type;
            Assert.IsTrue(type == "Text", "The paragraph type has not been changed");
        }
        #endregion Edit paragraphs

        #region Lists
        [TestMethod]
        public void CreateNumberedList()
        {
            var numberedList = _document.AddNumberedList();
            numberedList.AddItem(0, "Numbered Item 1");
            numberedList.AddItem(0, "Numbered Item 2");
            numberedList.AddItem(1, "Numbered Item 2a");
            numberedList.AddItem(2, "Numbered Item 2aa");
            numberedList.AddItem(0, "Numbered Item 3");

            var list = _document.Paragraphs[0] as NumberedList;
            var item1 = list.Items[0];
            var item2 = list.Items[1];
            var item3 = list.Items[2];
            var item4 = list.Items[3];
            var item5 = list.Items[4];

            Assert.IsTrue(_document.Paragraphs.Count == 1, "List not added to document");
            Assert.IsTrue(list.Items.Count == 5, "Not the correct number of items in the list");
            Assert.IsTrue(item1.Text == "Numbered Item 1", "Text of item 1 in the list is not correct");
            Assert.IsTrue(item1.Level == 0, "Level of item 1 in the list is not correct");
            Assert.IsTrue(item2.Text == "Numbered Item 2", "Text of item 2 in the list is not correct");
            Assert.IsTrue(item2.Level == 0, "Level of item 2 in the list is not correct");
            Assert.IsTrue(item3.Text == "Numbered Item 2a", "Text of item 3 in the list is not correct");
            Assert.IsTrue(item3.Level == 1, "Level of item 3 in the list is not correct");
            Assert.IsTrue(item4.Text == "Numbered Item 2aa", "Text of item 4 in the list is not correct");
            Assert.IsTrue(item4.Level == 2, "Level of item 4 in the list is not correct");
            Assert.IsTrue(item5.Text == "Numbered Item 3", "Text of item 5 in the list is not correct");
            Assert.IsTrue(item5.Level == 0, "Level of item 5 in the list is not correct");
        }

        [TestMethod]
        public void CreateBullitList()
        {
            //Add the same list but then as a bullet list
            var bulletList = _document.AddBulletList();
            bulletList.AddItem(0, "Bullet Item 1");
            bulletList.AddItem(0, "Bullet Item 2");
            bulletList.AddItem(1, "Bullet Item 2a");
            bulletList.AddItem(2, "Bullet Item 2aa");
            bulletList.AddItem(0, "Bullet Item 3");

            var list = _document.Paragraphs[0] as BulletList;
            var item1 = list.Items[0];
            var item2 = list.Items[1];
            var item3 = list.Items[2];
            var item4 = list.Items[3];
            var item5 = list.Items[4];

            Assert.IsTrue(_document.Paragraphs.Count == 1, "List not added to document");
            Assert.IsTrue(list.Items.Count == 5, "Not the correct number of items in the list");
            Assert.IsTrue(item1.Text == "Bullet Item 1", "Text of item 1 in the list is not correct");
            Assert.IsTrue(item1.Level == 0, "Level of item 1 in the list is not correct");
            Assert.IsTrue(item2.Text == "Bullet Item 2", "Text of item 2 in the list is not correct");
            Assert.IsTrue(item2.Level == 0, "Level of item 2 in the list is not correct");
            Assert.IsTrue(item3.Text == "Bullet Item 2a", "Text of item 3 in the list is not correct");
            Assert.IsTrue(item3.Level == 1, "Level of item 3 in the list is not correct");
            Assert.IsTrue(item4.Text == "Bullet Item 2aa", "Text of item 4 in the list is not correct");
            Assert.IsTrue(item4.Level == 2, "Level of item 4 in the list is not correct");
            Assert.IsTrue(item5.Text == "Bullet Item 3", "Text of item 5 in the list is not correct");
            Assert.IsTrue(item5.Level == 0, "Level of item 5 in the list is not correct");
        }
        #endregion Lists

        #region Images and Tables
        [TestMethod]
        public void AddPictureFromFile()
        {
            var whale_Location = string.Format("{0}\\{1}.{2}", IMAGE_PATH, WHALE_IMAGE_NAME, WHALE_IMAGE_EXTENSION);
            var image = _document.AddPicture(whale_Location, "The whale");

            Assert.IsNotNull(image, "Images not created");
            Assert.IsTrue(_document.Paragraphs.Count == 1, "Image not added to the document");
            
            var document_image = _document.Paragraphs[0] as Picture;
            Assert.IsTrue(document_image.Text == "The whale", "Image title is not correct");
            Assert.AreSame(image, _document.Paragraphs[0], "Image not added correct to the document");
        }

        [TestMethod]
        public void AddPictureFromObject()
        {
            var image = _document.AddPicture(Properties.Resources.Shakespeare, "Shakespeare");

            Assert.IsNotNull(image, "Images not created");
            Assert.IsTrue(_document.Paragraphs.Count == 1, "Image not added to the document");

            var document_image = _document.Paragraphs[0] as Picture;
            Assert.IsTrue(document_image.Text == "Shakespeare", "Image title is not correct");
            Assert.AreSame(image, _document.Paragraphs[0], "Image not added correct to the document");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException), "The document did not return the correct excpetion")]
        public void AddEmptyPictureObject()
        {
            Bitmap picture = null;
            var image = _document.AddPicture(picture, "A Picture");
        }

        [TestMethod]
        [ExpectedException(typeof(FileNotFoundException), "The document did not return the correct excpetion")]
        public void AddNonExistingPictureFile()
        {
            var image = _document.AddPicture("Picture", "A Picture");
        }

        [TestMethod]
        [ExpectedException(typeof(NotImplementedException), "The document did not return the correct excpetion")]
        public void GetOOXMLParagraphForPicture()
        {
            var image = _document.AddPicture(Properties.Resources.Shakespeare, "Shakespeare");
            var paragraph = image.GetOOXMLParagraph();
        }

        [TestMethod]
        public void AddTable()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            Assert.IsNotNull(table, "Table is not created");
            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            Assert.IsTrue(_document.Paragraphs.Count == 1, "Table not added to the document");
            Assert.IsTrue(table.ColCount == 2, "Number of columns is incorrect");
            Assert.IsTrue(table.RowCount == 2, "Number of rows is incorrect");

            Assert.IsTrue(table.Text == "Table 1", "Table name not correct");
            Assert.IsTrue(table.Col[0] == "col 1", "The text of column 1 is not correct");
            Assert.IsTrue(table.Col[1] == "col 2", "The text of column 2 is not correct");
            Assert.IsTrue(table.Row[0][0] == "row 1 col 1", "The text of row 1 in column 1 is not correct");
            Assert.IsTrue(table.Row[0][1] == "row 1 col 2", "The text of row 1 in column 2 is not correct");
            Assert.IsTrue(table.Row[1][0] == "row 2 col 1", "The text of row 2 in column 1 is not correct");
            Assert.IsTrue(table.Row[1][1] == "row 2 col 2", "The text of row 2 in column 2 is not correct");

            Assert.AreSame(table, _document.Paragraphs[0], "The table is not added correctly to the document");
        }

        [TestMethod]
        public void ChangeTableRow()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            table.Row[0][0] = "Row 0 Col 0";

            Assert.IsTrue(table.Row[0][0] == "Row 0 Col 0", "The text of row 1 in column 1 is not changed correctly.");
        }

        [TestMethod]
        public void RemoveTableRow()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            table.RemoveRow(0);

            Assert.IsTrue(table.RowCount == 1, "The row has not been removed");
            Assert.IsTrue(table.Row[0][0] == "row 2 col 1", "The correct row has not been removed.");
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException), "Index out of range not thrown for deleting row")]
        public void RemoveTableRowIndexLow()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            table.RemoveRow(-1);
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException), "Index out of range not thrown for deleting row")]
        public void RemoveTableRowIndexHigh()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            table.RemoveRow(3);
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException), "An error was not thrown while getting an invalid column through the row.")]
        public void GetInvalidRowColumnWithHighIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            var value = table.Row[0][2];
        }

      
        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException), "An error was not thrown while setting an invalid column through the row.")]
        public void SetInvalidRowColumnWithHighIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            table.Row[0][2] = "Row 0 Col 0";
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException), "An error was not thrown while getting an invalid column through the row.")]
        public void GetInvalidRowColumnWithNegativeIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            var value = table.Row[0][-1];
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException), "An error was not thrown while setting an invalid column through the row.")]
        public void SetInvalidRowColumnWithNegativeIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            table.Row[0][-1] = "Row 0 Col 0";
        }

        [TestMethod]
        public void ChangeTableColumns()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            table.Col[0] = "Col 0";

            Assert.IsTrue(table.Col[0] == "Col 0", "The text of column 1 is not changed correctly.");
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException), "An error was not thrown while getting an invalid column")]
        public void GetInvalidColumnWithHighIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            var value = table.Col[2];
        }

        [TestMethod]
        public void RemoveTableCol()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            table.RemoveColumn(0);

            Assert.IsTrue(table.ColCount == 1, "The column has not been removed");
            Assert.IsTrue(table.Col[0] == "col 2", "The correct column has not been removed.");
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException), "Index out of range not thrown for deleting row")]
        public void RemoveTableColIndexLow()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            table.RemoveColumn(-1);
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException), "Index out of range not thrown for deleting row")]
        public void RemoveTableColIndexHigh()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            table.RemoveColumn(3);
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException), "An error was not thrown while setting an invalid column")]
        public void SetInvalidColumnWithHighIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            table.Col[2] = "Col 2";
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException), "An error was not thrown while getting an invalid column")]
        public void GetInvalidColumnWithNegativeIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            var value = table.Col[-1];
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException), "An error was not thrown while setting an invalid column")]
        public void SetInvalidColumnWithNegativeIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            table.Col[-1] = "Col -1";
        }
        #endregion Images and Tables

        #region ContentControls
        [TestMethod]
        public void AddCustomTag()
        {
            _document.ReplaceCustomTag("Title", "Hamlet");
            _document.ReplaceCustomTag("Author", "W. Shakespeare");

            Assert.IsTrue(_document.CustomParts.Count() == 2, "Custom parts not added to the document");
            var part1 = _document.CustomParts.ToList()[0];
            var part2 = _document.CustomParts.ToList()[1];

            Assert.IsTrue(part1.Name == "Title", "The name of the first custom part not added correctly");
            Assert.IsTrue(part1.Value.ToString() == "Hamlet", "The value of the first custom part not added correctly");
            Assert.IsTrue(part2.Name == "Author", "The name of the second custom part not added correctly");
            Assert.IsTrue(part2.Value.ToString() == "W. Shakespeare", "The value of the second custom part not added correctly");
        }

        [TestMethod]
        public void AddCustomImageTag()
        {
            _document.ReplaceCustomTag("Shakespeare", Properties.Resources.Shakespeare);

            Assert.IsTrue(_document.CustomParts.Count() == 1, "Custom part not added to the document");
            var part1 = _document.CustomParts.ToList()[0];

            var image = part1.Value as Image;

            //Convert each image to a byte array 

            ImageConverter ic = new ImageConverter();
            byte[] btImageExpected = new byte[1];
            btImageExpected = (byte[])ic.ConvertTo(Properties.Resources.Shakespeare, btImageExpected.GetType());
            byte[] btImageActual = new byte[1];
            btImageActual = (byte[])ic.ConvertTo(image, btImageActual.GetType());

            //Compute a hash for each image 

            var shaM = new SHA256Managed();
            byte[] hash1 = shaM.ComputeHash(btImageExpected);
            byte[] hash2 = shaM.ComputeHash(btImageActual);


            Assert.IsTrue(part1.Name == "Shakespeare", "The title of the image custom part not added correctly");
            //Compare the hash values 

            for (int i = 0; i < hash1.Length && i < hash2.Length; i++)
            {
                if (hash1[i] != hash2[i])
                    throw new AssertFailedException(
                     string.Format("Images are not the same. Expected:<hash value " + hash1[i] + ">. Actual:<hash value " + hash2[i] + ">. "));

            } 
        }

        [TestMethod]
        [ExpectedException(typeof(DuplicateCustomTagException), "No exception thrown when a duplicate tag is added")]
        public void AddDuplicateCustomTag()
        {
            _document.ReplaceCustomTag("Title", "Hamlet");
            _document.ReplaceCustomTag("Title", "The Merchant of Venice");
        }

        [TestMethod]
        [ExpectedException(typeof(DuplicateCustomTagException), "No exception thrown when a duplicate tag is added")]
        public void AddDuplicateImageCustomTag()
        {
            _document.ReplaceCustomTag("AuthorImage", Properties.Resources.Shakespeare);
            _document.ReplaceCustomTag("AuthorImage", Properties.Resources.Shakespeare);
        }
        #endregion ContentControls
    }
}
