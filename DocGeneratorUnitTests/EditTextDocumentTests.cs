using System;
using Xunit;
using DocumentGenerator.WordDocuments;
using System.IO;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using DocumentGenerator.Common;
using DocGenerator.UnitTests.Fixtures;

namespace DocGenerator.UnitTests
{
    [Collection("Shakespeare collection")]
    public class EditTextDocumentTests
    {
        private const string WHALE_IMAGE_NAME = "Humpback_Whale";
        private const string WHALE_IMAGE_EXTENSION = "jpg";
        private const string TURTLE_IMAGE_NAME = "Green_Sea_Turtle";
        private const string TURTLE_IMAGE_EXTENSION = "jpg";
        private const string IMAGE_PATH = "TestData";

        //The document used in the tests.
        TextDocument _document;

        //The id of the paragraph used in the tests.
        readonly string _paragraphId;

        ShakespeareFixtures _shakespeareFixtures;

        public EditTextDocumentTests(ShakespeareFixtures shakespeareFixtures)
        {
            _shakespeareFixtures = shakespeareFixtures;

            _document = TextDocument.Create();
            _paragraphId = Guid.NewGuid().ToString();
        }

        

        [Fact]
        public void AddParagraphAtTopLevel()
        {
            var paragraph = _document.AddParagraph("This is the paragraph", 0);
            Assert.NotNull(paragraph);
            Assert.True(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.Same(paragraph, _document.Paragraphs[0]);
        }

        [Fact]
        public void AddParagraph()
        {
            var paragraph = _document.AddParagraph("This is the paragraph", 1);
            Assert.NotNull(paragraph);
            Assert.True(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.Same(paragraph, _document.Paragraphs[0]);
        }

        [Fact]
        public void AddTwoParagraphs()
        {
            var paragraph1 = _document.AddParagraph("This is the first paragraph", 1);
            var paragraph2 = _document.AddParagraph("This is the second paragraph", 1);
            Assert.NotNull(paragraph1);
            Assert.NotNull(paragraph2);
            Assert.True(_document.Paragraphs.Count == 2, "Paragraphs not added to the document");
            Assert.Same(paragraph1, _document.Paragraphs[0]);
            Assert.Same(paragraph2, _document.Paragraphs[1]);
        }

        [Fact]
        public void AddParagraphWithHeading()
        {
            var paragraph = _document.AddParagraph("This is the paragraph", "Header for Paragraph", 1);
            Assert.NotNull(paragraph);
            Assert.True(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.Same(paragraph, _document.Paragraphs[0]);
        }

        [Fact]
        public void AddParagraphWithStyle()
        {
            var paragraph = _document.AddParagraph("This is the paragraph", "Normal");
            Assert.NotNull(paragraph);
            Assert.True(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.Same(paragraph, _document.Paragraphs[0]);
        }

        [Fact]
        public void AddParagraphWithHeadingAndStyle()
        {
            var paragraph = _document.AddParagraph("This is the paragraph", "Header for Paragraph", "Normal", "Header1");
            Assert.NotNull(paragraph);
            Assert.True(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.Same(paragraph, _document.Paragraphs[0]);
        }

        [Fact]
        public void AddNamedParagraph()
        {
            var paragraph = _document.AddNamedParagraph("This is the paragraph", 1, _paragraphId);
            Assert.NotNull(paragraph);
            Assert.True(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.True(((NamedParagraph)_document.Paragraphs[0]).Id == _paragraphId, "The id of the paragraph is not set correct");
            Assert.Same(paragraph, _document.Paragraphs[0]);
        }

        [Fact]
        public void AddNamedParagraphWithHeading()
        {
            var paragraph = _document.AddNamedParagraph("This is the paragraph", "Header for Paragraph", 1, _paragraphId);
            Assert.NotNull(paragraph);
            Assert.True(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.True(((NamedParagraph)_document.Paragraphs[0]).Id == _paragraphId, "The id of the paragraph is not set correct");
            Assert.Same(paragraph, _document.Paragraphs[0]);
        }

        [Fact]
        public void AddNamedParagraphWithStyle()
        {
            var paragraph = _document.AddNamedParagraph("This is the paragraph", "Normal", _paragraphId);
            Assert.NotNull(paragraph);
            Assert.True(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.True(((NamedParagraph)_document.Paragraphs[0]).Id == _paragraphId, "The id of the paragraph is not set correct");
            Assert.Same(paragraph, _document.Paragraphs[0]);
        }

        [Fact]
        public void AddNamedParagraphWithHeadingAndStyle()
        {
            var paragraph = _document.AddNamedParagraph("This is the paragraph", "Header for Paragraph", "Normal", "Header1", _paragraphId);
            Assert.NotNull(paragraph);
            Assert.True(_document.Paragraphs.Count == 1, "Paragraph not added to the document");
            Assert.True(((NamedParagraph)_document.Paragraphs[0]).Id == _paragraphId, "The id of the paragraph is not set correct");
            Assert.Same(paragraph, _document.Paragraphs[0]);
        }

        [Fact]
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
            Assert.True(text == "This is the new text", "The paragraph text has not been changed");
        }

        [Fact]
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
            Assert.True(type == "Text", "The paragraph type has not been changed");

            type = ((NamedParagraph)_document.Paragraphs[0]).Type;
            Assert.True(type == "Text", "The paragraph type has not been changed");
        }

        [Fact]
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

            Assert.True(_document.Paragraphs.Count == 1, "List not added to document");
            Assert.True(list.Items.Count == 5, "Not the correct number of items in the list");
            Assert.True(item1.Text == "Numbered Item 1", "Text of item 1 in the list is not correct");
            Assert.True(item1.Level == 0, "Level of item 1 in the list is not correct");
            Assert.True(item2.Text == "Numbered Item 2", "Text of item 2 in the list is not correct");
            Assert.True(item2.Level == 0, "Level of item 2 in the list is not correct");
            Assert.True(item3.Text == "Numbered Item 2a", "Text of item 3 in the list is not correct");
            Assert.True(item3.Level == 1, "Level of item 3 in the list is not correct");
            Assert.True(item4.Text == "Numbered Item 2aa", "Text of item 4 in the list is not correct");
            Assert.True(item4.Level == 2, "Level of item 4 in the list is not correct");
            Assert.True(item5.Text == "Numbered Item 3", "Text of item 5 in the list is not correct");
            Assert.True(item5.Level == 0, "Level of item 5 in the list is not correct");
        }

        [Fact]
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

            Assert.True(_document.Paragraphs.Count == 1, "List not added to document");
            Assert.True(list.Items.Count == 5, "Not the correct number of items in the list");
            Assert.True(item1.Text == "Bullet Item 1", "Text of item 1 in the list is not correct");
            Assert.True(item1.Level == 0, "Level of item 1 in the list is not correct");
            Assert.True(item2.Text == "Bullet Item 2", "Text of item 2 in the list is not correct");
            Assert.True(item2.Level == 0, "Level of item 2 in the list is not correct");
            Assert.True(item3.Text == "Bullet Item 2a", "Text of item 3 in the list is not correct");
            Assert.True(item3.Level == 1, "Level of item 3 in the list is not correct");
            Assert.True(item4.Text == "Bullet Item 2aa", "Text of item 4 in the list is not correct");
            Assert.True(item4.Level == 2, "Level of item 4 in the list is not correct");
            Assert.True(item5.Text == "Bullet Item 3", "Text of item 5 in the list is not correct");
            Assert.True(item5.Level == 0, "Level of item 5 in the list is not correct");
        }

        [Fact]
        public void AddPictureFromFile()
        {
            var whale_Location = string.Format("{0}\\{1}.{2}", IMAGE_PATH, WHALE_IMAGE_NAME, WHALE_IMAGE_EXTENSION);
            var image = _document.AddPicture(whale_Location, "The whale");

            Assert.NotNull(image);
            Assert.True(_document.Paragraphs.Count == 1, "Image not added to the document");
            
            var document_image = _document.Paragraphs[0] as Picture;
            Assert.True(document_image.Text == "The whale", "Image title is not correct");
            Assert.Same(image, _document.Paragraphs[0]);
        }

        [Fact]
        public void AddPictureFromObject()
        {
            var image = _document.AddPicture(_shakespeareFixtures.Shakespeare, "Shakespeare");

            Assert.NotNull(image);
            Assert.True(_document.Paragraphs.Count == 1, "Image not added to the document");

            var document_image = _document.Paragraphs[0] as Picture;
            Assert.True(document_image.Text == "Shakespeare", "Image title is not correct");
            Assert.Same(image, _document.Paragraphs[0]);
        }

        [Fact]
        public void AddEmptyPictureObject()
        {
            Bitmap picture = null;
            var image = Assert.Throws<ArgumentNullException>(() => _document.AddPicture(picture, "A Picture"));
        }

        [Fact]
        public void AddNonExistingPictureFile()
        {
            var image = Assert.Throws<FileNotFoundException>(() => _document.AddPicture("Picture", "A Picture"));
        }

        [Fact]
        public void GetOOXMLParagraphForPicture()
        {
            var image = _document.AddPicture(_shakespeareFixtures.Shakespeare, "Shakespeare");
            var paragraph = Assert.Throws<NotImplementedException>(() => image.GetOOXMLParagraph());
        }

        [Fact]
        public void AddTable()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            Assert.NotNull(table);
            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            Assert.True(_document.Paragraphs.Count == 1, "Table not added to the document");
            Assert.True(table.ColCount == 2, "Number of columns is incorrect");
            Assert.True(table.RowCount == 2, "Number of rows is incorrect");

            Assert.True(table.Text == "Table 1", "Table name not correct");
            Assert.True(table.Col[0] == "col 1", "The text of column 1 is not correct");
            Assert.True(table.Col[1] == "col 2", "The text of column 2 is not correct");
            Assert.True(table.Row[0][0] == "row 1 col 1", "The text of row 1 in column 1 is not correct");
            Assert.True(table.Row[0][1] == "row 1 col 2", "The text of row 1 in column 2 is not correct");
            Assert.True(table.Row[1][0] == "row 2 col 1", "The text of row 2 in column 1 is not correct");
            Assert.True(table.Row[1][1] == "row 2 col 2", "The text of row 2 in column 2 is not correct");

            Assert.Same(table, _document.Paragraphs[0]);
        }

        [Fact]
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

            Assert.True(table.Row[0][0] == "Row 0 Col 0", "The text of row 1 in column 1 is not changed correctly.");
        }

        [Fact]
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

            Assert.True(table.RowCount == 1, "The row has not been removed");
            Assert.True(table.Row[0][0] == "row 2 col 1", "The correct row has not been removed.");
        }

        [Fact]
        public void RemoveTableRowIndexLow()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            Assert.Throws<IndexOutOfRangeException>(() => table.RemoveRow(-1));
        }

        [Fact]
        public void RemoveTableRowIndexHigh()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            Assert.Throws<IndexOutOfRangeException>(() => table.RemoveRow(3));
        }

        [Fact]
        public void GetInvalidRowColumnWithHighIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            var value = Assert.Throws<IndexOutOfRangeException>(() => table.Row[0][2]);
        }

      
        [Fact]
        public void SetInvalidRowColumnWithHighIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            Assert.Throws<IndexOutOfRangeException>(() => table.Row[0][2] = "Row 0 Col 0");
        }

        [Fact]
        public void GetInvalidRowColumnWithNegativeIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            var value = Assert.Throws<IndexOutOfRangeException>(() => table.Row[0][-1]);
        }

        [Fact]
        public void SetInvalidRowColumnWithNegativeIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            Assert.Throws<IndexOutOfRangeException>(() => table.Row[0][-1] = "Row 0 Col 0");
        }

        [Fact]
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

            Assert.True(table.Col[0] == "Col 0", "The text of column 1 is not changed correctly.");
        }

        [Fact]
        public void GetInvalidColumnWithHighIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            var value = Assert.Throws<IndexOutOfRangeException>(() => table.Col[2]);
        }

        [Fact]
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

            Assert.True(table.ColCount == 1, "The column has not been removed");
            Assert.True(table.Col[0] == "col 2", "The correct column has not been removed.");
        }

        [Fact]
        public void RemoveTableColIndexLow()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            Assert.Throws<IndexOutOfRangeException>(() => table.RemoveColumn(-1));
        }

        [Fact]
        public void RemoveTableColIndexHigh()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            Assert.Throws<IndexOutOfRangeException>(() => table.RemoveColumn(3));
        }

        [Fact]
        public void SetInvalidColumnWithHighIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            Assert.Throws<IndexOutOfRangeException>(() => table.Col[2] = "Col 2");
        }

        [Fact]
        public void GetInvalidColumnWithNegativeIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            var value = Assert.Throws<IndexOutOfRangeException>(() => table.Col[-1]);
        }

        [Fact]
        public void SetInvalidColumnWithNegativeIndex()
        {
            //Lets add a table
            var table = _document.AddTable("Table 1");

            //With two columns
            table.AddColumns(new List<string>() { "col 1", "col 2" });
            //And two rows
            table.AddRow(new List<string>() { "row 1 col 1", "row 1 col 2" });
            table.AddRow(new List<string>() { "row 2 col 1", "row 2 col 2" });

            Assert.Throws<IndexOutOfRangeException>(() => table.Col[-1] = "Col -1");
        }

        [Fact]
        public void AddCustomTag()
        {
            _document.ReplaceCustomTag("Title", "Hamlet");
            _document.ReplaceCustomTag("Author", "W. Shakespeare");

            Assert.True(_document.CustomParts.Count() == 2, "Custom parts not added to the document");
            var part1 = _document.CustomParts.ToList()[0];
            var part2 = _document.CustomParts.ToList()[1];

            Assert.True(part1.Name == "Title", "The name of the first custom part not added correctly");
            Assert.True(part1.Value.ToString() == "Hamlet", "The value of the first custom part not added correctly");
            Assert.True(part2.Name == "Author", "The name of the second custom part not added correctly");
            Assert.True(part2.Value.ToString() == "W. Shakespeare", "The value of the second custom part not added correctly");
        }

        [Fact]
        public void AddCustomImageTag()
        {
            _document.ReplaceCustomTag("Shakespeare", _shakespeareFixtures.Shakespeare);

            Assert.True(_document.CustomParts.Count() == 1, "Custom part not added to the document");
            var part1 = _document.CustomParts.ToList()[0];

            var image = part1.Value as Image;

            //Convert each image to a byte array 
            int hashExpected = _shakespeareFixtures.Shakespeare.GetHashCode();
            int hashActual = image.GetHashCode();

            Assert.True(part1.Name == "Shakespeare", "The title of the image custom part not added correctly");
            //Compare the hash values 
            Assert.True(hashExpected == hashActual, $"Images are not the same. Expected:<hash value {hashExpected}>. Actual:<hash value {hashActual}>. ");
        }

        [Fact]
        public void AddDuplicateCustomTag()
        {
            _document.ReplaceCustomTag("Title", "Hamlet");
            Assert.Throws<DuplicateCustomTagException>(() => _document.ReplaceCustomTag("Title", "The Merchant of Venice"));
        }

        [Fact]
        public void AddDuplicateImageCustomTag()
        {
            _document.ReplaceCustomTag("AuthorImage", _shakespeareFixtures.Shakespeare);
            Assert.Throws<DuplicateCustomTagException>(() => _document.ReplaceCustomTag("AuthorImage", _shakespeareFixtures.Shakespeare));
        }
    }
}
