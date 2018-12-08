using System;
using System.Collections.Generic;
using OOXMLTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using AODL.Document.Content;
using AODL.Document.Content.Tables;
using System.Globalization;

namespace DocumentGenerator.WordDocuments
{
    /// <summary>
    /// Represents a table in a document
    /// </summary>
    public class Table :Paragraph
    {
        #region Properties
        private readonly List<Row> _rows;
        /// <summary>
        /// Returns a list of rows in the table
        /// </summary>
        public IList<Row> Row
        {
            get { return _rows.AsReadOnly(); }
        }
        private Column column;
        /// <summary>
        /// Returns the columns in the table
        /// </summary>
        public Column Col
        {
            get { return column; }
        }

        /// <summary>
        /// Returns the number of rows in the table
        /// </summary>
        public int RowCount
        {
            get { return _rows.Count; }
        }

        /// <summary>
        /// Returns the number of columns in the table
        /// </summary>
        public int ColCount
        {
            get { return column.Colcount; }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a table for the document
        /// </summary>
        /// <param name="caption">The caption of the table</param>
        /// <param name="level">The paragraph level to show the table add. This sets the indent of the table.</param>
        public Table(string caption, int level):base(caption, level)
        {
            _rows = new List<Row>();
            column = new Column(new List<string>());
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Adds a row to the table
        /// </summary>
        /// <param name="values">The values for the row.</param>
        public void AddRow(List<string> values)
        {
            Row row = new Row(values);
            _rows.Add(row);
        }

        /// <summary>
        /// Adds a list of columns to the table
        /// </summary>
        /// <param name="columnNames">The column names</param>
        public void AddColumns(List<string> columnNames)
        {
            column = new Column(columnNames);
        }

        /// <summary>
        /// Removes the column at the given index.
        /// </summary>
        /// <param name="index">The index of the column.</param>
        /// <exception cref="IndexOutOfRangeException">If the index is out of the range of the columns.</exception>
        public void RemoveColumn(int index)
        {
            column.Remove(index);
        }

        /// <summary>
        /// Removes the row at the given index.
        /// </summary>
        /// <param name="index">The index of the row.</param>
        /// <exception cref="IndexOutOfRangeException">If the index is out of the range of the columns.</exception>
        public void RemoveRow(int index)
        {
            if (index < 0 || index >= _rows.Count)
            {
                throw new IndexOutOfRangeException(string.Format(CultureInfo.CurrentCulture, "Index must be between {0} and {1}", 0, (_rows.Count - 1)));
            }
            _rows.RemoveAt(index);
        }

        /// <summary>
        /// Gets the paragraph(s) in OOXMLFormat.
        /// This is returned as a list as an object may have multiple paragraphs.
        /// </summary>
        /// <returns>List of paragraphs in OOXML Format.</returns>
        public OOXMLTable GetOOXMLTable()
        {
            //Create the rows of the table
            List<OpenXmlElement> childElements = new List<OpenXmlElement>();

            //Create the table border.
            var tableBorders = new TableBorders(
                           new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                           new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                           new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                           new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                           new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                           new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U });

            childElements.Add(tableBorders);

            int numberOfColumns = 0; //Gets the number of columns to generate in the gridcolumns.
            //The first row are the column names if they are defined
            if (ColCount > 0)
            {
                childElements.Add(Col.GetOOXML());
                numberOfColumns = ColCount;
            }
            foreach (Row row in Row)
            {
                childElements.Add(row.GetOOXML());
                if (row.RowColcount > numberOfColumns) numberOfColumns = row.RowColcount;
            }
         
            //Create the table with the rows.
            var element = new OOXMLTable(childElements);
            return element;
        }

         /// <summary>
        /// Gets a paragraph in ODF format.
        /// </summary>
        /// <param name="doc">The ODF Textdocument that is being created</param>
        /// <returns>List of paragraphs in ODF format.</returns>
        public IContent GetODFTable(AODL.Document.TextDocuments.TextDocument doc)
        {
            //Get the max columns in a row .
            var colCount = 0;
            foreach (var row in _rows)
            {
                if (row.RowColcount > colCount) colCount = row.RowColcount;
            }

            //Create a table for a text document using the TableBuilder
            var table = TableBuilder.CreateTextDocumentTable(
                doc,
                Text,
                "table1",
                RowCount +1,
                colCount,
                16.99,
                Col.Colcount > 0,
                false);

            //Fill the cells
            foreach (var row in Row)
            {
                //Pass the current row in to Row object.
                //This seems the easiest way to fill them.
                //row.GetODF(table.Rows[Row.IndexOf(row)]);
            }

            //Fill the header
            if (Col != null &&Col.Colcount > 0 && table.RowHeader.RowCollection.Count > 0)
            {
                Col.GetODF(table.RowHeader.RowCollection[0]);
            }
            
            return table;
        }

       

        #endregion

    }
}
