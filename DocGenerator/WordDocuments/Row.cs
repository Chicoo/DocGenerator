using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using OOXMLParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using DocumentFormat.OpenXml;
using System.Globalization;
using AODL.Document.Content.Text;

namespace DocumentGenerator.WordDocuments
{
    /// <summary>
    /// This class defines a row in a table.
    /// </summary>
    public class Row
    {
        #region Properties
        private List<string> _values;
        /// <summary>
        /// The text of a row for a given index
        /// </summary>
        /// <param name="index">The index of the row</param>
        /// <returns>The row text.</returns>
        public string this[int index]
        {
            get 
            {
                if (index < 0 || index > _values.Count - 1)
                {
                    throw new IndexOutOfRangeException(string.Format(CultureInfo.CurrentCulture, "Index must be between {0} and {1}", 0, (_values.Count - 1)));
                }
                return _values[index]; 
            }
            set 
            {
                if (index < 0 || index > _values.Count - 1)
                {
                    throw new IndexOutOfRangeException(string.Format(CultureInfo.CurrentCulture, "Index must be between {0} and {1}", 0, (_values.Count - 1)));
                }
                _values[index] = value; 
            }
        }

        /// <summary>
        /// Returns the number of columns in the row
        /// </summary>
        internal int RowColcount
        {
            get { return _values.Count; }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new row with values
        /// </summary>
        /// <param name="values">The column values for the row</param>
        internal Row(List<string> values)
        {
            _values = values;
        }
        #endregion

        /// <summary>
        /// Creates an OOXML Table row.
        /// </summary>
        /// <returns></returns>
        internal TableRow GetOOXML()
        {
            List<OpenXmlElement> cells = new List<OpenXmlElement>();
            foreach(string s in _values)
            {
                TableCell cell = new TableCell(
                            new TableCellProperties(new TableCellWidth(){ Type = TableWidthUnitValues.Auto }),
                            new OOXMLParagraph(new Run(new Text(s))));
                cells.Add(cell);
            }
           TableRow row= new TableRow(cells);
            return row;
        }

        /// <summary>
        /// Creates a ODF table row.
        /// </summary>
        /// <param name="row">The row in the ODF document.</param>
        /// <returns>The filled ODF row.</returns>
        internal AODL.Document.Content.Tables.Row GetODF(AODL.Document.Content.Tables.Row row)
        {
            foreach (string s in _values)
            {
                    //Create a standard paragraph
                    var paragraph = ParagraphBuilder.CreateStandardTextParagraph(row.Document);
                    //Add the text
                    foreach (var formatedText in CommonDocumentFunctions.ParseParagraphForODF(row.Document, s))
                    {
                        paragraph.TextContent.Add(formatedText);
                    }
                    //Add the content to the cell
                    //row.Cells[_values.IndexOf(s)].Content.Add(paragraph);
            }
            return row;
        }
    }
}
