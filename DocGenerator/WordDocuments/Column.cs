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
    /// This class defines the columns in a table.
    /// These will be used for the heading row.
    /// </summary>
    public class Column
    {
           #region Properties
        private List<string> _values;
        /// <summary>
        /// The text of a column for a given index
        /// </summary>
        /// <param name="index">The index of the column</param>
        /// <returns>The column text.</returns>
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
        /// Returns the number of columns
        /// </summary>
        internal int Colcount
        {
            get { return _values.Count; }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new list of columns
        /// </summary>
        /// <param name="values">The column headers</param>
        internal Column(List<string> values)
        {
            _values = values;
        }
        #endregion

        /// <summary>
        /// Removes the columns at the given index.
        /// </summary>
        /// <param name="index">The index of the column</param>
        /// <exception cref="IndexOutOfRangeException">If the index is out of the range of the columns.</exception>
        internal void Remove(int index)
        {
            if (index < 0 || index > _values.Count - 1)
            {
                throw new IndexOutOfRangeException(string.Format(CultureInfo.CurrentCulture, "Index must be between {0} and {1}", 0, (_values.Count - 1)));
            }
            _values.RemoveAt(index);
        }

        /// <summary>
        /// Creates an OOXML header row.
        /// </summary>
        /// <returns>An OOXML Table row.</returns>
        internal TableRow GetOOXML()
        {
            List<OpenXmlElement> cells = new List<OpenXmlElement>();
            //Tell the table that this is the header row.
            var properties = new TableRowProperties(new TableHeader());
            cells.Add(properties);

            foreach (string s in _values)
            {
                TableCell cell = new TableCell(
                            new TableCellProperties(
                                new TableCellWidth() { Type = TableWidthUnitValues.Auto },
                                new Shading(){ Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D9D9D9"}),
                            new OOXMLParagraph(new Run(new Text(s))));
                cells.Add(cell);
            }
            TableRow row = new TableRow(cells);
            return row;
        }

        /// <summary>
        /// Creates a ODF table header.
        /// This is the first row in the table
        /// </summary>
        /// <param name="row">The headerrow in the ODF document.</param>
        /// <returns>The filled ODF headerrow.</returns>
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
                //row.Cells[_values.IndexOf(s)].CellStyle.CellProperties.BackgroundColor = "#d9d9d9";
                //row.Cells[_values.IndexOf(s)].Content.Add(paragraph);
            }
            return row;
        }

        
    }


}
