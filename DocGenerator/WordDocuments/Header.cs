using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OOXMLParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using DocumentFormat.OpenXml.Wordprocessing;
using AODL.Document.Content;
using AODL.Document.Content.Text;

namespace DocumentGenerator.WordDocuments
{
    /// <summary>
    /// Represents the header in a document.
    /// Inherets from paragraph.
    /// </summary>
    public class Header : Paragraph
    {
        #region Constructors
        /// <summary>
        /// Creates a new header on a given heading level.
        /// </summary>
        /// <param name="text">The text of the header</param>
        /// <param name="headerLevel">The level of the header</param>
        public Header(string text, int headerLevel) : base(text, headerLevel)
        {
        }

        /// <summary>
        /// Creates a new header with a given st5yle.
        /// </summary>
        /// <param name="text">The text of the paragraph</param>
        /// <param name="styleName">The style of the header</param>
        public Header(string text, string styleName) : base(text, styleName)
        {

        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Gets a paragraph in ODF format.
        /// </summary>
        /// <param name="doc">The ODF Textdocument that is being created</param>
        /// <returns>List of paragraphs in ODF format.</returns>
        public override List<IContent> GetODFParagraph(AODL.Document.TextDocuments.TextDocument doc)
        {
            //Get the headings enum
            Headings headingEnum = Headings.Heading;
            string headingString = string.Format("Heading{0}{1}", _paragraphLevel > 0 ? "_20" : string.Empty, _paragraphLevel > 0 ? "_" + (_paragraphLevel ).ToString() : string.Empty);
            if(Enum.IsDefined(typeof(Headings), headingString))
            {
                headingEnum = (Headings)Enum.Parse(typeof(Headings), headingString);
            }

            //Create the header
            AODL.Document.Content.Text.Header header = new AODL.Document.Content.Text.Header(doc, headingEnum);
            //Add the formated text
            foreach (var formatedText in CommonDocumentFunctions.ParseParagraphForODF(doc, _text))
            {
                header.TextContent.Add(formatedText);
            }
            //Add the header properties
            return new List<IContent>() { header };

        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Add a style based on the given paragraph level.
        /// </summary>
        /// <param name="paragraphLevel">The level of the paragraph.</param>
        /// <returns></returns>
        protected override ParagraphProperties ooxmlParagraphProp(int paragraphLevel)
        {
            var paraProp = new ParagraphProperties();
            var styleId = new ParagraphStyleId
            {
                Val = string.Format(CultureInfo.CurrentCulture, "Heading{0}", paragraphLevel == 0 ? "1" : paragraphLevel.ToString(CultureInfo.CurrentCulture))
            };
            paraProp.AppendChild(styleId);

            return paraProp;
        }
        #endregion
    }
}
