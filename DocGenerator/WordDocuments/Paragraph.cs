using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using OOXMLParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using DocumentFormat.OpenXml.Wordprocessing;

using ODFParagraph =  AODL.Document.Content.Text.Paragraph;
using AODL.Document.Content;

namespace DocumentGenerator.WordDocuments
{
    /// <summary>
    /// Represents the header in a document.
    /// </summary>
    public class Paragraph
    {
        #region Private Fields
        /// <summary>
        /// The text of the paragraph.
        /// </summary>
        protected string  _text;
        /// <summary>
        /// The level of the paragraph.
        /// This determines the style of the paragraph.
        /// </summary>
        protected int _paragraphLevel = -1;
        protected string _styleName = string.Empty;
        private readonly Header _header;
        #endregion

        #region Properties
        /// <summary>
        /// The text of the paragraph.
        /// </summary>
        public string Text
        {
            get {return _text;}
            set { _text = value; }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new text paragrah on a given heading level, without a header.
        /// </summary>
        /// <param name="text">The text of the paragraph.</param>
        /// <param name="paragraphLevel">The level of the paragraph.</param>
        public Paragraph(string text, int paragraphLevel)
        {
            _text = text;
            _paragraphLevel = paragraphLevel;
        }

        /// <summary>
        /// Creates a new paragrah with a header and text and on a given heading level.
        /// </summary>
        /// <param name="headerText">The header of the paragrah</param>
        /// <param name="text">The text of the paragraph</param>
        /// <param name="paragraphLevel">The level of the paragraph</param>
        public Paragraph(string text, string headerText, int paragraphLevel) : this(text, paragraphLevel)
        {
            _header = new Header(headerText, paragraphLevel);
        }

        /// <summary>
        /// Creates a new paragrah with a header and text with a give style for the header and the paragraph.
        /// </summary>
        /// <param name="headerText">The header of the paragrah</param>
        /// <param name="text">The text of the paragraph</param>
        /// <param name="paragraphStyle">The style of the paragraph</param>
        /// <param name="headerStyle">The style of the header</param>
        public Paragraph(string headerText, string text, string paragraphStyle, string headerStyle)
            : this(text, paragraphStyle)
        {
            _header = new Header(headerText, headerStyle);
        }

        /// <summary>
        /// Creates a new text paragrah with a given style, without a header.
        /// </summary>
        /// <param name="text">The text of the paragraph.</param>
        /// <param name="paragraphStyle">The style of the paragraph.</param>
        public Paragraph(string text, string paragraphStyle)
        {
            _text = text;
            _styleName = paragraphStyle;
        }

        /// <summary>
        /// Creates a new paragrah
        /// </summary>
        public Paragraph()
        {
        } 
        #endregion

        #region Public Methods
        /// <summary>
        /// Returns an empty OOXML paragraph.
        /// </summary>
        /// <returns></returns>
        public static OOXMLParagraph EmptyParagraph()
        {
            return new OOXMLParagraph();
        }

        /// <summary>
        /// Gets the paragraph(s) in OOXMLFormat.
        /// This is returned as a list as an object may have multiple paragraphs.
        /// </summary>
        /// <returns>List of paragraphs in OOXML Format.</returns>
        public virtual List<OOXMLParagraph> GetOOXMLParagraph()
        {
            List<OOXMLParagraph> paragraphs = new List<OOXMLParagraph>();
            var header = CreateOOXMLHeaderPart();
            if(header != null) paragraphs.Add(header);

            //Split the paragraph by the lines and create a separate one for each line.
            var lines = Regex.Split(Text, "\r\n");

            //Process each line as a paragraph.
            paragraphs.AddRange(lines.Select(line => CreateOOXMLTextPart(line)).Where(text => text!=null));

            return paragraphs;
        }

        /// <summary>
        /// Gets a paragraph in ODF format.
        /// </summary>
        /// <param name="doc">The ODF Textdocument that is being created</param>
        /// <returns>List of paragraphs in ODF format.</returns>
        public virtual List<IContent> GetODFParagraph(AODL.Document.TextDocuments.TextDocument doc)
        {
            List<IContent> paragraphs = new List<IContent>();

            //Create the header, if there is any.
            if (_header != null)
            {
                paragraphs.Add(_header.GetODFParagraph(doc)[0]);
            }

            //Create the text part if there is any
            if (!string.IsNullOrEmpty(_text))
            {
                //Create the paragraph
                ODFParagraph paragraph = AODL.Document.Content.Text.ParagraphBuilder.CreateStandardTextParagraph(doc);
                //Add the formated text
                foreach(var formatedText in CommonDocumentFunctions.ParseParagraphForODF(doc, _text))
                {
                    paragraph.TextContent.Add(formatedText);
                }
                //Add the header properties
                paragraphs.Add(paragraph);
            }
            return paragraphs;
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Add a style based on the given paragraph level.
        /// </summary>
        /// <param name="paragraphLevel">The level of the paragraph.</param>
        /// <returns></returns>
        protected virtual ParagraphProperties ooxmlParagraphProp(int paragraphLevel)
        {
            var paraProp = new ParagraphProperties();
            var styleId = new ParagraphStyleId
                              {
                                  Val=string.Format(CultureInfo.CurrentCulture, "normal{0}", paragraphLevel==0 ? string.Empty : paragraphLevel.ToString(CultureInfo.CurrentCulture))
                              };
            paraProp.AppendChild(styleId);

            return paraProp;
        }

        /// <summary>
        /// Add a style based on the given style name.
        /// </summary>
        /// <param name="style">The name of the style.</param>
        /// <returns></returns>
        protected virtual ParagraphProperties ooxmlParagraphProp(string style)
        {
            var paraProp = new ParagraphProperties();
            var styleId = new ParagraphStyleId
            {
                Val = string.Format(CultureInfo.CurrentCulture, style)
            };
            paraProp.AppendChild(styleId);

            return paraProp;
        }

        /// <summary>
        /// Returns the header part of the paragraph.
        /// </summary>
        /// <returns>The header or null if no header is defined.</returns>
        protected OOXMLParagraph CreateOOXMLHeaderPart()
        {
            if (_header != null) return _header.CreateOOXMLTextPart(_header.Text);
            return null;
        }

        /// <summary>
        /// Creates the text part of the paragraph.
        /// </summary>
        /// <param name="text">The text of the pargraph.</param>
        /// <returns>The paragraph or null if no text is defiend</returns>
        protected OOXMLParagraph CreateOOXMLTextPart(string text)
        {
            //Create the text part if there is any
            if (!string.IsNullOrEmpty(text))
            {
                OOXMLParagraph paragraph = new OOXMLParagraph(CommonDocumentFunctions.ParseParagraphForOOXML(text));
                //Add the header properties
                if(_paragraphLevel >= 0) paragraph.PrependChild(ooxmlParagraphProp(_paragraphLevel));
                else if (!string.IsNullOrWhiteSpace(_styleName)) paragraph.PrependChild(ooxmlParagraphProp(_styleName));
                else paragraph.PrependChild(ooxmlParagraphProp(0));
                return paragraph;
            }

            return null;
        }
        #endregion
    }
}
