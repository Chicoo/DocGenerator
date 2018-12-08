using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using OOXMLParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace DocumentGenerator.WordDocuments
{
    /// <summary>
    /// Represents a paragraph that is identified with an id.
    /// 
    /// </summary>
    public class NamedParagraph : Paragraph
    {
        #region Fields
        private string _type;
        #endregion Fields

        #region Properties
        private readonly string _id;
        /// <summary>
        /// Gets the id (name) of the paragraph.
        /// </summary>
        public string Id
        {
            get { return _id; }
        }

        /// <summary>
        /// Gets the type associated with the paragraph.
        /// </summary>
        public string Type
        {
            get { return _type; }
        }
        #endregion Properties

        #region Constructors
        /// <summary>
        /// Creates a new text paragrah on on given heading level, without a header.
        /// </summary>
        /// <param name="text">The text of the paragraph.</param>
        /// <param name="paragraphLevel">The level of the paragraph.</param>
        /// <param name="id">The id (name) of the paragraph</param>
        public NamedParagraph(string text, int paragraphLevel, string id) : base(text, paragraphLevel)
        {
            _id = id;
        }

        /// <summary>
        /// Creates a new paragrah with a header and text with a give style for the header and the paragraph.
        /// </summary>
        /// <param name="text">The text of the paragraph</param>
        /// <param name="paragraphStyle">The style of the paragraph</param>
        /// <param name="id">The id (name) of the paragraph</param>
        public NamedParagraph(string text, string paragraphStyle, string id) : base(text, paragraphStyle)
        {
            _id = id;
        }

        /// <summary>
        /// Creates a new paragrah with a header and text with a give style for the header and the paragraph.
        /// </summary>
        /// <param name="headerText">The header of the paragrah</param>
        /// <param name="text">The text of the paragraph</param>
        /// <param name="paragraphStyle">The style of the paragraph</param>
        /// <param name="headerStyle">The style of the header</param>
        /// <param name="id">The id (name) of the paragraph</param>
        public NamedParagraph(string text, string headerText, string paragraphStyle, string headerStyle, string id) : base(text, headerText, paragraphStyle, headerStyle)
        {
            _id = id;
        }

        /// <summary>
        /// Creates a new paragrah with a header and text and on on given heading level.
        /// </summary>
        /// <param name="headerText">The header of the paragrah</param>
        /// <param name="text">The text of the paragraph</param>
        /// <param name="paragraphLevel">The level of the paragraph</param>
        /// <param name="id">The id (name) of the paragraph</param>
        public NamedParagraph(string text, string headerText, int paragraphLevel, string id): base(text, headerText, paragraphLevel)
        {
            _id = id;
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Gets the paragraph(s) in OOXMLFormat.
        /// This is returned as a list as an object may have multiple paragraphs.
        /// </summary>
        /// <returns>List of paragraphs in OOXML Format.</returns>
        public override List<OOXMLParagraph> GetOOXMLParagraph()
        {
            List<OOXMLParagraph> paragraphs = new List<OOXMLParagraph>();
            var header = CreateOOXMLHeaderPart();
            if (header != null) paragraphs.Add(header);

            //Split the paragraph by the lines and create a separate one for each line.
            var lines = Regex.Split(Text, "\r\n");

            //Process each line as a paragraph.
            foreach(var line in lines)
            {
                var text = CreateOOXMLTextPart(line);
                if(text==null) continue;
                if (!string.IsNullOrWhiteSpace(_id)) text.SetAttribute(new OpenXmlAttribute("name", string.Empty, _id));
                if (!string.IsNullOrWhiteSpace(_type)) text.SetAttribute(new OpenXmlAttribute("type", string.Empty, _type));
                paragraphs.Add(text);
            }
            
            return paragraphs;
        }

        /// <summary>
        /// Get the type associated with this paragraph.
        /// </summary>
        /// <returns>A string representing the type of the paragraph.</returns>
        public string GetParagraphType()
        {
            return _type;
        }

        /// <summary>
        /// Set the type that is associated with this paragraph.
        /// </summary>
        /// <param name="type">The type associated with the paragraph.</param>
        public void SetParagraphType(string type)
        {
            _type = type;
        }
        #endregion Public Methods
    }
}
