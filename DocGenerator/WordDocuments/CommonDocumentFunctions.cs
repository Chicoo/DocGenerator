using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using AODL.Document;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using System.Drawing;
using System.Drawing.Imaging;

namespace DocumentGenerator.WordDocuments
{
    /// <summary>
    /// Contains common functions to generate a word document.
    /// </summary>
    class CommonDocumentFunctions
    {
        #region Textpart struct
        /// <summary>
        /// Contains information about text parts
        /// </summary>
        private struct TextPart
        {

            private int _startIndex;
            public int StartIndex { get { return _startIndex; } }

            private int _length;
            public int Length { get { return _length; } }

            private object _element;
            public OpenXmlElement OOXMLElement { get { return _element as OpenXmlElement; } }
            public AODL.Document.Content.Text.FormatedText ODFElement { get { return _element as AODL.Document.Content.Text.FormatedText; } }

            public TextPart(int startIndex, int length, OpenXmlElement element)
            {
                _startIndex = startIndex;
                _length = length;
                _element = element;
            }

            public TextPart(int startIndex, int length, AODL.Document.Content.Text.FormatedText element)
            {
                _startIndex = startIndex;
                _length = length;
                _element = element;
            }
        }
        #endregion

        #region Regular Expressions
        private static readonly Regex BoldRegex = new Regex(@"<b>.+?</b>", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex ItalicRegex = new Regex(@"<i>.+?</i>", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex BoldItalicRegex = new Regex(@"<b><i>'.+?</i></b>", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex UnderlinedRegex = new Regex(@"<u>.+?</u>", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex StrikedRegex = new Regex(@"<s>.+?</s>", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex TabRegex = new Regex(@"\t", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex NewLineRegex = new Regex(@"\r\n", RegexOptions.Compiled | RegexOptions.Singleline);
        #endregion

        #region Document parts creation
        /// <summary>
        /// Parses a string in OOXML format. The function replaces the foolowing special tags to format the text:
        ///     Bold:<b></b> 
        ///     Italic:<i></i> 
        ///     Underline:<u></u> 
        ///     StrikeThrough:<s></s>
        ///     Bullet list:<ul></ul> and <li></li> for items.
        ///     Tab: \t
        /// </summary>
        /// <param name="paragraph">The text to parse.</param>
        /// <returns>A list pf OpenXMLElements that make up the paragraph.</returns>
        internal static List<OpenXmlElement> ParseParagraphForOOXML(string paragraph)
        {
            List<TextPart> pars = new List<TextPart>();

            MatchCollection matches;

            //Bold Text
            matches = BoldRegex.Matches(paragraph);
            if (matches.Count > 0)
            {
                foreach (Match m in matches)
                {
                    var text = new Run(
                               new RunProperties(
                                   new Bold()),
                               new Text(m.Value.Substring(3, m.Value.Length - 7)) { Space = SpaceProcessingModeValues.Preserve }
                               );

                    pars.Add(new TextPart(m.Index, m.Value.Length, text));
                }
            }

            //Italic Text
            matches = ItalicRegex.Matches(paragraph);
            if (matches.Count > 0)
            {
                foreach (Match m in matches)
                {
                    var text = new Run(
                               new RunProperties(
                                   new Italic()),
                               new Text(m.Value.Substring(3, m.Value.Length - 7)) { Space = SpaceProcessingModeValues.Preserve }
                               );

                    pars.Add(new TextPart(m.Index, m.Value.Length, text));
                }
            }


            //Underlined Text
            matches = UnderlinedRegex.Matches(paragraph);
            if (matches.Count > 0)
            {
                foreach (Match m in matches)
                {
                    var text = new Run(
                               new RunProperties(
                                   new Underline() { Val = UnderlineValues.Single }),
                               new Text(m.Value.Substring(3, m.Value.Length - 7)) { Space = SpaceProcessingModeValues.Preserve }
                               );

                    pars.Add(new TextPart(m.Index, m.Value.Length, text));
                }
            }


            //Strike through Text
            matches = StrikedRegex.Matches(paragraph);
            if (matches.Count > 0)
            {
                foreach (Match m in matches)
                {
                    var text = new Run(
                               new RunProperties(
                                   new Strike()),
                               new Text(m.Value.Substring(3, m.Value.Length - 7)) { Space = SpaceProcessingModeValues.Preserve }
                               );

                    pars.Add(new TextPart(m.Index, m.Value.Length, text));
                }
            }

            //Tabs
            matches = TabRegex.Matches(paragraph);
            if (matches.Count > 0)
            {
                foreach (Match m in matches)
                {
                    var text = new Run(
                               new TabChar()
                               );

                    pars.Add(new TextPart(m.Index, m.Value.Length, text));
                }
            }

            //Create the list of OpenXmlElements in the correct order for the paragraph
            List<OpenXmlElement> p = new List<OpenXmlElement>();
            int index = 0;

            var sorted = from par in pars
                         orderby par.StartIndex
                         select par;

            foreach (var element in sorted)
            {
                //If the current index is not equal to the start index of the text paert then there is plain text in between.
                if (element.StartIndex > index)
                {
                    var text = new Run(
                   new Text(paragraph.Substring(index, element.StartIndex - index)) { Space = SpaceProcessingModeValues.Preserve }
                   );
                    p.Add(text);
                }

                p.Add(element.OOXMLElement);
                index = element.StartIndex + element.Length;
            }

            //Add any left over text
            if (paragraph.Substring(index).Length > 0)
            {
                var text = new Run(
                 new Text(paragraph.Substring(index)) { Space = SpaceProcessingModeValues.Preserve }
                 );
                p.Add(text);
            }

            return p;
        }

        /// <summary>
        /// Parses a string in ODF format. The function replaces the foolowing special tags to format the text:
        ///     Bold:<b></b> 
        ///     Italic:<i></i> 
        ///     Underline:<u></u> 
        ///     StrikeThrough:<s></s>
        /// </summary>
        /// <param name="paragraph">The text to parse.</param>
        /// <param name="doc">The ODF Textdoument that is being created</param>
        /// <returns>A list pf FormattedText that make up the paragraph.</returns>
        internal static List<AODL.Document.Content.Text.FormatedText> ParseParagraphForODF(IDocument doc, string paragraph)
        {
            List<TextPart> pars = new List<TextPart>();

            MatchCollection matches;

            //Bold Text
            matches = BoldRegex.Matches(paragraph);
            if (matches.Count > 0)
            {
                foreach (Match m in matches)
                {
                    //Create the formated text
                    var text = new AODL.Document.Content.Text.FormatedText(doc, "bold", m.Value.Substring(3, m.Value.Length - 7));
                    //Apply the style
                    text.TextStyle.TextProperties.Bold = "bold";
                    pars.Add(new TextPart(m.Index, m.Value.Length, text));
                }
            }

            //Italic Text
            matches = ItalicRegex.Matches(paragraph);
            if (matches.Count > 0)
            {
                foreach (Match m in matches)
                {
                    //Create the formated text
                    var text = new AODL.Document.Content.Text.FormatedText(doc, "italic", m.Value.Substring(3, m.Value.Length - 7));
                    //Apply the style
                    text.TextStyle.TextProperties.Italic = "italic";
                    pars.Add(new TextPart(m.Index, m.Value.Length, text));
                }
            }


            //Underlined Text
            matches = UnderlinedRegex.Matches(paragraph);
            if (matches.Count > 0)
            {
                foreach (Match m in matches)
                {
                    //Create the formated text
                    var text = new AODL.Document.Content.Text.FormatedText(doc, "underline", m.Value.Substring(3, m.Value.Length - 7));
                    //Apply the style
                    text.TextStyle.TextProperties.Underline = "solid";
                    pars.Add(new TextPart(m.Index, m.Value.Length, text));
                }
            }


            //Strike through Text
            matches = StrikedRegex.Matches(paragraph);
            if (matches.Count > 0)
            {
                foreach (Match m in matches)
                {
                    //Create the formated text
                    var text = new AODL.Document.Content.Text.FormatedText(doc, "strikethrough", m.Value.Substring(3, m.Value.Length - 7));
                    //Apply the style
                    text.TextStyle.TextProperties.TextLineThrough = "solid";
                    pars.Add(new TextPart(m.Index, m.Value.Length, text));
                }
            }

            //Create the list of OpenXmlElements in the correct order for the paragraph
            List<AODL.Document.Content.Text.FormatedText> p = new List<AODL.Document.Content.Text.FormatedText>();
            int index = 0;

            var sorted = from par in pars
                         orderby par.StartIndex
                         select par;

            foreach (var element in sorted)
            {
                //If the current index is not equal to the start index of the text part then there is plain text in between.
                if (element.StartIndex > index)
                {
                    //Create the formated text
                    var text = new AODL.Document.Content.Text.FormatedText(doc, "normal", paragraph.Substring(index, element.StartIndex - index));
                    p.Add(text);
                }
                p.Add(element.ODFElement);
                index = element.StartIndex + element.Length;
            }

            //Add any left over text
            if (paragraph.Substring(index).Length > 0)
            {
                //Create the formated text
                var text = new AODL.Document.Content.Text.FormatedText(doc, "T1", paragraph.Substring(index));
                p.Add(text);
            }

            return p;
        }

        internal static string AddPictureToOOXMLDocument(OpenXmlPartContainer mainPart, Image image, int index, out string imageName)
        {
            //Set variables for image
            var rid = string.Format(CultureInfo.CurrentCulture, "rId{0}", index);
            imageName = string.Format(CultureInfo.CurrentCulture, "Picture {0}", index);

            //Create the imagepart as an JPEG
            while (mainPart.Parts.Any(p => p.RelationshipId == rid))
            {
                rid = string.Format(CultureInfo.CurrentCulture, "image{0}", index++);
                imageName = string.Format(CultureInfo.CurrentCulture, "Picture {0}", index);
            }
            var imagePart = mainPart.AddNewPart<ImagePart>("image/jpeg", rid);

            //Write the image with in image part
            var bmp = new Bitmap(image);
            bmp.Save(imagePart.GetStream(), ImageFormat.Jpeg);

            return rid;
        } 
        #endregion Document parts creation
    }
}
