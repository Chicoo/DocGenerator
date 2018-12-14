using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

using OOXMLParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml;
using a = DocumentFormat.OpenXml.Drawing;
using pic = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Globalization;
using AODL.Document.Content.Text;
using AODL.Document.Content.Draw;



namespace DocumentGenerator.WordDocuments
{
    /// <summary>
    /// Defines a picture in a document.
    /// </summary>
    public class Picture : Paragraph
    {
        #region Private Fields
        private readonly Image _image;
        private readonly string _filePath;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates an image with a title on on given heading level.
        /// </summary>
        /// <param name="image">The image to show in the document</param>
        /// <param name="title">The title of the image. If this is null or empty then no title will be shown.</param>
        /// <param name="paragraphLevel">The level of the paragraph</param>
        public Picture(Image image,  string title, int paragraphLevel) : base(title, paragraphLevel)
        {
            _image = image ?? throw new ArgumentNullException("The image cannot be null");
        }

        /// <summary>
        /// Creates an image with a title on on given heading level.
        /// </summary>
        /// <param name="filePath">The path to the image to show in the document</param>
        /// <param name="title">The title of the image. If this is null or empty then no title will be shown.</param>
        /// <param name="paragraphLevel">The level of the paragraph</param>
        public Picture(string filePath, string title, int paragraphLevel)
            : base(title, paragraphLevel)
        {
            _filePath = filePath;
            using (FileStream fs = new FileStream(_filePath, FileMode.Open, FileAccess.Read))
            {
                _image = Image.FromStream(fs);
            }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Gets the paragraph in OOXMLFormat
        /// This method cannot be used for images as it requires a reference to the MainDocument part.
        /// </summary>
        /// <returns>Paragraph in OOXML</returns>
        public override List<OOXMLParagraph> GetOOXMLParagraph()
        {
            throw new NotImplementedException("Cannot use this method");

        }
        /// <summary>
        /// Gets the picture in OOXMLFormat
        /// </summary>
        /// <param name="mainPart">The main part of the OOXML Document</param>
        /// <param name="index">The index of the image in the list of pargraphs.
        /// This is used for generating an identifier for the image</param>
        /// <param name="imageCount">The number of images in the document.
        ///  This is needed for the caption number</param>
        /// <returns>Picture in OOXML Document</returns>
        public List<OOXMLParagraph> GetOOXMLParagraph(OpenXmlPartContainer mainPart, int index, ref int imageCount)
        {
            var imageName = string.Empty;
            var rid = CommonDocumentFunctions.AddPictureToOOXMLDocument(mainPart, _image, index, out imageName);

            var paragraph =
                           new OOXMLParagraph(
                               new Run(
                                   new RunProperties(
                                       new NoProof()),
                                   new Drawing(
                                       new wp.Inline(
                                           new wp.Extent { Cx = 5274945L, Cy = 3956050L },
                                           new wp.EffectExtent { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 1905L, BottomEdge = 0L },
                                           new wp.DocProperties { Id = (UInt32Value)1U, Name = imageName, Description = _text },
                                           new wp.NonVisualGraphicFrameDrawingProperties(
                                               new a.GraphicFrameLocks { NoChangeAspect = true }),
                                           new a.Graphic(
                                               new a.GraphicData(
                                                   new pic.Picture(
                                                       new pic.NonVisualPictureProperties(
                                                           new pic.NonVisualDrawingProperties { Id = (UInt32Value)0U, Name = _text },
                                                           new pic.NonVisualPictureDrawingProperties()),
                                                       new pic.BlipFill(
                                                           new a.Blip { Embed = rid, CompressionState = a.BlipCompressionValues.Print },
                                                           new a.Stretch(
                                                               new a.FillRectangle())),
                                                       new pic.ShapeProperties(
                                                           new a.Transform2D(
                                                               new a.Offset { X = 0L, Y = 0L },
                                                               new a.Extents { Cx = 5274945L, Cy = 3956050L }),
                                                           new a.PresetGeometry(
                                                               new a.AdjustValueList()
                                                           ) { Preset = a.ShapeTypeValues.Rectangle }))
                                               ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                                       ) { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U }))
                           );

            var paragraphs = new List<DocumentFormat.OpenXml.Wordprocessing.Paragraph> { paragraph };

            //Add the caption if there is one.
            if(!string.IsNullOrEmpty(_text))
            {
                paragraphs.Add(captionParagraph(_text, ref imageCount));
            }
            return paragraphs;

        }

        /// <summary>
        /// Gets the picture in OOXMLFormat
        /// </summary>
        /// <param name="doc">The document to add the image to</param>
        /// <param name="index">The index of the image in the list of pargraphs.
        /// This is used for generating an identifier for the image</param>
        /// <param name="imageCount">The number of images in the document.
        ///  This is needed for the caption number</param>
        /// <returns>Picture in OOXML Document</returns>
        public List<AODL.Document.Content.IContent> GetODFParagraph(AODL.Document.TextDocuments.TextDocument doc, int index, ref int imageCount)
        {   
            //Create a string with the filename to use.
            //This is the given filename or a temp file if there is only an image.
            string tempFileName = _filePath;
            if (string.IsNullOrEmpty(tempFileName))
            {
                //If there is only an image and not a filepath create a tempfile.
                string dir = string.Format("{0}\\DocumentGenerator", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData));
                if(!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                tempFileName = string.Format("{0}\\Image{1}.jpg", dir, index);
                _image.Save(tempFileName); 
            }

            //Create the main paragraph.
            AODL.Document.Content.Text.Paragraph p = ParagraphBuilder.CreateStandardTextParagraph(doc);
            if (!string.IsNullOrEmpty(_text))
            {
                //Draw the text box for the label
                var drawTextBox = new DrawTextBox(doc);
                var frameTextBox = new AODL.Document.Content.Draw.Frame(doc, "fr_txt_box") {DrawName = "fr_txt_box", ZIndex = "0"};
                //Create the paragraph for the image
                var pInner = ParagraphBuilder.CreateStandardTextParagraph(doc);
                pInner.StyleName = "Illustration";
                //Create the image frame
                var frame = new AODL.Document.Content.Draw.Frame(doc, "frame1", string.Format("graphic{0}", index), tempFileName) {ZIndex = "1"};
                //Add the frame.
                pInner.Content.Add(frame);
                //Add the title
                pInner.TextContent.Add(new SimpleText(doc, _text));
                //Draw the text and image in the box.
                drawTextBox.Content.Add(pInner);
                frameTextBox.SvgWidth = frame.SvgWidth;
                drawTextBox.MinWidth = frame.SvgWidth;
                drawTextBox.MinHeight = frame.SvgHeight;
                frameTextBox.Content.Add(drawTextBox);
                p.Content.Add(frameTextBox);
            }
            else
            {
                //Create a frame with the image.
                var frame = new AODL.Document.Content.Draw.Frame(doc, "Frame1", string.Format("graphic{0}", index), tempFileName);
                p.Content.Add(frame);
            }

            //Finally delete image file if it is a temp file.
            //if (tempFileName != _filePath) File.Delete(tempFileName);

            return new List<AODL.Document.Content.IContent> { p };
        }
        #endregion

        #region Private Methods
        private OOXMLParagraph captionParagraph(string text, ref int imageCount)
        {
            //Place a space before the text, otherwise it does not show correct.
            text = string.Format(" {0}", text);
             var paragraph =
                new OOXMLParagraph(
                    new ParagraphProperties(
                        new ParagraphStyleId{ Val = "Caption" }),
                    new Run(
                        new Text("Figure ") { Space = SpaceProcessingModeValues.Preserve }),
                    new SimpleField(
                        new Run(
                            new RunProperties(
                                new NoProof()),
                            new Text(imageCount.ToString(CultureInfo.CurrentCulture)))
                    ){ Instruction = " SEQ Figure \\* ARABIC " },
                    new Run(
                        new Text(text) { Space = SpaceProcessingModeValues.Preserve })
                );

             imageCount++;
            return paragraph;

        }
        #endregion
    }
}

