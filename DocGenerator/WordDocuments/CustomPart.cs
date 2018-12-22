using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing;

namespace DocumentGenerator.WordDocuments
{
    /// <summary>
    /// Represents a custom part, that will replace the contents of a Word custom part
    /// with the contents defined in here.
    /// </summary>
    public class CustomPart
    {
        #region Fields
        //Booleans to indicate the type of part that is being replaced.
        //This reduces the casting needed to determine the type of object.
        private readonly bool _isParagraph;
        private readonly bool _isImage;
        #endregion Fields

        #region Properties
        /// <summary>
        /// The name of the custom property to change in the Word document
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// The value of the custom property
        /// </summary>
        public object Value { get; set; }
        #endregion Properties

        #region Constructors
        /// <summary>
        /// Creates a new custom part object for an image
        /// </summary>
        /// <param name="name">The name of the custom part to replace the value for.</param>
        /// <param name="image">The new image for the part.</param>
        internal CustomPart(string name, Image image) : this(name, (object)image)
        {
            _isImage = true;
        }

        /// <summary>
        /// Creates a new custom part object for a paragraph.
        /// </summary>
        /// <param name="name">The name of the custom part to replace the value for.</param>
        /// <param name="text">The new text for the part.</param>
        internal CustomPart(string name, string text) : this(name, (object)text)
        {
            _isParagraph = true;
        }

        /// <summary>
        /// Creates a new custom part object
        /// </summary>
        /// <param name="name">The name of the custom part to replace the value for.</param>
        /// <param name="value">The new value of the part.</param>
        private CustomPart(string name, object value)
        {
            Name = name;
            Value = value;
        }

        #endregion Constructors

        #region Public methods

        /// <summary>
        /// Replaces the value of a custom part in the word document.
        /// It replaces tags in the main body, headers and footers.
        /// </summary>
        /// <param name="mainPart">The document to search in.</param>
        /// <param name="imageCount">The number of images</param>
        public void Replace(MainDocumentPart mainPart, ref int imageCount)
        {
            //Replace the text if this is a paragraph type.
            if (_isParagraph) ReplaceText(mainPart);
            if (_isImage) ReplaceImage(mainPart, ref imageCount);
        }

        #endregion Public methods

        #region Private methods
        /// <summary>
        /// Replaces the text parts in the documents.
        /// </summary>
        /// <param name="mainPart">The document to search in.</param>
        private void ReplaceText(MainDocumentPart mainPart)
        {
            if (!(Value is string)) return;
            //Find the tags in the body of the document. The tag value must correspond to the name provided in this object.
            var sdtList = mainPart.Document.Descendants<SdtElement>().Where(s => s.SdtProperties.GetFirstChild<Tag>().Val.Value == Name).ToList();
            //Search the header part for the custom tag.
            foreach (var headerPart in mainPart.HeaderParts)
            {
                sdtList.AddRange(headerPart.Header.Descendants<SdtElement>().Where(s => s.SdtProperties.GetFirstChild<Tag>().Val.Value == Name).ToList());
            }
            //Search the footer for the custom tag.
            foreach (var footerPart in mainPart.FooterParts)
            {
                sdtList.AddRange(footerPart.Footer.Descendants<SdtElement>().Where(s => s.SdtProperties.GetFirstChild<Tag>().Val.Value == Name).ToList());
            }
            //Loop through the found items and replace the text. This will keep the formatting as is in the document.
            if (sdtList.Count <= 0) return;
            foreach (var run in sdtList.Select(block => block.Descendants<Run>().ToList()).SelectMany(runs => runs))
            {
                run.InnerXml = GenerateRunBlock().InnerXml;
            }
        }

        /// <summary>
        /// Replaces the text parts in the documents.
        /// </summary>
        /// <param name="mainPart">The document to search in.</param>
        /// <param name="imageCount">The number of images</param>
        private void ReplaceImage(MainDocumentPart mainPart, ref int imageCount)
        {
            if (!(Value is Image image)) return;

            //Find the tags in the body of the document. The tag value must correspond to the name provided in this object.
            //For images this is the SdtCell object.
            var sdtList = mainPart.Document.Descendants<SdtElement>().Where(s => s.SdtProperties.GetFirstChild<Tag>().Val.Value == Name).ToList();
            //Search the header part for the custom tag.
            foreach (var headerPart in mainPart.HeaderParts)
            {
                foreach(var element in  headerPart.Header.Descendants<SdtElement>().Where(s => s.SdtProperties.GetFirstChild<Tag>().Val.Value == Name).ToList())
                {
                    foreach (var run in element.Descendants<Run>().ToList().SelectMany(runs => runs))
                    {
                        //Find the image.
                        var images = run.Descendants<Drawing>().ToList();
                        if (images.Count != 1) continue;
                        var drawing = images[0];
                        //When the image is found add the new image to the document.
                        imageCount++;
                        var rid = CommonDocumentFunctions.AddPictureToOOXMLDocument(headerPart, image, imageCount, out string imageName);
                        //Replace the old image name and id wityh the new one.
                        var properties = drawing.Inline.DocProperties;
                        properties.Name = imageName;
                        //Find the picture object.
                        var picture = drawing.Inline.Graphic.GraphicData.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().ToList();
                        if (picture.Count != 1) continue;
                        picture[0].BlipFill.Blip.Embed = rid;
                    }
                }
                //sdtList.AddRange(headerPart.Header.Descendants<SdtElement>().Where(s => s.SdtProperties.GetFirstChild<Tag>().Val.Value == Name).ToList());
            }
            //Search the footer for the custom tag.
            foreach (var footerPart in mainPart.FooterParts)
            {
                sdtList.AddRange(footerPart.Footer.Descendants<SdtElement>().Where(s => s.SdtProperties.GetFirstChild<Tag>().Val.Value == Name).ToList());
            }
            //Loop through the found items and replace the image. This will keep the formatting as is in the document.
            if (sdtList.Count <= 0) return;
            foreach (var run in sdtList.Select(block => block.Descendants<Run>().ToList()).SelectMany(runs => runs))
            {
                //Find the image.
                var images = run.Descendants<Drawing>().ToList();
                if (images.Count != 1) continue;
                var drawing = images[0];
                //When the image is found add the new image to the document.
                var rid = CommonDocumentFunctions.AddPictureToOOXMLDocument(mainPart, image, 0, out string imageName);
                //Replace the old image name and id wityh the new one.
                var properties = drawing.Inline.DocProperties;
                properties.Name = imageName;
                //Find the picture object.
                var picture = drawing.Inline.Graphic.GraphicData.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().ToList();
                if (picture.Count != 1) continue;
                picture[0].BlipFill.Blip.Embed = rid;
            }
        }

        /// <summary>
        /// Creates an Run to replace the value in the custom tag.
        /// </summary>
        /// <returns></returns>
        private Run GenerateRunBlock()
        {
            Run run=new Run();
            var stringValue=Value as string;
            if(!string.IsNullOrWhiteSpace(stringValue))
            {
                run.Append(CommonDocumentFunctions.ParseParagraphForOOXML(stringValue));
            }
            return run;
        }
        #endregion Private methods
    }
}
