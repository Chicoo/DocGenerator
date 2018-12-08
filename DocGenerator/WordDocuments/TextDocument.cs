using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing;
using DocumentFormat.OpenXml;
using DocumentGenerator.Common;
using DocumentGenerator.WordDocuments.StyleFormatting;

namespace DocumentGenerator.WordDocuments
{
    /// <summary>
    /// Creates/Modifies a text document.
    /// </summary>
    public class TextDocument
    {
        private string _filenameTemplate;
        private Stream _template;
        private bool _isNew; //Boolean to indicate if it is a new file or an existing file.
        private readonly List<CustomPart> _customParts;

        /// <summary>
        /// Gets the list of tables in the document.
        /// </summary>
        public IList<Table> Tables
        {
            get 
            {
                List<Table> tables =Paragraphs.Values.OfType<Table>().ToList();
                return tables.AsReadOnly();
            }
        }
        /// <summary>
        /// Gets a <see cref="SortedDictionary<int, Paragraph>"/> of the paragraphs of the document.
        /// </summary>
        public SortedDictionary<int, Paragraph> Paragraphs { get; }

        /// <summary>
        /// Gets a <see cref="IEnumerable<CustomPart>"/> of the <see cref="CustomPart"/>s of the document.
        /// </summary>
        public IEnumerable<CustomPart> CustomParts
        {
            get { return _customParts.AsReadOnly(); }
        }

        /// <summary>
        /// Gets or sets the filename of the document.
        /// </summary>
        public string Filename { get; set; }

        /// <summary>
        /// Creates an instance of a an OOXMLWordDocument
        /// This is an internal method as the document needs to be creates through the Open method 
        /// </summary>
        internal TextDocument()
        {
            Paragraphs =new SortedDictionary<int,Paragraph>();
            _customParts = new List<CustomPart>();
        }
        
        /// <summary>
        /// Adds a new <see cref="Paragraph"/>  to the document with a header and text and on a given heading level.
        /// </summary>
        /// <param name="paragraphHeader">The header of the paragrah</param>
        /// <param name="text">The text of the paragraph</param>
        /// <param name="headingLevel">The level of the paragraph</param>
        /// <returns>The paragraph without the header</returns>
        public Paragraph AddParagraph(string paragraphHeader, string text, int headingLevel)
        {
            //Get the new index
            var index = GetNewIndex();
            //Now add the paragraph
            var paragraph = new Paragraph(text, paragraphHeader, headingLevel);
            Paragraphs.Add(index, paragraph);

            return paragraph;
        }

        /// <summary>
        /// Adds a new paragrah to the document on a given heading level.
        /// </summary>
        /// <param name="text">The text of the paragraph</param>
        /// <param name="headingLevel">The level of the paragraph</param>
        /// <returns>The paragraph without the header</returns>
        public Paragraph AddParagraph(string text, int headingLevel)
        {
            //Get the new index
            var index = GetNewIndex();
            //Now add the paragraph
            var paragraph = new Paragraph(text, headingLevel);
            Paragraphs.Add(index, paragraph);

            return paragraph;
        }

        /// <summary>
        /// Adds a new paragrah with a header and text with a given style for the header and the paragraph.
        /// </summary>
        /// <param name="headerText">The header of the paragrah</param>
        /// <param name="text">The text of the paragraph</param>
        /// <param name="paragraphStyle">The style of the paragraph</param>
        /// <param name="headerStyle">The style of the header</param>
        public Paragraph AddParagraph(string headerText, string text, string paragraphStyle, string headerStyle)
        {
            //Get the new index
            var index = GetNewIndex();
            //Now add the paragraph
            var paragraph = new Paragraph(headerText, text, paragraphStyle, headerStyle);
            Paragraphs.Add(index, paragraph);

            return paragraph;
        }

        /// <summary>
        /// Adds a new text paragrah with a given style, without a header.
        /// </summary>
        /// <param name="text">The text of the paragraph.</param>
        /// <param name="paragraphStyle">The style of the paragraph.</param>
        public Paragraph AddParagraph(string text, string paragraphStyle)
        {
            //Get the new index
            var index = GetNewIndex();
            //Now add the paragraph
            var paragraph = new Paragraph(text, paragraphStyle);
            Paragraphs.Add(index, paragraph);

            return paragraph;
        }

        /// <summary>
        /// Adds a new paragrah, with an id, to the document with a header and text and on a given heading level.
        /// </summary>
        /// <param name="paragraphHeader">The header of the paragrah</param>
        /// <param name="text">The text of the paragraph</param>
        /// <param name="headingLevel">The level of the paragraph</param>
        /// <param name="id">The id (name) of the paragraph</param>
        /// <returns>The paragraph without the header</returns>
        public NamedParagraph AddNamedParagraph(string paragraphHeader, string text, int headingLevel, string id)
        {
            //Get the new index
            var index = GetNewIndex();
            //Now add the paragraph
            var paragraph = new NamedParagraph(text, paragraphHeader, headingLevel, id);
            Paragraphs.Add(index, paragraph);

            return paragraph;
        }

        /// <summary>
        /// Adds a new paragrah, with an id, to the document on a given heading level.
        /// </summary>
        /// <param name="text">The text of the paragraph</param>
        /// <param name="headingLevel">The level of the paragraph</param>
        /// <param name="id">The id (name) of the paragraph</param>
        /// <returns>The paragraph without the header</returns>
        public NamedParagraph AddNamedParagraph(string text, int headingLevel, string id)
        {
            //Get the new index
            var index = GetNewIndex();
            //Now add the paragraph
            var paragraph = new NamedParagraph(text, headingLevel, id);
            Paragraphs.Add(index, paragraph);

            return paragraph;
        }

        /// <summary>
        /// Adds a new paragrah with a header and text with a give style for the header and the paragraph.
        /// </summary>
        /// <param name="text">The text of the paragraph</param>
        /// <param name="paragraphStyle">The style of the paragraph</param>
        /// <param name="id">The id (name) of the paragraph</param>
        public NamedParagraph AddNamedParagraph(string text, string paragraphStyle, string id)
        {
            //Get the new index
            var index = GetNewIndex();
            //Now add the paragraph
            var paragraph = new NamedParagraph(text, paragraphStyle, id);
            Paragraphs.Add(index, paragraph);

            return paragraph;
        }

        /// <summary>
        /// Adds a new paragrah with a header and text with a give style for the header and the paragraph.
        /// </summary>
        /// <param name="headerText">The header of the paragrah</param>
        /// <param name="text">The text of the paragraph</param>
        /// <param name="paragraphStyle">The style of the paragraph</param>
        /// <param name="headerStyle">The style of the header</param>
        /// <param name="id">The id (name) of the paragraph</param>
        public NamedParagraph AddNamedParagraph(string text, string headerText, string paragraphStyle, string headerStyle, string id)
        {
            //Get the new index
            var index = GetNewIndex();
            //Now add the paragraph
            var paragraph = new NamedParagraph(text, headerText, paragraphStyle, headerStyle, id);
            Paragraphs.Add(index, paragraph);

            return paragraph;
        }

        /// <summary>
        /// Adds a picture to the document
        /// </summary>
        /// <param name="image">The image to add</param>
        /// <param name="title">The title of the image</param>
        /// <returns>An picture object</returns>
        public Picture AddPicture(Image image, string title)
        {
            //Get the new index
            int index = GetNewIndex();
            //Add the picture
            Picture picture = new Picture(image, title, 0);
            Paragraphs.Add(index, picture);

            return picture;
        }

        /// <summary>
        /// Adds a picture to the document
        /// </summary>
        /// <param name="filePath">The path to the image to add</param>
        /// <param name="title">The title of the image</param>
        /// <returns>An picture object</returns>
        public Picture AddPicture(string filePath, string title)
        {
            //Get the new index
            int index = GetNewIndex();
            //Add the picture
            Picture picture = new Picture(filePath, title, 0);
            Paragraphs.Add(index, picture);

            return picture;
        } 
        
        /// <summary>
        /// Adds a table to the document
        /// </summary>
        /// <param name="title">The title of the table</param>
        /// <returns>A table</returns>
        public Table AddTable(string title)
        {
            //Get the new index
            int index = GetNewIndex();
            //Add the table
            Table table = new Table(title, 0);
            Paragraphs.Add(index, table);

            return table;
        }

        /// <summary>
        /// Creates a bullet list for the document.
        /// </summary>
        /// <returns></returns>
        public BulletList AddBulletList()
        {
            //Get the new index
            int index = GetNewIndex();
            //Add the list
            BulletList list = new BulletList();
            Paragraphs.Add(index, list);

            return list;
        }

        /// <summary>
        /// Creates a numbered list for the document.
        /// </summary>
        /// <returns></returns>
        public NumberedList AddNumberedList()
        {
            //Get the new index
            int index = GetNewIndex();
            //Add the list
            NumberedList list = new NumberedList();
            Paragraphs.Add(index, list);

            return list;
        } 

        /// <summary>
        /// Replaces a custom tag in a Word document with the provided text.
        /// </summary>
        /// <param name="name">The name of the tag to replace the value for.</param>
        /// <param name="text">The new text.</param>
        public void ReplaceCustomTag(string name, string text)
        {
            if (_customParts.Any(c => c.Name == name))
            {
                throw new DuplicateCustomTagException("The name already exists in the custom parts list.");
            }
            var newPart = new CustomPart(name, text);
            _customParts.Add(newPart);
        }

        /// <summary>
        /// Replaces a custom tag in a Word document with the provided image.
        /// </summary>
        /// <param name="name">The name of the tag to replace the value for.</param>
        /// <param name="image">The new image.</param>
        public void ReplaceCustomTag(string name, Image image)
        {
            if (_customParts.Any(c => c.Name == name))
            {
                throw new DuplicateCustomTagException("The name already exists in the custom parts list.");
            }
            var newPart = new CustomPart(name, image);
            _customParts.Add(newPart);
        }

        /// <summary>
        /// Saves the document in a given type
        /// </summary>
        /// <param name="documentType">The type to save the document int</param>
        /// <exception cref="FilenameEmptyException">Thrown when the filename is null or empty.</exception>
        public void Save(Common.DocumentType documentType)
        {
            if (string.IsNullOrWhiteSpace(Filename))
            {
                throw new FilenameEmptyException("The filename for the document is not provided.");
            }
            switch (documentType)
            {
                case Common.DocumentType.OOXMLTextDocument:
                    SaveInOOXML();
                    break;
                case Common.DocumentType.ODFTextDocument:
                    SaveInODF();
                    break;
            }
        } 

        /// <summary>
        /// Saves the document in OOXML format.
        /// </summary>
        private void SaveInOOXML()
        {
            //First create the file if this is a new file
            if (_isNew && _template != null)
            {
                //Create a new file based on the template
                //Read the template stream inot a new document.
                using (_template)
                {
                    using (var documentStream = new MemoryStream((int)_template.Length))
                    {
                        _template.Position = 0L;
                        byte[] buffer = new byte[2048];
                        int length = buffer.Length;
                        int size;
                        while ((size = _template.Read(buffer, 0, length)) != 0)
                        {
                            documentStream.Write(buffer, 0, size);
                        }
                        documentStream.Position = 0L;
                        // Modify the document to ensure it is correctly marked as a Document and not Template    
                        using (var document = WordprocessingDocument.Open(documentStream, true))
                        {
                            document.ChangeDocumentType(WordprocessingDocumentType.Document);
                        }
                        //Save the stream to a new file.
                        File.WriteAllBytes(Filename, documentStream.ToArray());
                    }
                }
            }
            else if (_isNew && !string.IsNullOrWhiteSpace(_filenameTemplate))
            {
                //Create a new file based on the template
                //Read the file in to a stream.
                using (var templateStream = File.OpenRead(_filenameTemplate))
                {
                    using (var documentStream = new MemoryStream((int)templateStream.Length))
                    {
                        templateStream.Position = 0L;
                        byte[] buffer = new byte[2048];
                        int length = buffer.Length;
                        int size;
                        while ((size = templateStream.Read(buffer, 0, length)) != 0)
                        {
                            documentStream.Write(buffer, 0, size);
                        }
                        documentStream.Position = 0L;
                        // Modify the document to ensure it is correctly marked as a Document and not Template    
                        using (var document = WordprocessingDocument.Open(documentStream, true))
                        {
                            document.ChangeDocumentType(WordprocessingDocumentType.Document);
                        }
                        //Save the stream to a new file.
                        File.WriteAllBytes(Filename, documentStream.ToArray());
                    }
                }
            }
            else if (_isNew)
            {
                using (var doc = WordprocessingDocument.Create(Filename, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    var document = new Document(
                        new Body());
                    document.Save(mainPart);

                    //Create the styles part.
                    var styleDefinitionsPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    StyleGenerator.DefaultStyle().Save(styleDefinitionsPart);
                    
                    //Close the document
                    doc.Close();
                }
            }

            using (var doc = WordprocessingDocument.Open(Filename, true))
            {
                try
                {
                    var mainPart = doc.MainDocumentPart;
                    /* Create the contents */
                    var body = mainPart.Document.Body;
                    var imageCount = 1; //This is a counter for setting the image number.
                    var listCount = 0; //This is the counter for the number of lists.
                    //Create the list templates
                    var numberingDefinitionsPart = mainPart.NumberingDefinitionsPart??mainPart.AddNewPart<NumberingDefinitionsPart>();
                    var numbering = ListGenerator.CreateNumbering(numberingDefinitionsPart);

                    foreach (var paragraph in Paragraphs)
                    {
                        //If this is an image we also need to pass the MainpArt as this is needed to generate the image part
                        var img = paragraph.Value as Picture;
                        var table = paragraph.Value as Table;
                        var list = paragraph.Value as List;
                        if (img != null)
                        {
                            foreach (var p in img.GetOOXMLParagraph(mainPart, paragraph.Key, ref imageCount))
                            {
                                body.Append(p);
                            }
                           
                        }
                        else if (table != null)
                        {
                            //Add the table
                            body.Append(table.GetOOXMLTable());
                            //Add an extra paragraph so two following tables will not be concatenated.
                            body.Append(Paragraph.EmptyParagraph());
                        }
                        else if (list != null)
                        {
                            //Add the list
                            listCount++;
                            //Add the reference to the list
                            NumberingInstance numberingInstance = new NumberingInstance { NumberID = listCount };
                            AbstractNumId abstractNumId = new AbstractNumId { Val = list.Type == ListType.Bullet ? 0 : 1 };

                            numberingInstance.Append(abstractNumId);
                            numbering.Append(numberingInstance);

                            foreach (var item in list.GetOOXMLList(listCount))
                            {
                                body.Append(item);
                            }
                        }
                        else
                        {
                            //Add the text paragraph(s)
                            foreach (var p in paragraph.Value.GetOOXMLParagraph())
                            {
                                body.Append(p);
                            }
                        }
                    }

                    //All the paragraphs are done.Save the numbering.
                    numbering.Save(numberingDefinitionsPart);

                    //Replace the custom tags.
                    foreach(var customPart in _customParts)
                    {
                        customPart.Replace(mainPart, ref imageCount);
                    }

                    /* Save the results and close */
                    mainPart.Document.Save();
                    doc.Close();
                }
                catch (Exception ex)
                {
                    throw new FileLoadException("An error occured when saving the document", Filename, ex);
                }
            }
        }
        /// <summary>
        /// Gets a new index for the paragraphs. This is the last index +1.
        /// </summary>
        /// <returns></returns>
        private int GetNewIndex()
        {
            int index = 0;
            //Get the last index if there are entries in the dictionairy.
            if(Paragraphs.Count !=0)  
            {
                index = Paragraphs.Max(p => p.Key);
                index++;
            }
            //Add one for the new index.
            return index;
        }

        private void SaveInODF()
        {
            using (AODL.Document.TextDocuments.TextDocument doc = new AODL.Document.TextDocuments.TextDocument())
            {
                if (_isNew)
                {
                    doc.New();
                }
                else
                {
                    doc.Load(Filename);
                }

                try
                {
                    int imageCount = 1; //This is a counter for setting the image number.
                    foreach (KeyValuePair<int, Paragraph> paragraph in Paragraphs)
                    {
                        //If this is an image we also need to pass the MainpArt as this is needed to generate the image part
                        Picture img = paragraph.Value as Picture;
                        Table table = paragraph.Value as Table;
                        List list = paragraph.Value as List;
                        if (img != null)
                        {
                            foreach (var p in img.GetODFParagraph(doc, paragraph.Key, ref imageCount))
                            {
                                doc.Content.Add(p);
                            }

                        }
                        else if (table != null)
                        {
                            //Add table to the document
                            doc.Content.Add(table.GetODFTable(doc));
                        }
                        else if (list != null)
                        {
                            //Add the list.
                            doc.Content.Add(list.GetODFList(doc));
                        }
                        else
                        {
                            //Add the text paragraph(s)
                            foreach (var p in paragraph.Value.GetODFParagraph(doc))
                            {
                                doc.Content.Add(p);
                            }

                        }
                    }

                    /* Save the results and close */
                    doc.SaveTo(Filename);
                }
                catch (Exception ex)
                {
                    throw new FileLoadException("An error occured when saving the document: " + ex.Message, Filename, ex);
                }
            }
        }

        /// <summary>
        /// Copies the stream that is supplied into a local stream..
        /// </summary>
        /// <param name="source">The source stream</param>
        /// <param name="destination">The destination</param>
        static private void CopyStream(Stream source, Stream destination)
        {
            byte[] buffer = new byte[32768];
            int bytesRead;
            do
            {
                bytesRead = source.Read(buffer, 0, buffer.Length);
                destination.Write(buffer, 0, bytesRead);
            } while (bytesRead != 0);
        }

        /// <summary>
        /// Opens a OOXML word document based on a filename
        /// </summary>
        /// <param name="fileName"></param>
        public static TextDocument Open(string fileName)
        {
            if (!File.Exists(fileName))
                throw new FileNotFoundException("The document could not be found", fileName);

            var oOXMLWordDocument = new TextDocument
                                        {
                                            Filename=fileName, _isNew=false
                                        };

            return oOXMLWordDocument;
        }

        /// <summary>
        /// Creates new document based on a filename
        /// </summary>
        /// <param name="fileName">The name of the document</param>
        /// <param name="replace">Indicates if an existing file can be replaced</param>
        public static TextDocument Create(string fileName, bool replace)
        {
            if (File.Exists(fileName) && !replace)
                throw new FileNotFoundException("The document already exists", fileName);

            var oOXMLWordDocument = new TextDocument
                                        {
                                            Filename=fileName, _isNew=true
                                        };

            return oOXMLWordDocument;
        }

        /// <summary>
        /// Creates new document
        /// </summary>
        public static TextDocument Create()
        {
            var oOXMLWordDocument = new TextDocument
                                        {
                                            Filename=string.Empty, _isNew=true
                                        };

            return oOXMLWordDocument;
        }

        /// <summary>
        /// Creates a new document with the given filename based on a template.
        /// </summary>
        /// <param name="filename">The name of the document</param>
        /// <param name="template">The filename of the template to base the document on.</param>
        /// <param name="replace">Indicates if an existing file can be replaced</param>
        /// <returns>A new text document</returns>
        public static TextDocument Create(string filename, string template, bool replace)
        {
            if (File.Exists(filename) && !replace)
                throw new FileNotFoundException("The document already exists", filename);

            var oOXMLWordDocument = new TextDocument
            {
                Filename = filename,
                _filenameTemplate = template,
                _isNew = true
            };

            return oOXMLWordDocument;

        }

        /// <summary>
        /// Creates a new document based on a template.
        /// </summary>
        /// <param name="template">A stream containing the template.</param>
        /// <returns>A new text document</returns>
        public static TextDocument Create(Stream template)
        {
            
            var oOXMLWordDocument = new TextDocument
                                        {
                                            _isNew=true, _template=new MemoryStream(),
                                        };
            CopyStream(template, oOXMLWordDocument._template);

            return oOXMLWordDocument;
        }
    }
}
