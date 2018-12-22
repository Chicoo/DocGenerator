/*************************************************************************
 *
 * DO NOT ALTER OR REMOVE COPYRIGHT NOTICES OR THIS FILE HEADER
 * 
 * Copyright 2008 Sun Microsystems, Inc. All rights reserved.
 * 
 * Use is subject to license terms.
 * 
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not
 * use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0. You can also
 * obtain a copy of the License at http://odftoolkit.org/docs/license.txt
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * 
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 ************************************************************************/

using System;
using System.Collections;
using System.IO;
using System.Xml;
using System.Reflection;
using AODL.Document.TextDocuments;
using AODL.Document.Content.Tables;
using AODL.Document;
using AODL.Document.Export;
using AODL.Document.Import;
using AODL.Document.Import.OpenDocument;
using AODL.Document.Import.OpenDocument.NodeProcessors;
using AODL.Document.Styles;
using AODL.Document.Content;
using AODL.Document.Exceptions;
using AODL.Document .Content .EmbedObjects  ;

namespace AODL.Document.SpreadsheetDocuments
{
	/// <summary>
	/// The SpreadsheetDocument class represent an OpenDocument
	/// spreadsheet document.
	/// </summary>
	public class SpreadsheetDocument : IDocument, IDisposable
	{
		private IImporter m_importer = null;
		
		private readonly StyleFactory m_styleFactory = null;
		
		private int _tableCount		= 0;
		/// <summary>
		/// Gets the table count.
		/// </summary>
		/// <value>The table count.</value>
		public int TableCount
		{
			get { return _tableCount;  }
		}

		public StyleFactory StyleFactory {
			get { return m_styleFactory; }
		}
		
		private DirInfo m_dirInfo = null;
		
		public DirInfo DirInfo {
			get { return m_dirInfo; }
		}
		
		
		private XmlDocument _xmldoc;
		/// <summary>
		/// The xml document the spreadsheet document based on.
		/// </summary>
		public XmlDocument XmlDoc
		{
			get { return _xmldoc; }
			set { _xmldoc = value; }
		}

		private bool _isLoadedFile;
		/// <summary>
		/// If this file was loaded
		/// </summary>
		/// <value></value>
		public bool IsLoadedFile
		{
			get { return _isLoadedFile; }
		}

		private ArrayList _fontList;
		/// <summary>
		/// The font list
		/// </summary>
		/// <value></value>
		public ArrayList FontList
		{
			get { return _fontList; }
			set { _fontList = value; }
		}

		private StyleCollection _styles;
		/// <summary>
		/// Collection of local styles used with this document.
		/// </summary>
		/// <value></value>
		public StyleCollection Styles
		{
			get { return _styles; }
			set { _styles = value; }
		}

		private StyleCollection _commonStyles;
		/// <summary>
		/// Collection of common styles used with this document.
		/// </summary>
		/// <value></value>
		public StyleCollection CommonStyles
		{
			get { return _commonStyles; }
			set { _commonStyles = value; }
		}

		private ContentCollection _contents;
		/// <summary>
		/// Collection of contents used by this document.
		/// </summary>
		/// <value></value>
		public ContentCollection Content
		{
			get { return _contents; }
			set { _contents = value; }
		}

		private TableCollection _tableCollection;
		/// <summary>
		/// Gets or sets the table collection.
		/// </summary>
		/// <value>The table collection.</value>
		public TableCollection TableCollection
		{
			get { return _tableCollection; }
			set { _tableCollection = value; }
		}

		private DocumentStyles _documentStyles;
		/// <summary>
		/// Gets or sets the document styes.
		/// </summary>
		/// <value>The document styes.</value>
		public DocumentStyles DocumentStyles
		{
			get { return _documentStyles; }
			set { _documentStyles = value; }
		}

		private DocumentSetting _documentSetting;
		/// <summary>
		/// Gets or sets the document setting.
		/// </summary>
		/// <value>The document setting.</value>
		public DocumentSetting DocumentSetting
		{
			get { return _documentSetting; }
			set { _documentSetting = value; }
		}

		private DocumentMetadata _documentMetadata;
		/// <summary>
		/// Gets or sets the document metadata.
		/// </summary>
		/// <value>The document metadata.</value>
		public DocumentMetadata DocumentMetadata
		{
			get { return _documentMetadata; }
			set { _documentMetadata = value; }
		}

		private DocumentConfiguration2 _documentConfiguations2;
		/// <summary>
		/// Gets or sets the document configuration2.
		/// </summary>
		/// <value>The document configuration2.</value>
		public DocumentConfiguration2 DocumentConfigurations2
		{
			get { return _documentConfiguations2; }
			set { _documentConfiguations2 = value; }
		}

		private DocumentManifest _documentManifest;
		/// <summary>
		/// Gets or sets the document manifest.
		/// </summary>
		/// <value>The document manifest.</value>
		public DocumentManifest DocumentManifest
		{
			get { return _documentManifest; }
			set { _documentManifest = value; }
		}

		private DocumentPictureCollection _documentPictures;
		/// <summary>
		/// Gets or sets the document pictures.
		/// </summary>
		/// <value>The document pictures.</value>
		public DocumentPictureCollection DocumentPictures
		{
			get { return _documentPictures; }
			set { _documentPictures = value; }
		}

		private DocumentPictureCollection _documentThumbnails;
		/// <summary>
		/// Gets or sets the document thumbnails.
		/// </summary>
		/// <value>The document thumbnails.</value>
		public DocumentPictureCollection DocumentThumbnails
		{
			get { return _documentThumbnails; }
			set { _documentThumbnails = value; }
		}

		private readonly ArrayList _graphics;
		/// <summary>
		/// Gets the graphics.
		/// </summary>
		/// <value>The graphics.</value>
		public ArrayList Graphics
		{
			get { return _graphics; }
		}

		private ArrayList _embedobjects;
		/// <summary>
		/// Gets the embedobject.
		/// </summary>
		/// <value>The embedobject.</value>
		public ArrayList EmbedObjects
		{
			get { return _embedobjects; }

			set { _embedobjects =value; }
		}

		private readonly string _mimeTyp		= "application/vnd.oasis.opendocument.spreadsheet";
		/// <summary>
		/// Gets the MIME typ.
		/// </summary>
		/// <value>The MIME typ.</value>
		public string MimeTyp
		{
			get { return _mimeTyp; }
		}

		private XmlNamespaceManager _namespacemanager;
		/// <summary>
		/// The namespacemanager to access the documents namespace uris.
		/// </summary>
		public XmlNamespaceManager NamespaceManager
		{
			get { return _namespacemanager; }
			set { _namespacemanager = value; }
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="SpreadsheetDocument"/> class.
		/// </summary>
		public SpreadsheetDocument()
		{
			TableCollection			= new TableCollection();
			Styles						= new StyleCollection();
			m_styleFactory             = new StyleFactory(this);
			CommonStyles				= new StyleCollection();
			Content					= new ContentCollection();
			_graphics					= new ArrayList();
			FontList					= new ArrayList();
			EmbedObjects               = new ArrayList ();
			TableCollection.Inserted	+= TableCollection_Inserted;
			TableCollection.Removed	+= TableCollection_Removed;
			Content.Inserted			+= Content_Inserted;
			Content.Removed			+= Content_Removed;
		}

		/// <summary>
		/// Create a new blank spreadsheet document.
		/// </summary>
		public void New()
		{
			LoadBlankContent();

			NamespaceManager			= TextDocumentHelper.NameSpace(_xmldoc.NameTable);

			DocumentConfigurations2	= new DocumentConfiguration2();

			DocumentManifest			= new DocumentManifest();
			DocumentManifest.New();

			DocumentMetadata			= new DocumentMetadata(this);
			DocumentMetadata.New();

			DocumentPictures			= new DocumentPictureCollection();

			DocumentSetting			= new DocumentSetting();
			DocumentSetting.New();

			DocumentStyles				= new DocumentStyles();
			DocumentStyles.New();
			ReadCommonStyles();

			DocumentThumbnails			= new DocumentPictureCollection();
			
		}

		public void DeleteUnpackedFiles()
		{
			if (m_importer != null)
				m_importer.DeleteUnpackedFiles();
		}
		
		/// <summary>
		/// Reads the common styles.
		/// </summary>
		private void ReadCommonStyles()
		{
            OpenDocumentImporter odImporter = new OpenDocumentImporter
            {
                Document = this
            };
            odImporter.ImportCommonStyles();

			LocalStyleProcessor lsp			= new LocalStyleProcessor(this, true);
			lsp.ReadStyles();
		}

		/// <summary>
		/// Save the SpreadsheetDocument as OpenDocument spreadsheet document
		/// </summary>
		/// <param name="filename">The filename. With or without full path. Without will save the file to application path!</param>
		public void SaveTo(string filename)
		{
			CreateContentBody();

			//Call the ExportHandler for the first matching IExporter
			ExportHandler exportHandler		= new ExportHandler();
			IExporter iExporter				= exportHandler.GetFirstExporter(
				DocumentTypes.SpreadsheetDocument, filename);
			iExporter.Export(this, filename);
		}

		/// <summary>
		/// Save the document by using the passed IExporter
		/// with the passed file name.
		/// </summary>
		/// <param name="filename">The name of the new file.</param>
		/// <param name="iExporter"></param>
		public void SaveTo(string filename, IExporter iExporter)
		{
			CreateContentBody();
			iExporter.Export(this, filename);
		}

		/// <summary>
		/// Load the given file.
		/// </summary>
		/// <param name="file"></param>
		public void Load(string file)
		{
			_isLoadedFile				= true;
			LoadBlankContent();

			NamespaceManager			= TextDocumentHelper.NameSpace(_xmldoc.NameTable);

			ImportHandler importHandler		= new ImportHandler();
			m_importer				= importHandler.GetFirstImporter(DocumentTypes.SpreadsheetDocument, file);

			m_dirInfo = m_importer.DirInfo;
			if (m_importer != null)
			{
				if (m_importer.NeedNewOpenDocument)
					New();
				m_importer.Import(this,file);

				if (m_importer.ImportError != null)
					if (m_importer.ImportError.Count > 0)
					foreach(object ob in m_importer.ImportError)
					if (ob is AODLWarning)
				{
					if (((AODLWarning)ob).Message != null)
						Console.WriteLine("Err: {0}", ((AODLWarning)ob).Message);
					if (((AODLWarning)ob).Node != null)
					{
                                    XmlTextWriter writer = new XmlTextWriter(Console.Out)
                                    {
                                        Formatting = Formatting.Indented
                                    };
                                    ((AODLWarning)ob).Node.WriteContentTo(writer);
					}
				}
			}
		}


		/// <summary>
		/// Create a new XmlNode for this document.
		/// </summary>
		/// <param name="name">The elementname.</param>
		/// <param name="prefix">The prefix.</param>
		/// <returns>The new XmlNode.</returns>
		public XmlNode CreateNode(string name, string prefix)
		{
			
			if (XmlDoc == null)
				throw new NullReferenceException("There is no XmlDocument loaded. Couldn't create Node "+name+" with Prefix "+prefix+". "+GetType().ToString());
			string nuri = GetNamespaceUri(prefix);
			return XmlDoc.CreateElement(prefix, name, nuri);
		}

		/// <summary>
		/// Create a new XmlAttribute for this document.
		/// </summary>
		/// <param name="name">The attributename.</param>
		/// <param name="prefix">The prefixname.</param>
		/// <returns>The new XmlAttribute.</returns>
		public XmlAttribute CreateAttribute(string name, string prefix)
		{
			if (XmlDoc == null)
				throw new NullReferenceException("There is no XmlDocument loaded. Couldn't create Attribue "+name+" with Prefix "+prefix+". "+GetType().ToString());
			string nuri = GetNamespaceUri(prefix);
			return XmlDoc.CreateAttribute(prefix, name, nuri);
		}

		/// <summary>
		/// Return the namespaceuri for the given prefixname.
		/// </summary>
		/// <param name="prefix">The prefixname.</param>
		/// <returns></returns>
		public string GetNamespaceUri(string prefix)
		{
			foreach (string prefixx in NamespaceManager)
				if (prefix == prefixx)
				return NamespaceManager.LookupNamespace(prefixx);
			return null;
		}

		/// <summary>
		/// Load a blank the spreadsheet content document.
		/// </summary>
		private void LoadBlankContent()
		{
				Assembly ass			= Assembly.GetExecutingAssembly();
				Stream str				= ass.GetManifestResourceStream("DocumentGenerator.WordDocuments.AODL.Resources.OD.spreadsheetcontent.xml");
				_xmldoc			= new XmlDocument();
				_xmldoc.Load(str);
		}

		/// <summary>
		/// Tables the collection_ inserted.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void TableCollection_Inserted(int index, object value)
		{
			_tableCount++;
			if (!Content.Contains(value as IContent))
				Content.Add(value as IContent);
		}

		/// <summary>
		/// Tables the collection_ removed.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void TableCollection_Removed(int index, object value)
		{
			_tableCount--;
			if (Content.Contains(value as IContent))
				Content.Remove(value as IContent);
		}

		/// <summary>
		/// Content_s the inserted.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void Content_Inserted(int index, object value)
		{
			if (value is Table)
				if (!TableCollection.Contains((Table)value))
				TableCollection.Add(value as Table);
		}

		/// <summary>
		/// Content_s the removed.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void Content_Removed(int index, object value)
		{
			if (value is Table)
				if (TableCollection.Contains((Table)value))
				TableCollection.Remove(value as Table);
		}

		/// <summary>
		/// Creates the content body.
		/// </summary>
		private void CreateContentBody()
		{
			XmlNode nodeSpreadSheets		= XmlDoc.SelectSingleNode(
				"/office:document-content/office:body/office:spreadsheet", NamespaceManager);

			foreach(IContent iContent in Content)
			{
				//only table content
				if (iContent is Table)
					nodeSpreadSheets.AppendChild(((Table)iContent).BuildNode());
			}

			CreateLocalStyleContent();
			CreateCommonStyleContent();
		}

		/// <summary>
		/// Creates the content of the local style.
		/// </summary>
		private void CreateLocalStyleContent()
		{
			XmlNode nodeAutomaticStyles		= XmlDoc.SelectSingleNode(
				"/office:document-content/office:automatic-styles", NamespaceManager);

			foreach(IStyle style in Styles.ToValueList())
				nodeAutomaticStyles.AppendChild(style.Node);
		}

		/// <summary>
		/// Creates the content of the common style.
		/// </summary>
		private void CreateCommonStyleContent()
		{
			XmlNode nodeCommonStyles		= DocumentStyles.Styles.SelectSingleNode(
				"office:document-styles/office:styles", NamespaceManager);
			nodeCommonStyles.InnerXml	= "";

			foreach(IStyle style in CommonStyles.ToValueList())
			{
				XmlNode nodeStyle			= DocumentStyles.Styles.ImportNode(style.Node, true);
				nodeCommonStyles.AppendChild(nodeStyle);
			}

			//Remove styles node
			nodeCommonStyles				= XmlDoc.SelectSingleNode(
				"office:document-content/office:styles", NamespaceManager);
			if (nodeCommonStyles != null)
				XmlDoc.SelectSingleNode(
					"office:document-content", NamespaceManager).RemoveChild(nodeCommonStyles);
		}

		/// <summary>
		/// Styles the node exists.
		/// </summary>
		/// <param name="styleName">Name of the style.</param>
		/// <returns></returns>
		private bool StyleNodeExists(string styleName)
		{
			XmlNode styleNode				= XmlDoc.SelectSingleNode(
				"/office:document-content/office:automatic-styles/style:style[@style:name='"+styleName+"']", NamespaceManager);

			if (styleNode == null)
				return false;

			return true;
		}

		public EmbedObject  GetObjectByName(string ObjectName)
		{
			for(int i=0; i<EmbedObjects.Count ; i++)
			{
				if (((EmbedObject)EmbedObjects[i]).ObjectName==ObjectName)

					return (EmbedObject)EmbedObjects[i];

			}

			return null;
		}

		#region IDisposable Member

		private bool _disposed = false;

		/// <summary>
		/// Releases unmanaged resources and performs other cleanup operations before the
		/// is reclaimed by garbage collection.
		/// </summary>
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		/// <summary>
		/// Disposes the specified disposing.
		/// </summary>
		/// <param name="disposing">if set to <c>true</c> [disposing].</param>
		private void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				if (disposing)
				{
					try
					{
						DeleteUnpackedFiles();
					}
					catch(Exception ex)
					{
						throw ex;
					}
				}
			}
			_disposed = true;
		}

        /// <summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// <see cref="TextDocument"/> is reclaimed by garbage collection.
        /// </summary>
        ~SpreadsheetDocument()
		{
			Dispose();
		}

		#endregion
	}
}

/*
 * $Log: SpreadsheetDocument.cs,v $
 * Revision 1.3  2008/04/29 15:39:53  mt
 * new copyright header
 *
 * Revision 1.2  2008/02/08 07:12:20  larsbehr
 * - added initial chart support
 * - several bug fixes
 *
 * Revision 1.1  2007/02/25 08:58:47  larsbehr
 * initial checkin, import from Sourceforge.net to OpenOffice.org
 *
 * Revision 1.7  2007/02/04 22:52:57  larsbm
 * - fixed bug in resize algorithm for rows and cells
 * - extending IDocument, overload SaveTo to accept external exporter impl.
 * - initial version of AODL PDF exporter add on
 *
 * Revision 1.6  2006/02/21 19:34:55  larsbm
 * - Fixed Bug text that contains a xml tag will be imported  as UnknowText and not correct displayed if document is exported  as HTML.
 * - Fixed Bug [ 1436080 ] Common styles
 *
 * Revision 1.5  2006/02/06 19:27:23  larsbm
 * - fixed bug in spreadsheet document
 * - added smal OpenOfficeLib for document printing
 *
 * Revision 1.4  2006/02/05 20:03:32  larsbm
 * - Fixed several bugs
 * - clean up some messy code
 *
 * Revision 1.3  2006/02/02 21:55:59  larsbm
 * - Added Clone object support for many AODL object types
 * - New Importer implementation PlainTextImporter and CsvImporter
 * - New tests
 *
 * Revision 1.2  2006/01/29 18:52:14  larsbm
 * - Added support for common styles (style templates in OpenOffice)
 * - Draw TextBox import and export
 * - DrawTextBox html export
 *
 * Revision 1.1  2006/01/29 11:28:23  larsbm
 * - Changes for the new version. 1.2. see next changelog for details
 *
 */