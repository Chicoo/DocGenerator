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
using System.IO;
using System.Collections;
using AODL.Document;
using AODL.Document.Import;
using AODL.Document.Export;
using AODL.Document.Exceptions;
using AODL.Document.Content;
using AODL.Document.Content.Text;
using AODL.Document.Import.OpenDocument;

namespace AODL.Document.Import.PlainText
{
	/// <summary>
	/// Plain Text Importer.
	/// </summary>
	public class PlainTextImporter : IImporter, IPublisherInfo
	{
		private readonly DirInfo m_dirInfo = null;
		
		public DirInfo DirInfo {
			get { return m_dirInfo; }
		}
		
		
		/// <summary>
		/// The document to fill with content.
		/// </summary>
		private IDocument _document;

		/// <summary>
		/// Initializes a new instance of the <see cref="PlainTextImporter"/> class.
		/// </summary>
		public PlainTextImporter()
		{
			m_dirInfo = new DirInfo(string.Empty, string.Empty);
			_importError					= new ArrayList();

            _supportedExtensions = new ArrayList
            {
                new DocumentSupportInfo(".txt", DocumentTypes.TextDocument)
            };

            _author						= "Lars Behrmann, lb@OpenDocument4all.com";
			_infoUrl						= "http://AODL.OpenDocument4all.com";
			_description					= "This the standard importer for plain text files of the OpenDocument library AODL.";
		}

		#region IExporter Member

		private readonly ArrayList _supportedExtensions;
		/// <summary>
		/// Gets the document support infos.
		/// </summary>
		/// <value>The document support infos.</value>
		public ArrayList DocumentSupportInfos
		{
			get { return _supportedExtensions; }
		}

		/// <summary>
		/// Imports the specified filename.
		/// </summary>
		/// <param name="document">The TextDocument to fill.</param>
		/// <param name="filename">The filename.</param>
		/// <returns>The created TextDocument</returns>
		public void Import(IDocument document, string filename)
		{
			_document			= document;
			string text				= ReadContentFromFile(filename);
			
			if (text.Length > 0)
				ReadTextToDocument(text);
			else
			{
				AODLWarning warning	= new AODLWarning("Empty file. ["+filename+"]");
				ImportError.Add(warning);
			}
		}

		private readonly ArrayList _importError;
		/// <summary>
		/// Gets the import errors as ArrayList of strings.
		/// </summary>
		/// <value>The import errors.</value>
		public ArrayList ImportError
		{
			get
			{
				return _importError;
			}
		}

		/// <summary>
		/// If the import file format isn't any OpenDocument
		/// format you have to return true and AODL will
		/// create a new one.
		/// </summary>
		/// <value></value>
		public bool NeedNewOpenDocument
		{
			get { return true; }
		}

		#endregion

		#region IPublisherInfo Member

		private readonly string _author;
		/// <summary>
		/// The name the Author
		/// </summary>
		/// <value></value>
		public string Author
		{
			get
			{
				return _author;
			}
		}

		private readonly string _infoUrl;
		/// <summary>
		/// Url to a info site
		/// </summary>
		/// <value></value>
		public string InfoUrl
		{
			get
			{
				return _infoUrl;
			}
		}

		private readonly string _description;
		/// <summary>
		/// Description about the exporter resp. importer
		/// </summary>
		/// <value></value>
		public string Description
		{
			get
			{
				return _description;
			}
		}

		#endregion

		public void DeleteUnpackedFiles()
		{
			
		}
		
		/// <summary>
		/// Reads the text to document.
		/// </summary>
		/// <param name="text">The text.</param>
		private void ReadTextToDocument(string text)
		{
			ParagraphCollection parCol	= ParagraphBuilder.CreateParagraphCollection(
				_document, text, false, ParagraphBuilder.ParagraphSeperator);

			if (parCol != null)
				foreach(Paragraph paragraph in parCol)
				_document.Content.Add(paragraph);
		}

		/// <summary>
		/// Reads the content from file.
		/// </summary>
		/// <param name="fileName">Name of the file.</param>
		/// <returns></returns>
		private string ReadContentFromFile(string fileName)
		{
			string text					= "";

			try
			{
				StreamReader sReader	= File.OpenText(fileName);
				text					= sReader.ReadToEnd();
				sReader.Close();
			}
			catch(Exception ex)
			{
				throw ex;
			}

			return SetConformLineBreaks(text);
		}

		/// <summary>
		/// Sets the conform line breaks.
		/// </summary>
		/// <param name="text">The text.</param>
		/// <returns></returns>
		private string SetConformLineBreaks(string text)
		{
			return text.Replace(
				ParagraphBuilder.ParagraphSeperator2, ParagraphBuilder.ParagraphSeperator);
		}
	}
}

/*
 * $Log: PlainTextImporter.cs,v $
 * Revision 1.2  2008/04/29 15:39:53  mt
 * new copyright header
 *
 * Revision 1.1  2007/02/25 08:58:46  larsbehr
 * initial checkin, import from Sourceforge.net to OpenOffice.org
 *
 * Revision 1.1  2006/02/02 21:55:59  larsbm
 * - Added Clone object support for many AODL object types
 * - New Importer implementation PlainTextImporter and CsvImporter
 * - New tests
 *
 */