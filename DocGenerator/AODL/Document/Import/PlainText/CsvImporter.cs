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
using System.Diagnostics;
using AODL.Document;
using AODL.Document.Import;
using AODL.Document.Export;
using AODL.Document.Exceptions;
using AODL.Document.Content;
using AODL.Document.Content.Text;
using AODL.Document.Content.Tables;
using AODL.Document.SpreadsheetDocuments;
using AODL.Document.Import.OpenDocument;

namespace AODL.Document.Import.PlainText
{
	/// <summary>
	/// CsvImporter, a class for importing csv files into
	/// OpenDocument spreadsheet documents.
	/// </summary>
	public class CsvImporter : IImporter, IPublisherInfo
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
		/// Initializes a new instance of the <see cref="CsvImporter"/> class.
		/// </summary>
		public CsvImporter()
		{
			m_dirInfo = new DirInfo(string.Empty, string.Empty);
			_importError					= new ArrayList();

            _supportedExtensions = new ArrayList
            {
                new DocumentSupportInfo(".csv", DocumentTypes.SpreadsheetDocument)
            };

            _author						= "Lars Behrmann, lb@OpenDocument4all.com";
			_infoUrl						= "http://AODL.OpenDocument4all.com";
			_description					= "This the standard importer for comma seperated text files of the OpenDocument library AODL.";
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
			ArrayList lines			= GetFileContent(filename);
			
			if (lines.Count > 0)
				CreateTables(lines);
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
		/// Creates the tables.
		/// </summary>
		/// <param name="lines">The lines.</param>
		private void CreateTables(ArrayList lines)
		{
			string unicodeDelimiter				= "\u00BF"; // turned question mark

			if (lines != null)
			{
				Table table						= TableBuilder.CreateSpreadsheetTable(
					(SpreadsheetDocument)_document, "Table1", "table1");
				//First line must specify the used delimiter
				string delimiter				= lines[0] as string;
				lines.RemoveAt(0);

				try
				{
					//Perform lines
					foreach(string line in lines)
					{
						string lineContent			= line.Replace(delimiter, unicodeDelimiter);
						string[] cellContents		= lineContent.Split(unicodeDelimiter.ToCharArray());
						Row row						= new Row(table);
						foreach(string cellContent in cellContents)
						{
							Cell cell				= new Cell(table.Document);
							Paragraph paragraph		= ParagraphBuilder.CreateSpreadsheetParagraph(_document);
							paragraph.TextContent.Add(new SimpleText(_document, cellContent));
							cell.Content.Add(paragraph);
							row.InsertCellAt(row.Cells.Count, cell);
						}
						table.Rows.Add(row);
					}
				}
				catch(Exception ex)
				{
					throw new AODLException("Error while proccessing the csv file.", ex);
				}

				_document.Content.Add(table);
			}
		}

		/// <summary>
		/// Gets the content of the file.
		/// </summary>
		/// <param name="fileName">Name of the file.</param>
		/// <returns>All text lines as an ArrayList of strings.</returns>
		private ArrayList GetFileContent(string fileName)
		{
			ArrayList lines						= new ArrayList();

			try
			{
				StreamReader sReader	= File.OpenText(fileName);
				string currentLine		= null;

				while((currentLine = sReader.ReadLine()) != null)
				{
					lines.Add(currentLine);
				}
				sReader.Close();
			}
			catch(Exception ex)
			{
				throw ex;
			}

			return lines;
		}
	}
}

/*
 * $Log: CsvImporter.cs,v $
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