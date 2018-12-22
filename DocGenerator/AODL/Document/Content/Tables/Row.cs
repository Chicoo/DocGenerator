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
using System.Xml;
using AODL.Document.Styles;
using AODL.Document.Content;
using AODL.Document.SpreadsheetDocuments;

namespace AODL.Document.Content.Tables
{
	/// <summary>
	/// Row represent a row within a table. If the row is part of a table which is
	/// used in a text document, then Cell merging is possible.
	/// </summary>
	public class Row : IContent, IHtml
	{
		/// <summary>
		/// RowChanged delegate
		/// </summary>
		public delegate void RowChanged(int rowNumber, int cellCount);

		private Table _table;
		/// <summary>
		/// Gets or sets the node.
		/// </summary>
		/// <value>The node.</value>
		public Table Table
		{
			get { return _table; }
			set { _table = value; }
		}

		/// <summary>
		/// Gets or sets the row style.
		/// </summary>
		/// <value>The row style.</value>
		public RowStyle RowStyle
		{
			get { return (RowStyle)Style; }
			set
			{
				StyleName	= ((RowStyle)value).StyleName;
				Style		= value;
			}
		}

		private CellCollection _cells;
		/// <summary>
		/// Gets or sets the cell collection.
		/// </summary>
		/// <value>The cell collection.</value>
		public CellCollection Cells
		{
			get { return _cells; }
			set { _cells = value; }
		}

		private CellSpanCollection _cellSpanCollection;
		/// <summary>
		/// Gets or sets the cell collection.
		/// </summary>
		/// <value>The cell collection.</value>
		public CellSpanCollection CellSpanCollection
		{
			get { return _cellSpanCollection; }
			set { _cellSpanCollection = value; }
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="Row"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="node">The node.</param>
		public Row(IDocument document, XmlNode node)
		{
			Document					= document;
			Node						= node;

			InitStandards();
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="Row"/> class.
		/// </summary>
		/// <param name="table">The table.</param>
		public Row(Table table)
		{
			Table						= table;
			Document					= table.Document;
			NewXmlNode(null);
			InitStandards();
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="Row"/> class.
		/// </summary>
		/// <param name="table">The table.</param>
		/// <param name="styleName">Name of the style.</param>
		public Row(Table table, string styleName)
		{
			Table						= table;
			Document					= table.Document;
			NewXmlNode(styleName);
			RowStyle					= Document.StyleFactory.Request<RowStyle>(styleName);
			InitStandards();
		}

		/// <summary>
		/// Inits the standards.
		/// </summary>
		private void InitStandards()
		{
			Cells				= new CellCollection();
			Cells.Removed		+= CellCollection_Removed;
			Cells.Inserted	+= CellCollection_Inserted;

			CellSpanCollection				= new CellSpanCollection();
			CellSpanCollection.Inserted	+= CellSpanCollection_Inserted;
			CellSpanCollection.Removed		+= CellSpanCollection_Removed;
		}

		/// <summary>
		/// Inserts the cell at the given zero based position.
		/// </summary>
		/// <param name="position">The position.</param>
		/// <param name="cell">The cell.</param>
		public void InsertCellAt(int position, Cell cell)
		{
			// Phil Jollans 16-Feb-2008: Largely rewritten
			if ( _cells.Count > position )
			{
				// We need to ask Lars how he intended this list to work.
				// Note that Table.Row_OnRowChanged() adds cells to all rows when
				// we add a cell to a single row.
				_cells.RemoveAt ( position ) ;
				_cells.Insert ( position, cell ) ;
			}
			else
			{
				while ( _cells.Count < position  )
				{
					_cells.Add ( new Cell ( Table.Document ) ) ;
				}
				
				_cells.Add(cell);
			}
		}

		/// <summary>
		/// Merge cells. This is only possible if the rows table is part
		/// of a text document.
		/// </summary>
		/// <param name="document">The TextDocument this row belongs to.</param>
		/// <param name="cellStartIndex">Start index of the cell.</param>
		/// <param name="mergeCells">The count of cells to merge incl. the starting cell.</param>
		/// <param name="mergeContent">if set to <c>true</c> [merge content].</param>
		public void MergeCells(TextDocuments.TextDocument document,int cellStartIndex, int mergeCells, bool mergeContent)
		{
			try
			{
				Cells[cellStartIndex].ColumnRepeating		= mergeCells.ToString();
				
				if (mergeContent)
				{
					for(int i=cellStartIndex+1; i<cellStartIndex+mergeCells; i++)
					{
						foreach(IContent content in Cells[i].Content)
							Cells[cellStartIndex].Content.Add(content);
					}
				}

				for(int i=cellStartIndex+mergeCells-1; i>cellStartIndex; i--)
				{
					Cells.RemoveAt(i);
					CellSpanCollection.Add(new CellSpan(this, (TextDocuments.TextDocument)Document));
				}
			}
			catch(Exception)
			{
				throw;
			}
		}

		/// <summary>
		/// Gets the index of the cell.
		/// </summary>
		/// <param name="cell">The cell.</param>
		/// <returns>The index of the cell wthin the cell collection. If the
		/// cell isn't part of the collection -1 will be returned.</returns>
		public int GetCellIndex(Cell cell)
		{
			if (cell != null && Cells != null)
			{
				for(int i=0; i<Cells.Count; i++)
					if (Cells[i].Equals(cell))
					return i;
			}

			return -1;
		}

		/// <summary>
		/// Create a new Xml node.
		/// </summary>
		/// <param name="styleName">Name of the style.</param>
		private void NewXmlNode(string styleName)
		{
			Node		= Document.CreateNode("table-row", "table");

			if (styleName != null)
			{
				XmlAttribute xa = Document.CreateAttribute("style-name", "table");
				xa.Value		= styleName;
				Node.Attributes.Append(xa);
			}
		}

		/// <summary>
		/// Create a XmlAttribute for propertie XmlNode.
		/// </summary>
		/// <param name="name">The attribute name.</param>
		/// <param name="text">The attribute value.</param>
		/// <param name="prefix">The namespace prefix.</param>
		private void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = Document.CreateAttribute(name, prefix);
			xa.Value		= text;
			Node.Attributes.Append(xa);
		}

		/// <summary>
		/// Cells the collection_ removed.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void CellCollection_Removed(int index, object value)
		{
			Node.RemoveChild(((Cell)value).Node);
		}

		/// <summary>
		/// Cells the collection_ inserted.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void CellCollection_Inserted(int index, object value)
		{
			//Only Spreadsheet documents are automaticaly resized
			//not needed if the file is loaded (right order!);
			if (Document is SpreadsheetDocument
			    && !Document.IsLoadedFile)
			{
				if (Node.ChildNodes.Count == index)
					Node.AppendChild(((Cell)value).Node);
				else
				{
					XmlNode childNode		= Node.ChildNodes[index];
					Node.InsertAfter(((Cell)value).Node, childNode);
				}
				Row_OnRowChanged(GetRowIndex(), Cells.Count);
			}
			else
				Node.AppendChild(((Cell)value).Node);
		}
		
		#region Eventhandling
		/// <summary>
		/// After a row size changed all rows before the changed row
		/// maybe need to be resized. This also belongs to the columns.
		/// </summary>
		/// <param name="rowNumber">The row number.</param>
		/// <param name="cellCount">The cell count.</param>
		private void Row_OnRowChanged(int rowNumber, int cellCount)
		{
			if (Table.ColumnCollection.Count <= cellCount)
			{
				for(int i=0; i <= cellCount-Table.ColumnCollection.Count; i++)
				{
					Table.ColumnCollection.Add(new Column(Table, string.Format("col{0}", Table.ColumnCollection.Count)));
				}
			}

			for(int i=0; i < rowNumber; i++)
			{
				Row row = Table.Rows[i];
				if (row.Cells.Count >= cellCount)
					continue;
				for(int ii = 0; ii < cellCount - row.Cells.Count; ii++)
				//for(int ii = cellCount - row.Cells.Count; ii < cellCount; ii++)
				{
					row.Cells.Add(new Cell(Table.Document));
				}
			}
		}
		#endregion

		/// <summary>
		/// Cells the span collection_ inserted.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void CellSpanCollection_Inserted(int index, object value)
		{
			Node.AppendChild(((CellSpan)value).Node);
		}

		/// <summary>
		/// Cells the span collection_ removed.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void CellSpanCollection_Removed(int index, object value)
		{
			Node.RemoveChild(((CellSpan)value).Node);
		}

		/// <summary>
		/// Gets the index of the row.
		/// </summary>
		/// <returns>The index within the table rowcollection of this row.</returns>
		private int GetRowIndex()
		{
			for(int i=0; i < Table.Rows.Count; i++)
			{
				if (Table.Rows[i] == this)
					return i;
			}
			//Maybe this row isn't already added.
			//e.g. this is a new row which will be added
			//to the end of the collection
			// TODO - there is no "Maybe" :) Throw exception if not found
			return Table.Rows.Count;
		}

		#region IContent Member

		/// <summary>
		/// Gets or sets the name of the style.
		/// </summary>
		/// <value>The name of the style.</value>
		public string StyleName
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@table:style-name",
				                                         Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@table:style-name",
				                                         Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("style-name", value, "table");
				_node.SelectSingleNode("@table:style-name",
				                            Document.NamespaceManager).InnerText = value;
			}
		}

		private IDocument _document;
		/// <summary>
		/// Every object (typeof(IContent)) have to know his document.
		/// </summary>
		/// <value></value>
		public IDocument Document
		{
			get
			{
				return _document;
			}
			set
			{
				_document = value;
			}
		}

		private IStyle _style;
		/// <summary>
		/// A Style class wich is referenced with the content object.
		/// If no style is available this is null.
		/// </summary>
		/// <value></value>
		public IStyle Style
		{
			get
			{
				return _style;
			}
			set
			{
				StyleName	= value.StyleName;
				_style = value;
			}
		}

		private XmlNode _node;
		/// <summary>
		/// Gets or sets the node.
		/// </summary>
		/// <value>The node.</value>
		public XmlNode Node
		{
			get { return _node; }
			set { _node = value; }
		}

		#endregion

		#region IHtml Member

		/// <summary>
		/// Return the content as Html string
		/// </summary>
		/// <returns>The html string</returns>
		public string GetHtml()
		{
			string html		= "<tr ";

			if (((RowStyle)Style).RowProperties != null)
				html		+= ((RowStyle)Style).RowProperties.GetHtmlStyle();

			html			+= ">\n";

			foreach(Cell cell in Cells)
				if (cell is IHtml)
				html	+= cell.GetHtml()+"\n";

			html			+= "</tr>";

			return html;
		}

		#endregion
	}
}

/*
 * $Log: Row.cs,v $
 * Revision 1.4  2008/04/29 15:39:46  mt
 * new copyright header
 *
 * Revision 1.3  2008/04/10 17:33:15  larsbehr
 * - Added several bug fixes mainly for the table handling which are submitted by Phil  Jollans
 *
 * Revision 1.2  2007/04/08 16:51:23  larsbehr
 * - finished master pages and styles for text documents
 * - several bug fixes
 *
 * Revision 1.1  2007/02/25 08:58:36  larsbehr
 * initial checkin, import from Sourceforge.net to OpenOffice.org
 *
 * Revision 1.3  2006/02/16 18:35:41  larsbm
 * - Add FrameBuilder class
 * - TextSequence implementation (Todo loading!)
 * - Free draing postioning via x and y coordinates
 * - Graphic will give access to it's full qualified path
 *   via the GraphicRealPath property
 * - Fixed Bug with CellSpan in Spreadsheetdocuments
 * - Fixed bug graphic of loaded files won't be deleted if they
 *   are removed from the content.
 * - Break-Before property for Paragraph properties for Page Break
 *
 * Revision 1.2  2006/01/29 18:52:14  larsbm
 * - Added support for common styles (style templates in OpenOffice)
 * - Draw TextBox import and export
 * - DrawTextBox html export
 *
 * Revision 1.1  2006/01/29 11:28:22  larsbm
 * - Changes for the new version. 1.2. see next changelog for details
 *
 */
