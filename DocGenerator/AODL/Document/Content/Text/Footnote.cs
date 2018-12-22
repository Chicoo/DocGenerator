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
using AODL.Document.TextDocuments;

namespace AODL.Document.Content.Text
{
	/// <summary>
	/// Represent a Footnote which could be 
	/// a Foot- or a Endnote
	/// </summary>
	public class Footnote : IText, IHtml, ITextContainer
	{
		/// <summary>
		/// Gets or sets the id.
		/// </summary>
		/// <value>The id.</value>
		public string Id
		{
			get 
			{
				XmlNode xn = _node.SelectSingleNode("//@text:note-citation", 
					Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{
				XmlNode xn = _node.SelectSingleNode("//@text:note-citation",
					Document.NamespaceManager);
				if (xn != null)
				{
					_node.SelectSingleNode("//@text:note-citation",
						Document.NamespaceManager).InnerText = value;
					_node.SelectSingleNode("//@text:id",
						Document.NamespaceManager).InnerText = "ftn"+value;
				}
			}
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="Footnote"/> class.
		/// </summary>
		/// <param name="document">The text document.</param>
		/// <param name="notetext">The notetext.</param>
		/// <param name="id">The id.</param>
		/// <param name="type">The type.</param>
		public Footnote(IDocument document, string notetext, string id, FootnoteType type)
		{
			Document		= document;
			NewXmlNode(id, notetext, type);
			InitStandards();
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="Footnote"/> class.
		/// </summary>
		/// <param name="document">The text document.</param>
		public Footnote(IDocument document)
		{
			Document		= document;
			InitStandards();
		}

		/// <summary>
		/// Inits the standards.
		/// </summary>
		private void InitStandards()
		{
			TextContent			= new ITextCollection();
			TextContent.Inserted	+= TextContent_Inserted;
			TextContent.Removed	+= TextContent_Removed;
		}

		/// <summary>
		/// News the XML node.
		/// </summary>
		/// <param name="id">The id.</param>
		/// <param name="notetext">The notetext.</param>
		/// <param name="type">The type.</param>
		private void NewXmlNode(string id, string notetext, FootnoteType type)
		{			
			Node		= Document.CreateNode("note", "text");
			
			XmlAttribute xa = Document.CreateAttribute("id", "text");
			xa.Value		= "ftn"+id;
			Node.Attributes.Append(xa);

			xa				 = Document.CreateAttribute("note-class", "text");
			xa.Value		= type.ToString();
			Node.Attributes.Append(xa);

			//Node citation
			XmlNode node	 = Document.CreateNode("not-citation", "text");
			node.InnerText	 = id;

			_node.AppendChild(node);

			//Node Footnode body
			XmlNode nodebody = Document.CreateNode("note-body", "text");

			//Node Footnode text
			node			 = Document.CreateNode("p", "text");
			node.InnerXml	 = notetext;

			xa				 = Document.CreateAttribute("style-name", "text");
			xa.Value		 = (type == FootnoteType.footnode)?"Footnote":"Endnote";
			node.Attributes.Append(xa);

			nodebody.AppendChild(node);

			_node.AppendChild(nodebody);
		}

		/// <summary>
		/// Texts the content_ inserted.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void TextContent_Inserted(int index, object value)
		{
			Node.AppendChild(((IText)value).Node);
		}

		/// <summary>
		/// Texts the content_ removed.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void TextContent_Removed(int index, object value)
		{
			Node.RemoveChild(((IText)value).Node);
		}

		#region IText Member

		private XmlNode _node;
		/// <summary>
		/// The node that represent the text content.
		/// </summary>
		/// <value></value>
		public XmlNode Node
		{
			get
			{
				return _node;
			}
			set
			{
				_node = value;
			}
		}

		/// <summary>
		/// The text.
		/// </summary>
		/// <value></value>
		public string Text
		{
			get
			{
				return Node.InnerText;
			}
			set
			{
				Node.InnerText	= value;
			}
		}

		private IDocument _document;
		/// <summary>
		/// The document to which this text content belongs to.
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

		/// <summary>
		/// Is null no style is available.
		/// </summary>
		/// <value></value>
		public IStyle Style
		{
			get { return null; }
			set { }
		}

		/// <summary>
		/// No style name available
		/// </summary>
		/// <value></value>
		public string StyleName
		{
			get 
			{ 
				return null;
			}
			set
			{
			}
		}

		#endregion

		#region ITextContainer Member

		private ITextCollection _textContent;
		/// <summary>
		/// All Content objects have a Text container. Which represents
		/// his Text this could be SimpleText, FormatedText or mixed.
		/// </summary>
		/// <value></value>
		public ITextCollection TextContent
		{
			get
			{
				return _textContent;
			}
			set
			{
				if (_textContent != null)
					foreach(IText text in _textContent)
						Node.RemoveChild(text.Node);

				_textContent = value;
				
				if (_textContent != null)
					foreach(IText text in _textContent)
						Node.AppendChild(text.Node);
			}
		}

		#endregion

		#region IHtml Member

		/// <summary>
		/// Return the content as Html string
		/// </summary>
		/// <returns>The html string</returns>
		public string GetHtml()
		{
			string html			= "<sup>(";
			html				+= Id;
			html				+= ". "+Text;
			html				+= ")</sup>";

			return html;
		}

		#endregion
		
	}

	/// <summary>
	/// Represent the possible footnodes
	/// </summary>
	public enum FootnoteType
	{
		/// <summary>
		/// A footnode
		/// </summary>
		footnode,
		/// <summary>
		/// A endnote
		/// </summary>
		endnote
	}
}

/*
 * $LoG$
 */