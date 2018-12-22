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
using AODL.Document;
using AODL.Document.Styles;
using AODL.Document.Content;

namespace AODL.Document.Content.Text
{
	/// <summary>
	/// Represent a list which could be a numbered or bullet style list.
	/// </summary>
	public class List : IContent, IContentContainer, IHtml
	{
		private ParagraphStyle _paragraphstyle;
		/// <summary>
		/// The ParagraphStyle to which this List belongs.
		/// There is only one ParagraphStyle per List and
		/// its ListItems.
		/// </summary>
		public ParagraphStyle ParagraphStyle
		{
			get { return _paragraphstyle; }
			set { _paragraphstyle = value; }
		}

		/// <summary>
		/// Gets or sets the list style.
		/// </summary>
		/// <value>The list style.</value>
		public ListStyle ListStyle
		{
			get { return (ListStyle)Style; }
			set { Style = (IStyle)value; }
		}

		private ContentCollection _content;
		/// <summary>
		/// The ContentCollection of access
		/// to their list items.
		/// </summary>
		public ContentCollection Content
		{
			get { return _content; }
			set { _content = value; }
		}

		private readonly ListStyles _type;
		/// <summary>
		/// Gets the type of the list.
		/// </summary>
		/// <value>The type of the list.</value>
		public ListStyles ListType
		{
			get { return _type; }
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="List"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="node">The node.</param>
		public List(IDocument document, XmlNode node)
		{
			Document						= document;
			Node							= node;
			InitStandards();
		}

		/// <summary>
		/// Create a new List object
		/// </summary>
		/// <param name="document">The IDocument</param>
		/// <param name="styleName">The style name</param>
		/// <param name="typ">The list typ bullet, ..</param>
		/// <param name="paragraphStyleName">The style name for the ParagraphStyle.</param>
		public List(IDocument document, string styleName, ListStyles typ, string paragraphStyleName)
		{
			Document						= document;
			NewXmlNode();
			InitStandards();

			Style							= (IStyle)new ListStyle(Document, styleName);
			ParagraphStyle					= new ParagraphStyle(Document, paragraphStyleName);
			Document.Styles.Add(Style);
			Document.Styles.Add(ParagraphStyle);
			
			ParagraphStyle.ListStyleName	= styleName;
			_type							= typ;

			((ListStyle)Style).AutomaticAddListLevelStyles(typ);			
		}

		/// <summary>
		/// Create a new List which is used to represent a inner list.
		/// </summary>
		/// <param name="document">The IDocument</param>
		/// <param name="outerlist">The List to which this List belongs.</param>
		public List(IDocument document, List outerlist)
		{
			Document						= document;
			ParagraphStyle					= outerlist.ParagraphStyle;			
			InitStandards();
			_type							= outerlist.ListType;
			//Create an inner list node, don't need a style
			//use the parents list style
			NewXmlNode();
		}

		/// <summary>
		/// Inits the standards.
		/// </summary>
		private void InitStandards()
		{
			Content				= new ContentCollection();

			Content.Inserted		+= Content_Inserted;
			Content.Removed		+= Content_Removed;
		}

		/// <summary>
		/// Create a new XmlNode.
		/// </summary>
		private void NewXmlNode()
		{			
			Node					= Document.CreateNode("list", "text");
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

		#region IContent Member
		/// <summary>
		/// Gets or sets the name of the style.
		/// </summary>
		/// <value>The name of the style.</value>
		public string StyleName
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@text:style-name",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@text:style-name",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("style-name", value, "text");
				_node.SelectSingleNode("@text:style-name",
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

		/// <summary>
		/// Content_s the inserted.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void Content_Inserted(int index, object value)
		{
			Node.AppendChild(((IContent)value).Node);
		}

		/// <summary>
		/// Content_s the removed.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void Content_Removed(int index, object value)
		{
			Node.RemoveChild(((IContent)value).Node);
		}

		#region IHtml Member

		/// <summary>
		/// Return the content as Html string
		/// </summary>
		/// <returns>The html string</returns>
		public string GetHtml()
		{
			string html			= null;

			if (ListType == ListStyles.Bullet)
				html			= "<ul>\n";
			else if (ListType == ListStyles.Number)
				html			= "<ol>\n";
			
			foreach(IContent content in Content)
				if (content is IHtml)
					html		+= ((IHtml)content).GetHtml();

			if (ListType == ListStyles.Bullet)
				html			+= "</ul>\n";
			else if (ListType == ListStyles.Number)
				html			+= "</ol>\n";
			//html				+= "</ul>\n";

			return html;
		}

		#endregion
	}
}
