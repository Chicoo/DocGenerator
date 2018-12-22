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
using System.Drawing;
using System.Xml;
using AODL.Document.Styles;
using AODL.Document.Content.Text;
using AODL.Document.Content;
using AODL.Document;

namespace AODL.Document.Content.Draw
{
	/// <summary>
	/// DrawTextBox represent a draw text box which could be e.g. used
	/// to host graphic frame
	/// </summary>
	public class DrawTextBox : IContent, IContentContainer
	{
		/// <summary>
		/// If the context of the text box exceedes it's capacity,
		/// the content flows into the next text box in the chain
		/// and get or set the name of the next text box within this property.
		/// [optional]
		/// </summary>
		/// <value>The chain.</value>
		public string Chain
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:chain-next-name",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:chain-next-name",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("chain-next-name", value, "draw");
				_node.SelectSingleNode("@draw:chain-next-name",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the corner radius.
		/// [optional]
		/// </summary>
		/// <value>The corner radius.</value>
		public string CornerRadius
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:corner-radius",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:corner-radius",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("corner-radius", value, "draw");
				_node.SelectSingleNode("@draw:corner-radius",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the minimum height of the text box.
		/// [optional]
		/// </summary>
		/// <value>The height of the min.</value>
		public string MinHeight
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:min-height",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:min-height",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("min-height", value, "fo");
				_node.SelectSingleNode("@fo:min-height",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the minimum height of the text box.
		/// [optional]
		/// </summary>
		/// <value>The height of the min.</value>
		public string MinWidth
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:min-width",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:min-width",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("min-width", value, "fo");
				_node.SelectSingleNode("@fo:min-width",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the maximum width of the text box.
		/// [optional]
		/// </summary>
		/// <value>The height of the min.</value>
		public string MaxHeight
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:max-height",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:max-height",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("max-height", value, "fo");
				_node.SelectSingleNode("@fo:max-height",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the maximum width of the text box.
		/// [optional]
		/// </summary>
		/// <value>The height of the min.</value>
		public string MaxWidth
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:max-width",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:max-width",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("max-width", value, "fo");
				_node.SelectSingleNode("@fo:max-width",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="DrawTextBox"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="node">The node.</param>
		public DrawTextBox(IDocument document, XmlNode node)
		{
			Document				= document;
			InitStandards();
			Node					= node;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="DrawTextBox"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		public DrawTextBox(IDocument document)
		{
			Document				= document;
			InitStandards();
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
			Node		= Document.CreateNode("text-box", "draw");
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
				XmlNode xn = _node.SelectSingleNode("@draw:style-name",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:style-name",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("style-name", value, "draw");
				_node.SelectSingleNode("@draw:style-name",
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

		#region IContentContainer Member

		private ContentCollection _content;
		/// <summary>
		/// Gets or sets the content.
		/// </summary>
		/// <value>The content.</value>
		public ContentCollection Content
		{
			get
			{
				return _content;
			}
			set
			{
				_content = value;
			}
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
	}
}

/*
 * $Log: DrawTextBox.cs,v $
 * Revision 1.2  2008/04/29 15:39:44  mt
 * new copyright header
 *
 * Revision 1.1  2007/02/25 08:58:32  larsbehr
 * initial checkin, import from Sourceforge.net to OpenOffice.org
 *
 * Revision 1.1  2006/01/29 11:29:46  larsbm
 * *** empty log message ***
 *
 */