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
using AODL.Document.Import.OpenDocument.NodeProcessors;

namespace AODL.Document.Content.Text
{
	/// <summary>
	/// Represent a hyperlink, which could be 
	/// a web-, ftp- or telnet link
	/// </summary>
	public class XLink : IText, IHtml, ITextContainer, ICloneable
	{	
		/// <summary>
		/// Gets or sets the href. e.g http://www.sourceforge.net
		/// </summary>
		/// <value>The href.</value>
		public string Href
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@xlink:href", 
					Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@xlink:href",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("href", value, "xlink");
				_node.SelectSingleNode("@xlink:href",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the type of the X link.
		/// </summary>
		/// <value>The type of the X link.</value>
		public string XLinkType
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@xlink:type", 
					Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@xlink:type",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("type", value, "xlink");
				_node.SelectSingleNode("@xlink:type",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// This is not the display text! Gets or sets the office name 
		/// </summary>
		/// <value>The name of the office.</value>
		public string OfficeName
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@office:name", 
					Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@office:name",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("name", value, "office");
				_node.SelectSingleNode("@office:name",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the name of the target frame.
		/// e.g. _blank, _top
		/// </summary>
		/// <value>The name of the target frame.</value>
		public string TargetFrameName
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@office:target-frame-name", 
					Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@office:target-frame-name",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("target-frame-name", value, "office");
				_node.SelectSingleNode("@office:target-frame-name",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the show.
		/// Standard value is <b>new</b>
		/// </summary>
		/// <value>The show.</value>
		public string Show
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@xlink:show", 
					Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@xlink:show",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("show", value, "xlink");
				_node.SelectSingleNode("@xlink:show",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="XLink"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="href">The href.</param>
		/// <param name="name">The name.</param>
		public XLink(IDocument document, string href, string name)
		{
			Document			= document;
			NewXmlNode();
			InitStandards();

			Href				= href;
			TextContent.Add(new SimpleText(Document, name));
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="XLink"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		public XLink(IDocument document)
		{
			Document			= document;
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
		private void NewXmlNode()
		{
			Node		= Document.CreateNode("a", "text");
		}

		/// <summary>
		/// Creates the attribute.
		/// </summary>
		/// <param name="name">The name.</param>
		/// <param name="text">The text.</param>
		/// <param name="prefix">The prefix.</param>
		private void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = Document.CreateAttribute(name, prefix);
			xa.Value		= text;
			Node.Attributes.Append(xa);
		}

		/// <summary>
		/// Append the xml from added IText object.
		/// </summary>
		/// <param name="index"></param>
		/// <param name="value"></param>
		private void TextContent_Inserted(int index, object value)
		{
			Node.AppendChild(((IText)value).Node);

			if (((IText)value).Text != null)
			{
				try
				{
					if (Document is TextDocuments.TextDocument)
					{
						string text		= ((IText)value).Text;
						Document.DocumentMetadata.CharacterCount	+= text.Length;
						string[] words	= text.Split(' ');
						Document.DocumentMetadata.WordCount		+= words.Length;
					}
				}
				catch(Exception)
				{
					//unhandled, only word and character count wouldn' be correct
				}
			}
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
				Node.InnerText = value;
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

		private IStyle _style;
		/// <summary>
		/// The style which is referenced with this text object.
		/// This is null if no style is available.
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
				_style = value;
			}
		}

		private string _styleName;
		/// <summary>
		/// The style name which is used for the referenced style.
		/// This is null is no  style is available.
		/// </summary>
		/// <value></value>
		public string StyleName
		{
			get
			{
				return _styleName;
			}
			set
			{
				_styleName = value;
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
			string html		= "<a href=\"";

			if (Href != null)
				//html	+= this.Href+"\"";
				html	+= GetLink()+"\"";

			if (TargetFrameName != null)
				html	+= " target=\""+TargetFrameName+"\"";

			if (Href != null)
			{
				html	+= ">\n";
				html	+= Text;
				html	+= "</a>";
			}
			else
				html	= "";
				
			return html;
		}

		/// <summary>
		/// Gets the link.
		/// </summary>
		/// <returns></returns>
		private string GetLink()
		{
			//string outline		= "|outline";
			string link			= Href;
//			if (this.Document.TableofContentsCount > 0
//				&& link.StartsWith("#") && link.EndsWith(outline))
//			{
//				link			= link.Replace(outline, "");
//				int lastWS		= link.IndexOf(" ");
//				if (lastWS != -1)
//					link		= link.Substring(lastWS);
//				foreach(AODL.Document.TextDocuments.Content.IContent content in this.Document.Content)
//					if (content is Header)
//						foreach(IText text in ((Header)content).TextContent)
//							if (text.Text.EndsWith(link))
//							{
//								return "#"+text.Text;
//							}
//			}
			return Href;
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

		#region ICloneable Member
		/// <summary>
		/// Create a deep clone of this XLink object.
		/// </summary>
		/// <remarks>A possible Attached Style wouldn't be cloned!</remarks>
		/// <returns>
		/// A clone of this object.
		/// </returns>
		public object Clone()
		{
			XLink xLinkClone			= null;

			if (Document != null && Node != null)
			{
				TextContentProcessor tcp	= new TextContentProcessor();
				xLinkClone					= tcp.CreateXLink(Document, Node.CloneNode(true));
			}

			return xLinkClone;
		}

		#endregion
	}
}

/*
 * $Log: XLink.cs,v $
 * Revision 1.3  2008/04/29 15:39:46  mt
 * new copyright header
 *
 * Revision 1.2  2007/04/08 16:51:23  larsbehr
 * - finished master pages and styles for text documents
 * - several bug fixes
 *
 * Revision 1.1  2007/02/25 08:58:40  larsbehr
 * initial checkin, import from Sourceforge.net to OpenOffice.org
 *
 * Revision 1.2  2006/02/02 21:55:59  larsbm
 * - Added Clone object support for many AODL object types
 * - New Importer implementation PlainTextImporter and CsvImporter
 * - New tests
 *
 * Revision 1.1  2006/01/29 11:28:22  larsbm
 * - Changes for the new version. 1.2. see next changelog for details
 *
 * Revision 1.3  2006/01/05 10:31:10  larsbm
 * - AODL merged cells
 * - AODL toc
 * - AODC batch mode, splash screen
 *
 * Revision 1.2  2005/12/12 19:39:17  larsbm
 * - Added Paragraph Header
 * - Added Table Row Header
 * - Fixed some bugs
 * - better whitespace handling
 * - Implmemenation of HTML Exporter
 *
 * Revision 1.1  2005/11/20 17:31:20  larsbm
 * - added suport for XLinks, TabStopStyles
 * - First experimental of loading dcuments
 * - load and save via importer and exporter interfaces
 *
 */