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
using AODL.Document.Styles.Properties;
using AODL.Document.Content.Text;
using AODL.Document.Content.Text.TextControl;
using AODL.Document.TextDocuments;

namespace AODL.Document.Content.Text.Indexes
{
	/// <summary>
	/// TableOfContent represent a table of contents.
	/// </summary>
	public class TableOfContents : IContent, IContentContainer, IHtml
	{
		/// <summary>
		/// This node will represent the visible entries
		/// </summary>
		private XmlNode _indexBodyNode;
		/// <summary>
		/// The style name for content entries
		/// </summary>
		private readonly string _contentStyleName		= "Contents_20_";
		/// <summary>
		/// The display name for content entries
		/// </summary>
		private readonly string _contentStyleDisplayName	= "Contents ";
		private bool _useHyperlinks;
		/// <summary>
		/// Gets or sets a value indicating whether [use hyperlinks].
		/// If it's set to true, the text entries automaticaly will
		/// extended with Hyperlinks.
		/// </summary>
		/// <value><c>true</c> if [use hyperlinks]; otherwise, <c>false</c>.</value>
		public bool UseHyperlinks
		{
			get { return _useHyperlinks; }
			set { _useHyperlinks = value; }
		}

		public XmlNode IndexBodyNode {
			get { return _indexBodyNode; }
			set { _indexBodyNode = value; }
		}
		
		/// <summary>
		/// Gets the title which is displayed
		/// for this table of content in english
		/// this will always be Table of Content,
		/// but this is free of choice.
		/// </summary>
		/// <value>The title.</value>
		public string Title
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@text:name",
				                                         ((TextDocument)Document).NamespaceManager);
				if (xn != null)
					//Todo: Algo for more entries then 9
					return xn.InnerText.Substring(0, xn.InnerText.Length-1);
				return null;
			}
		}

		private Paragraph _titleParagraph;
		/// <summary>
		/// Gets or sets the title paragraph.
		/// </summary>
		/// <value>The title paragraph.</value>
		public Paragraph TitleParagraph
		{
			get { return _titleParagraph; }
			set { _titleParagraph = value; }
		}

		private TableOfContentsSource _tableOfContentsSource;
		/// <summary>
		/// Gets or sets the table of content source.
		/// </summary>
		/// <value>The table of content source.</value>
		public TableOfContentsSource TableOfContentsSource
		{
			get { return _tableOfContentsSource; }
			set { _tableOfContentsSource = value; }
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="TableOfContents"/> class.
		/// </summary>
		/// <param name="textDocument">The text document.</param>
		/// <param name="styleName">Name of the style.</param>
		/// <param name="useHyperlinks">if set to <c>true</c> [use hyperlinks].</param>
		/// <param name="protectChanges">if set to <c>true</c> [protect changes].</param>
		/// <param name="textName">Title for the Table of content e.g. Table of Content</param>
		public TableOfContents(TextDocument textDocument, string styleName, bool useHyperlinks, bool protectChanges, string textName)
		{
			Document				= textDocument;
			UseHyperlinks			= useHyperlinks;
			NewXmlNode(styleName, protectChanges, textName);
			Style					= new SectionStyle(this, styleName);
			Document.Styles.Add(Style);
			
			TableOfContentsSource	= new TableOfContentsSource(this);
			TableOfContentsSource.InitStandardTableOfContentStyle();
			Node.AppendChild(TableOfContentsSource.Node);

			CreateIndexBody();
			CreateTitlePargraph();
			InsertContentStyle();
			SetOutlineStyle();
			RegisterEvents();
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="TableOfContents"/> class.
		/// </summary>
		/// <param name="textDocument">The text document.</param>
		/// <param name="tocNode">The toc node.</param>
		public TableOfContents(TextDocument textDocument, XmlNode tocNode)
		{
			Document				= textDocument;
			Node					= tocNode;
			RegisterEvents();
		}

		/// <summary>
		/// Registers the events.
		/// </summary>
		private void RegisterEvents()
		{
			Content				= new ContentCollection();
			Content.Inserted		+= Content_Inserted;
			Content.Removed		+= Content_Removed;
		}
		
		/// <summary>
		/// News the XML node.
		/// </summary>
		/// <param name="stylename">The stylename.</param>
		/// <param name="protectChanges">if set to <c>true</c> [protect changes].</param>
		/// <param name="textName">Name of the text.</param>
		private void NewXmlNode(string stylename, bool protectChanges, string textName)
		{
			Node		= ((TextDocument)Document).CreateNode("table-of-content", "text");

			XmlAttribute xa = ((TextDocument)Document).CreateAttribute("style-name", "text");
			xa.Value		= stylename;
			Node.Attributes.Append(xa);

			xa				= ((TextDocument)Document).CreateAttribute("protected", "text");
			xa.Value		= protectChanges.ToString().ToLower();
			Node.Attributes.Append(xa);

			xa				= ((TextDocument)Document).CreateAttribute("use-outline-level", "text");
			xa.Value		= "true";
			Node.Attributes.Append(xa);

			xa				= ((TextDocument)Document).CreateAttribute("name", "text");
			xa.Value		= (textName ?? "Table of Contents1");//+((TextDocument)this.Document).TableofContentsCount;
			Node.Attributes.Append(xa);
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
		/// Creates the index body node.
		/// </summary>
		private void CreateIndexBody()
		{
			_indexBodyNode		= ((TextDocument)Document).CreateNode("index-body", "text");
			
			//First not is always the index title
			XmlNode indexTitleNode	= ((TextDocument)Document).CreateNode("index-title", "text");
			
			//Create attributes for the index title
			XmlAttribute xa			= ((TextDocument)Document).CreateAttribute("style-name", "text");
			xa.Value				= StyleName;
			indexTitleNode.Attributes.Append(xa);
			
			xa						= ((TextDocument)Document).CreateAttribute("name", "text");
			xa.Value				= Title+"_Head";
			indexTitleNode.Attributes.Append(xa);

			_indexBodyNode.AppendChild(indexTitleNode);
			Node.AppendChild(_indexBodyNode);
		}

		/// <summary>
		/// Creates the title pargraph.
		/// </summary>
		private void CreateTitlePargraph()
		{
			TitleParagraph		= new Paragraph(((TextDocument)Document), "Table_Of_Contents_Title");
			TitleParagraph.TextContent.Add(new SimpleText(Document, Title));
			//Set default styles
			TitleParagraph.ParagraphStyle.TextProperties.Bold		= "bold";
			TitleParagraph.ParagraphStyle.TextProperties.FontName	= FontFamilies.Arial;
			TitleParagraph.ParagraphStyle.TextProperties.FontSize	= "20pt";
			//Add to the index title
			_indexBodyNode.ChildNodes[0].AppendChild(TitleParagraph.Node);
		}

		/// <summary>
		/// Insert the content style nodes. These are 10 styles for
		/// each outline number one style.
		/// TODO: Section Style move to document common styles
		/// </summary>
		private void InsertContentStyle()
		{
			for(int i=1; i<=10; i++)
			{
				XmlNode styleNode	= Document.CreateNode(
					"style", "style");

				XmlAttribute xa		= Document.CreateAttribute(
					"name", "style");
				xa.InnerText		= _contentStyleName+i.ToString();
				styleNode.Attributes.Append(xa);

				xa					= Document.CreateAttribute(
					"display-name", "style");
				xa.InnerText		= _contentStyleDisplayName+i.ToString();
				styleNode.Attributes.Append(xa);

				xa					= Document.CreateAttribute(
					"parent-style-name", "style");
				xa.InnerText		= "Index";
				styleNode.Attributes.Append(xa);

				xa					= Document.CreateAttribute(
					"family", "style");
				xa.InnerText		= "paragraph";
				styleNode.Attributes.Append(xa);

				xa					= Document.CreateAttribute(
					"class","style");
				xa.InnerText		= "index";
				styleNode.Attributes.Append(xa);
				
				XmlNode ppNode		= Document.CreateNode(
					"paragraph-properties", "style");

				xa					= Document.CreateAttribute(
					"margin-left", "fo");
				xa.InnerText		= (0.499*(i-1)).ToString("F3").Replace(",",".")+"cm";
				ppNode.Attributes.Append(xa);

				xa					= Document.CreateAttribute(
					"margin-right", "fo");
				xa.InnerText		= "0cm";
				ppNode.Attributes.Append(xa);

				xa					= Document.CreateAttribute(
					"text-indent", "fo");
				xa.InnerText		= "0cm";
				ppNode.Attributes.Append(xa);

				xa					= Document.CreateAttribute(
					"auto-text-indent", "fo");
				xa.InnerText		= "0cm";
				ppNode.Attributes.Append(xa);

				XmlNode tabsNode		= Document.CreateNode(
					"tab-stops", "style");

				XmlNode tabNode			= Document.CreateNode(
					"tab-stop", "style");

				xa					= Document.CreateAttribute(
					"position", "style");
				xa.InnerText		= (16.999-(i*0.499)).ToString("F3").Replace(",",".")+"cm";
				tabNode.Attributes.Append(xa);

				xa					= Document.CreateAttribute(
					"type", "style");
				xa.InnerText		= "right";
				tabNode.Attributes.Append(xa);

				xa					= Document.CreateAttribute(
					"leader-style", "style");
				xa.InnerText		= "dotted";
				tabNode.Attributes.Append(xa);

				xa					= Document.CreateAttribute(
					"leader-text", "style");
				xa.InnerText		= ".";
				tabNode.Attributes.Append(xa);

				tabsNode.AppendChild(tabNode);
				ppNode.AppendChild(tabsNode);
				styleNode.AppendChild(ppNode);

				IStyle iStyle		= new UnknownStyle(Document, styleNode);
				Document.CommonStyles.Add(iStyle);

//				XmlNode styleNode	= ((TextDocument)this.Document).DocumentStyles.Styles.CreateElement(
//					"style", "style", ((TextDocument)this.Document).GetNamespaceUri("style"));
//
//				XmlAttribute xa		= ((TextDocument)this.Document).DocumentStyles.Styles.CreateAttribute(
//					"style", "name", ((TextDocument)this.Document).GetNamespaceUri("style"));
//				xa.InnerText		= this._contentStyleName+i.ToString();
//				styleNode.Attributes.Append(xa);
//
//				xa					= ((TextDocument)this.Document).DocumentStyles.Styles.CreateAttribute(
//					"style", "display-name", ((TextDocument)this.Document).GetNamespaceUri("style"));
//				xa.InnerText		= this._contentStyleDisplayName+i.ToString();
//				styleNode.Attributes.Append(xa);
//
//				xa					= ((TextDocument)this.Document).DocumentStyles.Styles.CreateAttribute(
//					"style", "parent-style-name", ((TextDocument)this.Document).GetNamespaceUri("style"));
//				xa.InnerText		= "Index";
//				styleNode.Attributes.Append(xa);
//
//				xa					= ((TextDocument)this.Document).DocumentStyles.Styles.CreateAttribute(
//					"style", "family", ((TextDocument)this.Document).GetNamespaceUri("style"));
//				xa.InnerText		= "paragraph";
//				styleNode.Attributes.Append(xa);
//
//				xa					= ((TextDocument)this.Document).DocumentStyles.Styles.CreateAttribute(
//					"style", "class", ((TextDocument)this.Document).GetNamespaceUri("style"));
//				xa.InnerText		= "index";
//				styleNode.Attributes.Append(xa);
//
//				XmlNode ppNode		= ((TextDocument)this.Document).DocumentStyles.Styles.CreateElement(
//					"style", "paragraph-properties", ((TextDocument)this.Document).GetNamespaceUri("style"));
//
//				xa					= ((TextDocument)this.Document).DocumentStyles.Styles.CreateAttribute(
//					"fo", "margin-left", ((TextDocument)this.Document).GetNamespaceUri("fo"));
//				xa.InnerText		= (0.499*(i-1)).ToString("F3").Replace(",",".")+"cm";
//				ppNode.Attributes.Append(xa);
//
//				xa					= ((TextDocument)this.Document).DocumentStyles.Styles.CreateAttribute(
//					"fo", "margin-right", ((TextDocument)this.Document).GetNamespaceUri("fo"));
//				xa.InnerText		= "0cm";
//				ppNode.Attributes.Append(xa);
//
//				xa					= ((TextDocument)this.Document).DocumentStyles.Styles.CreateAttribute(
//					"fo", "text-indent", ((TextDocument)this.Document).GetNamespaceUri("fo"));
//				xa.InnerText		= "0cm";
//				ppNode.Attributes.Append(xa);
//
//				xa					= ((TextDocument)this.Document).DocumentStyles.Styles.CreateAttribute(
//					"fo", "auto-text-indent", ((TextDocument)this.Document).GetNamespaceUri("fo"));
//				xa.InnerText		= "0cm";
//				ppNode.Attributes.Append(xa);
//
//				XmlNode tabsNode		= ((TextDocument)this.Document).DocumentStyles.Styles.CreateElement(
//					"style", "tab-stops", ((TextDocument)this.Document).GetNamespaceUri("style"));
//
//				XmlNode tabNode			= ((TextDocument)this.Document).DocumentStyles.Styles.CreateElement(
//					"style", "tab-stop", ((TextDocument)this.Document).GetNamespaceUri("style"));
//
//				xa					= ((TextDocument)this.Document).DocumentStyles.Styles.CreateAttribute(
//					"style", "position", ((TextDocument)this.Document).GetNamespaceUri("style"));
//				xa.InnerText		= (16.999-(i*0.499)).ToString("F3").Replace(",",".")+"cm";
//				tabNode.Attributes.Append(xa);
//
//				xa					= ((TextDocument)this.Document).DocumentStyles.Styles.CreateAttribute(
//					"style", "type", ((TextDocument)this.Document).GetNamespaceUri("style"));
//				xa.InnerText		= "right";
//				tabNode.Attributes.Append(xa);
//
//				xa					= ((TextDocument)this.Document).DocumentStyles.Styles.CreateAttribute(
//					"style", "leader-style", ((TextDocument)this.Document).GetNamespaceUri("style"));
//				xa.InnerText		= "dotted";
//				tabNode.Attributes.Append(xa);
//
//				xa					= ((TextDocument)this.Document).DocumentStyles.Styles.CreateAttribute(
//					"style", "leader-text", ((TextDocument)this.Document).GetNamespaceUri("style"));
//				xa.InnerText		= ".";
//				tabNode.Attributes.Append(xa);
//
//				tabsNode.AppendChild(tabNode);
//				ppNode.AppendChild(tabsNode);
//				styleNode.AppendChild(ppNode);
//
//				((TextDocument)this.Document).DocumentStyles.InsertOfficeStylesNode(
//					styleNode, ((TextDocument)this.Document));
			}
		}

		/// <summary>
		/// Set the outline style.
		/// </summary>
		private void SetOutlineStyle()
		{
			for(int i=1; i<=10; i++)
			{
				((TextDocument)Document).DocumentStyles.SetOutlineStyle(
					i, "1", ((TextDocument)Document));
			}
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

		#region IContentContainer Member

		private ContentCollection _content;
		/// <summary>
		/// Represent all visible entries of this Table of contents.
		/// Normally you should only insert paragraph objects as
		/// entry which match the following structure
		/// e.g. 3.1 My header text \t
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
		/// Content was inserted.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void Content_Inserted(int index, object value)
		{
			_indexBodyNode.AppendChild(((IContent)value).Node);
		}

		/// <summary>
		/// Content was removed.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void Content_Removed(int index, object value)
		{
			_indexBodyNode.RemoveChild(((IContent)value).Node);
		}

		/// <summary>
		/// Gets the tab stop style.
		/// </summary>
		/// <param name="leaderStyle">The leader style.</param>
		/// <param name="leadingChar">The leading char.</param>
		/// <param name="position">The position.</param>
		/// <returns>A for a table of contents optimized TabStopStyleCollection</returns>
		public TabStopStyleCollection GetTabStopStyle(string leaderStyle, string leadingChar, double position)
		{
			TabStopStyleCollection tabStopStyleCol	= new TabStopStyleCollection(((TextDocument)Document));
            //Create TabStopStyles
            TabStopStyle tabStopStyle = new TabStopStyle(((TextDocument)Document), position)
            {
                LeaderStyle = leaderStyle,
                LeaderText = leadingChar,
                Type = TabStopTypes.Center
            };
            //Add the tabstop
            tabStopStyleCol.Add(tabStopStyle);

			return tabStopStyleCol;
		}

		/// <summary>
		/// Insert the given text as an Table of contents entry.
		/// e.g. You just insert a Headline 1. My headline to
		/// the document and want this text also as an Table of
		/// contents entry, so you can simply add the text using
		/// this method.
		/// </summary>
		/// <param name="textEntry">The text entry.</param>
		/// <param name="outLineLevel">The outline level possible 1-10.</param>
		public void InsertEntry(string textEntry, int outLineLevel)
		{
			Paragraph paragraph				= new Paragraph(
				((TextDocument)Document), "P1_Toc_Entry"+outLineLevel.ToString());
			((ParagraphStyle)paragraph.Style).ParentStyle =
				_contentStyleName+outLineLevel.ToString();
			if (UseHyperlinks)
			{
				int firstWhiteSpace			= textEntry.IndexOf(" ");
				System.Text.StringBuilder sb = new System.Text.StringBuilder(textEntry);
				sb							= sb.Remove(firstWhiteSpace, 1);
				string link					= "#"+sb.ToString()+"|outline";
                XLink xlink = new XLink(Document, link, textEntry)
                {
                    XLinkType = "simple"
                };
                paragraph.TextContent.Add(xlink);
				paragraph.TextContent.Add(new TabStop(Document));
				paragraph.TextContent.Add(new SimpleText(Document, "1"));
			}
			else
			{
				//add the tabstop and the page number, the number is
				//always set to 1, but will be updated by the most
				//word processors immediately to the correct page number.
				paragraph.TextContent.Add(new SimpleText(Document, textEntry));
				paragraph.TextContent.Add(new TabStop(Document));
				paragraph.TextContent.Add(new SimpleText(Document, "1"));
			}
			//There is a bug which deny to add new simple ta
//			this.TitleParagraph.ParagraphStyle.ParagraphProperties.TabStopStyleCollection =
//				this.GetTabStopStyle(TabStopLeaderStyles.Dotted, ".", 16.999);
			paragraph.ParagraphStyle.ParagraphProperties.TabStopStyleCollection =
				GetTabStopStyle(TabStopLeaderStyles.Dotted, ".", 16.999);

			Content.Add(paragraph);
		}

		#region IHtml Member

		/// <summary>
		/// Return the content as Html string
		/// </summary>
		/// <returns>The html string</returns>
		public string GetHtml()
		{
			string html					= "<br>&nbsp;\n";
			try
			{
				foreach(IContent content in Content)
					if (content is IHtml)
					html				+= ((IHtml)content).GetHtml()+"\n";
			}
			catch(Exception)
			{
			}
			return html;
		}

		#endregion
	}
}

/*
 * $Log: TableOfContents.cs,v $
 * Revision 1.3  2008/04/29 15:39:47  mt
 * new copyright header
 *
 * Revision 1.2  2007/04/08 16:51:24  larsbehr
 * - finished master pages and styles for text documents
 * - several bug fixes
 *
 * Revision 1.1  2007/02/25 08:58:40  larsbehr
 * initial checkin, import from Sourceforge.net to OpenOffice.org
 *
 * Revision 1.4  2006/02/21 19:34:55  larsbm
 * - Fixed Bug text that contains a xml tag will be imported  as UnknowText and not correct displayed if document is exported  as HTML.
 * - Fixed Bug [ 1436080 ] Common styles
 *
 * Revision 1.3  2006/02/05 20:02:25  larsbm
 * - Fixed several bugs
 * - clean up some messy code
 *
 * Revision 1.2  2006/01/29 18:52:14  larsbm
 * - Added support for common styles (style templates in OpenOffice)
 * - Draw TextBox import and export
 * - DrawTextBox html export
 *
 * Revision 1.1  2006/01/29 11:28:22  larsbm
 * - Changes for the new version. 1.2. see next changelog for details
 *
 * Revision 1.1  2006/01/05 10:31:10  larsbm
 * - AODL merged cells
 * - AODL toc
 * - AODC batch mode, splash screen
 *
 */