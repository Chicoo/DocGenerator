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

namespace AODL.Document.Styles.Properties
{
	/// <summary>
	/// Represent access to the possible attributes of of an paragraph-propertie element.
	/// </summary>
	public class ParagraphProperties : IProperty, IHtmlStyle
	{
		/// <summary>
		/// The ParagraphStyle object to this object belongs.
		/// </summary>
		public ParagraphStyle Paragraphstyle
		{
			get { return (ParagraphStyle)Style; }
			set { Style = value; }
		}

		/// <summary>
		/// Margin left. in cm an object .MarginLeft = "1cm";
		/// </summary>
		public string MarginLeft
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:margin-left", 
					  Style.Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:margin-left",
					  Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("margin-left", value, "fo");
				_node.SelectSingleNode("@fo:margin-left",
					  Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the break before.
		/// e.g. set this to "page" if this paragraph
		/// should be start on the next page.
		/// </summary>
		/// <value>The break before.</value>
		public string BreakBefore
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:break-before", 
					Style.Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:break-before",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("break-before", value, "fo");
				_node.SelectSingleNode("@fo:break-before",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the break after.
		/// e.g. set this to "page" if after this paragraph
		/// should be added a new page.
		/// </summary>
		/// <value>The break before.</value>
		public string BreakAfter
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:break-after", 
					Style.Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:break-after",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("break-after", value, "fo");
				_node.SelectSingleNode("@fo:break-after",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the border.
		/// This could be e.g. 0.002cm solid #000000 (width, linestyle, color)
		/// or none
		/// </summary>
		/// <value>The border.</value>
		public string Border
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:border",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:border",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("border", value, "fo");
				_node.SelectSingleNode("@fo:border",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Set the background color
		/// Use Colors.GetColor(Color color) to set 
		/// on of the available .net colors
		/// </summary>
		public string BackgroundColor
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:background-color",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:background-color",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("background-color", value, "fo");
				_node.SelectSingleNode("@fo:background-color",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the padding. 
		/// Default 0.049cm 
		/// Use this in combination with the Border property
		/// </summary>
		/// <value>The padding.</value>
		public string Padding
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:padding",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:padding",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("padding", value, "fo");
				_node.SelectSingleNode("@fo:padding",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the line height in %.
		/// 200% means double space
		/// </summary>
		/// <value>The line height.</value>
		public string LineSpacing
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:line-height",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:line-height",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("line-height", value, "fo");
				_node.SelectSingleNode("@fo:line-height",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		private TabStopStyleCollection _tabstopstylecollection;
		/// <summary>
		/// Gets or sets the tab stop style collection.
		/// <b>Notice:</b> A TabStopStyleCollection will not work
		/// within a Standard Paragraph!
		/// </summary>
		/// <value>The tab stop style collection.</value>
		public TabStopStyleCollection TabStopStyleCollection
		{
			get { return _tabstopstylecollection; }
			set 
			{ 
				if (Style.StyleName == "Standard")
					return;
				if (_tabstopstylecollection != null)
				{
					//Remove node and reset the collection
					Node.RemoveChild(_tabstopstylecollection.Node);
					_tabstopstylecollection = null; 
				}
				
				_tabstopstylecollection	= value;
				if (Node.SelectSingleNode("style:tab-stops", Style.Document.NamespaceManager) == null)
					Node.AppendChild(_tabstopstylecollection.Node);
			}
		}

		/// <summary>
		/// Set paragraph alignment - object.Alignment = TextAlignments.right.ToString()
		/// </summary>
		public string Alignment
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:text-align",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:text-align",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("text-align", value, "fo");
				_node.SelectSingleNode("@fo:text-align",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="ParagraphProperties"/> class.
		/// </summary>
		/// <param name="style">The style.</param>
		public ParagraphProperties(IStyle style)
		{
			Style			= style;
			NewXmlNode();
		}
		
		/// <summary>
		/// Initializes a new instance of the <see cref="ParagraphProperties"/> class.
		/// </summary>
		/// <param name="style">The style.</param>
		/// <param name="node">The node.</param>
		public ParagraphProperties(IStyle style, XmlNode node)
		{
			Style				= style;
			Node				= node;
		}

		/// <summary>
		/// Create a XmlAttribute for propertie XmlNode.
		/// </summary>
		/// <param name="name">The attribute name.</param>
		/// <param name="text">The attribute value.</param>
		/// <param name="prefix">The namespace prefix.</param>
		private void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = Style.Document.CreateAttribute(name, prefix);
			xa.Value		= text;
			Node.Attributes.Append(xa);
		}

		/// <summary>
		/// News the XML node.
		/// </summary>
		private void NewXmlNode()
		{
			Node		= Style.Document.CreateNode("paragraph-properties", "style");
		}

		#region IProperty Member
		private XmlNode _node;
		/// <summary>
		/// The XmlNode which represent the property element.
		/// </summary>
		/// <value>The node</value>
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

		private IStyle _style;
		/// <summary>
		/// The style object to which this property object belongs
		/// </summary>
		/// <value></value>
		public IStyle Style
		{
			get { return _style; }
			set { _style = value; }
		}
		#endregion

		#region IHtmlStyle Member

		/// <summary>
		/// Get the css style fragement
		/// </summary>
		/// <returns>The css style attribute</returns>
		public string GetHtmlStyle()
		{
			string style		= "style=\"";

			if (Alignment != null)
				style	+= "text-align: "+Alignment+"; ";
			if (MarginLeft != null)
				style	+= "text-indent: "+MarginLeft+"; ";
			if (LineSpacing != null)
				style	+= "line-height: "+LineSpacing+"; ";
			if (Border != null && Padding == null)
				style	+= "border-width:1px; border-style:solid; padding: 0.5cm; ";
			if (Border != null && Padding != null)
				style	+= "border-width:1px; border-style:solid; padding:"+Padding+"; ";

			if (!style.EndsWith("; "))
				style	= "";
			else
				style	+= "\"";

			return style;
		}

		#endregion
	}

	/// <summary>
	/// Some helper constants for Paragraph properties
	/// </summary>
	public class ParagraphHelper
	{
		/// <summary>
		/// Line spacing 1.5 lines
		/// </summary>
		public static readonly string LineSpacing15	= "150%";
		/// <summary>
		/// Line spacing double
		/// </summary>
		public static readonly string LineDouble	= "200%";
		/// <summary>
		/// Line spacing three lines
		/// </summary>
		public static readonly string LineSpacing3	= "300%";
	}
}

/*
 * $Log: ParagraphProperties.cs,v $
 * Revision 1.2  2008/04/29 15:39:56  mt
 * new copyright header
 *
 * Revision 1.1  2007/02/25 08:58:54  larsbehr
 * initial checkin, import from Sourceforge.net to OpenOffice.org
 *
 * Revision 1.4  2007/02/13 17:58:49  larsbm
 * - add first part of implementation of master style pages
 * - pdf exporter conversations for tables and images and added measurement helper
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
 * Revision 1.2  2006/02/05 20:03:32  larsbm
 * - Fixed several bugs
 * - clean up some messy code
 *
 * Revision 1.1  2006/01/29 11:28:23  larsbm
 * - Changes for the new version. 1.2. see next changelog for details
 *
 * Revision 1.6  2005/12/12 19:39:17  larsbm
 * - Added Paragraph Header
 * - Added Table Row Header
 * - Fixed some bugs
 * - better whitespace handling
 * - Implmemenation of HTML Exporter
 *
 * Revision 1.5  2005/11/23 19:18:17  larsbm
 * - New Textproperties
 * - New Paragraphproperties
 * - New Border Helper
 * - Textproprtie helper
 *
 * Revision 1.4  2005/11/20 17:31:20  larsbm
 * - added suport for XLinks, TabStopStyles
 * - First experimental of loading dcuments
 * - load and save via importer and exporter interfaces
 *
 * Revision 1.3  2005/10/09 15:52:47  larsbm
 * - Changed some design at the paragraph usage
 * - add list support
 *
 * Revision 1.2  2005/10/08 07:55:35  larsbm
 * - added cvs tags
 *
 */