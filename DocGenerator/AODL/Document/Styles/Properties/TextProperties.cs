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
	/// Represent access to the possible attributes of of an text-propertie element.
	/// </summary>
	public class TextProperties : IProperty, IHtmlStyle
	{
		/// <summary>
		/// Set font-weight bold object.Bold = "bold";
		/// </summary>
		public string Bold
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:font-weight",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:font-weight",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("font-weight", value, "fo");
				_node.SelectSingleNode("@fo:font-weight",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Set text position, could be sub- ("sub 58%") or superscript ("super 58%")
		/// See TextPropertieHelper for possible settings
		/// </summary>
		public string Position
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:text-position",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@style:text-position",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("text-position", value, "style");
				_node.SelectSingleNode("@style:text-position",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Set text shadow e.g. "1pt 1pt"
		/// See TextPropertieHelper for possible settings
		/// </summary>
		public string Shadow
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:text-shadow",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@style:text-shadow",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("text-shadow", value, "style");
				_node.SelectSingleNode("@style:text-shadow",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Set text outline e.g. "true"
		/// </summary>
		public string Outline
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:text-outline",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@style:text-outline",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("text-outline", value, "style");
				_node.SelectSingleNode("@style:text-outline",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Set text to line through
		/// See LineStyles for possible settings
		/// </summary>
		public string TextLineThrough
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:text-line-through-style",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@style:text-line-through-style",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("text-line-through-style", value, "style");
				_node.SelectSingleNode("@style:text-line-through-style",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Set the font color
		/// Use Colors.GetColor(Color color) to set 
		/// on of the available .net colors
		/// </summary>
		public string FontColor
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:color",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:color",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("color", value, "fo");
				_node.SelectSingleNode("@fo:color",
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
		/// Set font-style italic object.Italic = "italic";
		/// </summary>
		public string Italic
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:font-style",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:font-style",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("font-style", value, "fo");
				_node.SelectSingleNode("@fo:font-style",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Set text-underline-style Underline object.Underline = "dotted";
		/// </summary>
		public string Underline
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:text-underline-style",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@style:text-underline-style",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("text-underline-style", value, "style");
				_node.SelectSingleNode("@style:text-underline-style",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Set text-underline-style Underline object.Underline = "dotted";
		/// </summary>
		public string UnderlineWidth
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:text-underline-width",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@style:text-underline-width",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("text-underline-width", value, "style");
				_node.SelectSingleNode("@style:text-underline-width",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Set text-underline-color - object.UnderlineColor = "font-color";
		/// </summary>
		public string UnderlineColor
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:text-underline-color",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@style:text-underline-color",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("text-underline-color", value, "style");
				_node.SelectSingleNode("@style:text-underline-color",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Set font-name you will find all available fonts in class FontFamilies
		/// </summary>
		public string FontName
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:font-name",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				Style.Document.FontList.Add(value);
				XmlNode xn = _node.SelectSingleNode("@style:font-name",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("font-name", value, "style");
				_node.SelectSingleNode("@style:font-name",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Set font-name - object.FontSize = "10pt";
		/// </summary>
		public string FontSize
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:font-size",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:font-size",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("font-size", value, "fo");
				_node.SelectSingleNode("@fo:font-size",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Create a new TextProperties object that belongs to the given TextStyle object.
		/// </summary>
		/// <param name="style">The TextStyle object.</param>
		public TextProperties(IStyle style)
		{
			Style				= style;
			NewXmlNode();
		}

		/// <summary>
		/// Use to set all possible underline styles. You can use this method
		/// instead to set all necessary properties by hand.
		/// </summary>
		/// <param name="style">The style. Sess enum LineStyles.</param>
		/// <param name="width">The width. Sess enum LineWidths.</param>
		/// <param name="color">The color. "font-color" or as hex #[0-9a-fA-F]{6}</param>
		public void SetUnderlineStyles(string style, string width, string color)
		{
			Underline		= style;
			UnderlineWidth	= width;
			UnderlineColor	= color;
		}

		/// <summary>
		/// Create the XmlNode which represent the propertie element.
		/// </summary>
		private void NewXmlNode()
		{
			Node		= Style.Document.CreateNode("text-properties", "style");
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

			if (Italic != null)
				if (Italic != "normal")
					style	+= "font-style: italic; ";
			if (Bold != null)
				style	+= "font-weight: "+Bold+"; ";
			if (Underline != null)
				style	+= "text-decoration: underline; ";
			if (TextLineThrough != null)
				style	+= "text-decoration: line-through; ";
			if (FontName != null)
				style	+= "font-family: "+FontFamilies.HtmlFont(FontName)+"; ";
			if (FontSize != null)
				style	+= "font-size: "+FontFamilies.PtToPx(FontSize)+"; ";
			if (FontColor != null)
				style	+= "color: "+FontColor+"; ";
			if (BackgroundColor != null)
				style	+= "background-color: "+BackgroundColor+"; ";

			if (!style.EndsWith("; "))
				style	= "";
			else
				style	+= "\"";

			return style;
		}

		#endregion
	}

	/// <summary>
	/// Propertie Helpers
	/// </summary>
	public class TextPropertieHelper
	{
		/// <summary>
		/// Subscript use within Postion
		/// </summary>
		public static readonly string Subscript		= "sub 58%";
		/// <summary>
		/// Superscript use within Postion
		/// </summary>
		public static readonly string Superscript	= "super 58%";
		/// <summary>
		/// Light shadow use within Shadow
		/// </summary>
		public static readonly string Shadowlight	= "1pt 1pt";
		/// <summary>
		/// Middle shadow use within Shadow
		/// </summary>
		public static readonly string Shadowmidlle	= "3pt 3pt";
		/// <summary>
		/// Heavy shadow use within Shadow
		/// </summary>
		public static readonly string Shadowheavy	= "6pt 6pt";
	}
}

/*
 * $Log: TextProperties.cs,v $
 * Revision 1.2  2008/04/29 15:39:56  mt
 * new copyright header
 *
 * Revision 1.1  2007/02/25 08:58:57  larsbehr
 * initial checkin, import from Sourceforge.net to OpenOffice.org
 *
 * Revision 1.1  2006/01/29 11:28:23  larsbm
 * - Changes for the new version. 1.2. see next changelog for details
 *
 * Revision 1.9  2006/01/05 10:31:10  larsbm
 * - AODL merged cells
 * - AODL toc
 * - AODC batch mode, splash screen
 *
 * Revision 1.8  2005/12/12 19:39:17  larsbm
 * - Added Paragraph Header
 * - Added Table Row Header
 * - Fixed some bugs
 * - better whitespace handling
 * - Implmemenation of HTML Exporter
 *
 * Revision 1.7  2005/11/23 19:18:17  larsbm
 * - New Textproperties
 * - New Paragraphproperties
 * - New Border Helper
 * - Textproprtie helper
 *
 * Revision 1.6  2005/10/23 16:47:48  larsbm
 * - Bugfix ListItem throws IStyleInterface not implemented exeption
 * - now. build the document after call saveto instead prepare the do at runtime
 * - add remove support for IText objects in the paragraph class
 *
 * Revision 1.5  2005/10/22 15:52:10  larsbm
 * - Changed some styles from Enum to Class with statics
 * - Add full support for available OpenOffice fonts
 *
 * Revision 1.4  2005/10/08 09:01:15  larsbm
 * --- uncommented ---
 *
 * Revision 1.3  2005/10/08 08:07:45  larsbm
 * - added cvs tags
 *
 * Revision 1.2  2005/10/08 07:55:35  larsbm
 * - added cvs tags
 *
 */
