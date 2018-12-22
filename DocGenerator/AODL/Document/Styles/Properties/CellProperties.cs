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
using AODL.Document;
using AODL.Document.Content;

namespace AODL.Document.Styles.Properties
{
	/// <summary>
	/// Represent the Cell Properties within a Cell which is used
	/// within a Tabe resp. a Row.
	/// </summary>
	public class CellProperties :IProperty, IHtmlStyle
	{
		/// <summary>
		/// Gets or sets the cell style.
		/// </summary>
		/// <value>The cell style.</value>
		public CellStyle CellStyle
		{
			get { return (CellStyle)Style; }
			set { Style = value; }
		}

		/// <summary>
		/// Gets or sets the padding. 
		/// Default 0.097cm
		/// </summary>
		/// <value>The padding.</value>
		public string Padding
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:padding",
					CellStyle.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:padding",
					CellStyle.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("padding", value, "fo");
				_node.SelectSingleNode("@fo:padding",
					CellStyle.Document.NamespaceManager).InnerText = value;
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
					CellStyle.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:border",
					CellStyle.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("border", value, "fo");
				_node.SelectSingleNode("@fo:border",
					CellStyle.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the border left.
		/// This could be e.g. 0.002cm solid #000000 (width, linestyle, color)
		/// or none
		/// </summary>
		/// <value>The border left.</value>
		public string BorderLeft
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:border-left",
					CellStyle.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:border-left",
					CellStyle.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("border-left", value, "fo");
				_node.SelectSingleNode("@fo:border-left",
					CellStyle.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the border right.
		/// This could be e.g. 0.002cm solid #000000 (width, linestyle, color)
		/// or none
		/// </summary>
		/// <value>The border right.</value>
		public string BorderRight
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:border-right",
					CellStyle.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:border-right",
					CellStyle.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("border-right", value, "fo");
				_node.SelectSingleNode("@fo:border-right",
					CellStyle.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the border top.
		/// This could be e.g. 0.002cm solid #000000 (width, linestyle, color)
		/// or none
		/// </summary>
		/// <value>The border top.</value>
		public string BorderTop
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:border-top",
					CellStyle.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:border-top",
					CellStyle.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("border-top", value, "fo");
				_node.SelectSingleNode("@fo:border-top",
					CellStyle.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the border bottom.
		/// This could be e.g. 0.002cm solid #000000 (width, linestyle, color)
		/// or none
		/// </summary>
		/// <value>The border bottom.</value>
		public string BorderBottom
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:border-bottom",
					CellStyle.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:border-bottom",
					CellStyle.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("border-bottom", value, "fo");
				_node.SelectSingleNode("@fo:border-bottom",
					CellStyle.Document.NamespaceManager).InnerText = value;
			}
		}
		
		/// <summary>
		/// Gets or sets the color of the background. e.g #000000 for black
		/// </summary>
		/// <value>The color of the background.</value>
		public string BackgroundColor
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@fo:background-color",
					CellStyle.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@fo:background-color",
					CellStyle.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("background-color", value, "fo");
				_node.SelectSingleNode("@fo:background-color",
					CellStyle.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="CellProperties"/> class.
		/// </summary>
		/// <param name="cellstyle">The cellstyle.</param>
		public CellProperties(CellStyle cellstyle)
		{
			CellStyle		= cellstyle;
			NewXmlNode();
			//TODO: Check localisations cm?? inch??
			//defaults 
			Padding		= "0.097cm";
		}

		/// <summary>
		/// Create the XmlNode which represent the propertie element.
		/// </summary>
		private void NewXmlNode()
		{
			Node		= Style.Document.CreateNode("table-cell-properties", "style");
		}

		/// <summary>
		/// Create a XmlAttribute for propertie XmlNode.
		/// </summary>
		/// <param name="name">The attribute name.</param>
		/// <param name="text">The attribute value.</param>
		/// <param name="prefix">The namespace prefix.</param>
		private void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = CellStyle.Document.CreateAttribute(name, prefix);
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

			if (BackgroundColor != null)
				if (BackgroundColor.ToLower() != "transparent")
					style	+= "background-color: "+BackgroundColor+"; ";
				else
					style	+= "background-color: #FFFFFF; ";
			else
				style	+= "background-color: #FFFFFF; ";

			if (!style.EndsWith("; "))
				style	= "";
			else
				style	+= "\"";

			return style;
		}

		#endregion

	}
}
