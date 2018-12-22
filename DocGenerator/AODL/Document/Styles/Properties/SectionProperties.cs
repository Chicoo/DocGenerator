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
	/// SectionProperties represent the section properties which is e.g used
	/// within a table of contents.
	/// </summary>
	public class SectionProperties : IProperty
	{
		/// <summary>
		/// Gets or sets a value indicating whether this <see cref="SectionProperties"/> is editable.
		/// </summary>
		/// <value><c>true</c> if editable; otherwise, <c>false</c>.</value>
		public bool Editable
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:editable",
					Style.Document.NamespaceManager);
				if (xn != null)
					return Convert.ToBoolean(xn.InnerText);
				return false;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@style:editable",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("editable", value.ToString(), "style");
				_node.SelectSingleNode("@style:editable",
					Style.Document.NamespaceManager).InnerText = value.ToString();
			}
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="SectionProperties"/> class.
		/// </summary>
		/// <param name="style">The style.</param>
		public SectionProperties(IStyle style)
		{
			Style					= style;
			NewXmlNode();
		}

		/// <summary>
		/// Adds the standard column style.
		/// While creating new TableOfContent objects
		/// AODL will only support a TableOfContent
		/// which use the Header styles with outlining
		/// without table columns
		/// </summary>
		public void AddStandardColumnStyle()
		{
			XmlNode standardColStyle	= Style.Document.CreateNode("columns", "style");

			XmlAttribute xa				= Style.Document.CreateAttribute("column-count", "fo");
			xa.Value					= "0";
			standardColStyle.Attributes.Append(xa);

			xa							= Style.Document.CreateAttribute("column-gap", "fo");
			xa.Value					= "0cm";
			standardColStyle.Attributes.Append(xa);

			Node.AppendChild(standardColStyle);
		}

		/// <summary>
		/// Create a new XmlNode.
		/// </summary>
		private void NewXmlNode()
		{			
			Node		= Style.Document.CreateNode("section-properties", "style");

			XmlAttribute xa = Style.Document.CreateAttribute("editable", "style");
			xa.Value		= "false";
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
	}
}

/*
 * $Log: SectionProperties.cs,v $
 * Revision 1.2  2008/04/29 15:39:56  mt
 * new copyright header
 *
 * Revision 1.1  2007/02/25 08:58:55  larsbehr
 * initial checkin, import from Sourceforge.net to OpenOffice.org
 *
 * Revision 1.2  2006/02/05 20:03:32  larsbm
 * - Fixed several bugs
 * - clean up some messy code
 *
 * Revision 1.1  2006/01/29 11:28:23  larsbm
 * - Changes for the new version. 1.2. see next changelog for details
 *
 * Revision 1.1  2006/01/05 10:31:10  larsbm
 * - AODL merged cells
 * - AODL toc
 * - AODC batch mode, splash screen
 *
 */