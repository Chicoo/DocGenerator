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
	/// Represent the properties of the list levels.
	/// </summary>
	public class ListLevelProperties : IProperty
	{
		/// <summary>
		/// Gets or sets the space before.
		/// </summary>
		/// <value>The space before.</value>
		public string SpaceBefore
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@text:space-before", 
					Style.Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@text:space-before",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("space-before", value, "text");
				_node.SelectSingleNode("@text:space-before",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the width of the min label.
		/// </summary>
		/// <value>The width of the min label.</value>
		public string MinLabelWidth
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@text:min-label-width", 
					Style.Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@text:min-label-width",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("min-label-width", value, "text");
				_node.SelectSingleNode("@text:min-label-width",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Constructor create a new ListLevelProperties object.
		/// </summary>
		public ListLevelProperties(IStyle style)
		{
			Style				= style;
			NewXmlNode();
		}

		/// <summary>
		/// Create the XmlNode which represent the propertie element.
		/// </summary>
		private void NewXmlNode()
		{
			Node			= Style.Document.CreateNode("list-level-properties", "style");
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
		/// The XmlNode.
		/// </summary>
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
		/// The style to this ListLevelProperties object belongs.
		/// </summary>
		public IStyle Style
		{
			get { return _style; }
			set { _style = value; }
		}

		#endregion
	}
}
