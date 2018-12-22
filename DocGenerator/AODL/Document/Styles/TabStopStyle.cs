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
using AODL.Document.Styles.Properties;

namespace AODL.Document.Styles
{
	/// <summary>
	/// Class represent a TabStopStyle.
	/// </summary>
	public class TabStopStyle : AbstractStyle
	{
		/// <summary>
		/// Position e.g = "4.98cm";
		/// </summary>
		public string Position
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:position", 
					Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:position",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("postion", value, "style");
				_node.SelectSingleNode("@style:position",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// A Tabstoptype e.g center
		/// </summary>
		public string Type
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:type", 
					Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:type",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("type", value, "style");
				_node.SelectSingleNode("@style:type",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// The Tabstop LeaderStyle e.g dotted
		/// </summary>
		public string LeaderStyle
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:leader-style", 
					Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:leader-style",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("leader-style", value, "style");
				_node.SelectSingleNode("@style:leader-style",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// The Tabstop Leader text e.g. "."
		/// Use this if you use the LeaderStyle property
		/// </summary>
		public string LeaderText
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:leader-text", 
					Document.NamespaceManager) ;
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:leader-text",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("leader-text", value, "style");
				_node.SelectSingleNode("@style:leader-text",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="TabStopStyle"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="position">The position.</param>
		public TabStopStyle(IDocument document, double position)
		{
			Document		= document;
			NewXmlNode(position);
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="TabStopStyle"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="node">The node.</param>
		public TabStopStyle(IDocument document, XmlNode node)
		{
			Document		= document;
			Node			= node;
		}

		/// <summary>
		/// Create the XmlNode that represent this element.
		/// </summary>
		/// <param name="position">The position.</param>
		private void NewXmlNode(double position)
		{			
			Node		= Document.CreateNode("tab-stop", "style");

			XmlAttribute xa = Document.CreateAttribute("position", "style");
			xa.Value		= position.ToString().Replace(",",".")+"cm";
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
	}
}

/*
 * $Log: TabStopStyle.cs,v $
 * Revision 1.2  2008/04/29 15:39:54  mt
 * new copyright header
 *
 * Revision 1.1  2007/02/25 08:58:50  larsbehr
 * initial checkin, import from Sourceforge.net to OpenOffice.org
 *
 * Revision 1.1  2006/01/29 11:28:23  larsbm
 * - Changes for the new version. 1.2. see next changelog for details
 *
 * Revision 1.1  2005/11/20 17:31:20  larsbm
 * - added suport for XLinks, TabStopStyles
 * - First experimental of loading dcuments
 * - load and save via importer and exporter interfaces
 *
 */