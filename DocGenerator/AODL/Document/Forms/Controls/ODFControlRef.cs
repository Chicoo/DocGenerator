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
using AODL.Document;
using AODL.Document.Forms;
using System.Xml;
using AODL.Document.TextDocuments;
using AODL.Document.Styles;
using AODL.Document.Content;

namespace AODL.Document.Forms.Controls
{
	/// <summary>
	/// Summary description for ODFControlRef.
	/// </summary>
	public enum AnchorType {Page, Frame, Paragraph, Char, AsChar, NotSet};
    
	public class ODFControlRef: IContent
	{
		protected IDocument _Document;
		protected XmlNode _node;

		public string X
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@svg:x", 
					Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@svg:x", 
					Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("x", "svg"));
				nd.InnerText = value;
			}
		}

		public string Y
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@svg:y", 
					Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@svg:y", 
					Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("y", "svg"));
				nd.InnerText = value;
			}
		}

		public string Width
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@svg:width", 
					Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@svg:width", 
					Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("width", "svg"));
				nd.InnerText = value;
			}
		}
		public string Height
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@svg:height", 
					Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@svg:height", 
					Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("height", "svg"));
				nd.InnerText = value;
			}
		}

		public string Layer
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@draw:layer", 
					Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@draw:layer", 
					Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("layer", "draw"));
				nd.InnerText = value;
			}
		}

		public string ID
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@draw:id", 
					Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@draw:id", 
					Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("id", "draw"));
				nd.InnerText = value;
			}
		}

		public string DrawControl
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@draw:control", 
					Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@draw:control", 
					Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("control", "draw"));
				nd.InnerText = value;
			}
		}

		public string Transform
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@draw:transform", 
					Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@draw:transform", 
					Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("transform", "draw"));
				nd.InnerText = value;
			}
		}

		public int ZIndex
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@svg:z-index", 
					Document.NamespaceManager);
				if (xn == null) return -1;
				return int.Parse(xn.InnerText);
			}
			set
			{
				if (value >0)
				{
					XmlNode nd = _node.SelectSingleNode("@svg:z-index", 
						Document.NamespaceManager);
					if (nd == null)
						nd = Node.Attributes.Append(Document.CreateAttribute("z-index", "svg"));
					nd.InnerText = value.ToString();
				}
				else
				{
					
				}
			}
		}

		public int AnchorPageNumber
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@text:anchor-page-number", 
					Document.NamespaceManager);
				if (xn == null) return -1;
				return int.Parse(xn.InnerText);
			}
			set
			{
				if (value >0)
				{
					XmlNode nd = _node.SelectSingleNode("@text:anchor-page-number", 
						Document.NamespaceManager);
					if (nd == null)
						nd = Node.Attributes.Append(Document.CreateAttribute("anchor-page-number", "text"));
					nd.InnerText = value.ToString();
				}
				else
				{
					
				}
			}
		}

		public AnchorType AnchorType
		{	
			get
			{
				XmlNode xn = _node.SelectSingleNode("@text:anchor-type", 
					Document.NamespaceManager);
				if (xn == null) return AnchorType.NotSet;

				string s = xn.InnerText;	
				AnchorType at;
				switch (s)
				{
					case "page": at = AnchorType.Page; break;
					case "frame": at = AnchorType.Frame; break;
					case "paragraph": at = AnchorType.Paragraph; break;
					case "char": at = AnchorType.Char; break;
					case "as-char": at = AnchorType.AsChar; break;
					default: at = AnchorType.NotSet; break;
				}
				return at;
			}
			set
			{
				string s;
				switch (value)
				{
					case AnchorType.Page: s = "page"; break;
					case AnchorType.Frame: s = "frame"; break;
					case AnchorType.Paragraph: s = "paragraph"; break;
					case AnchorType.Char: s = "char"; break;
					case AnchorType.AsChar: s = "as-char"; break;
					default: return;
				}
				XmlNode nd = _node.SelectSingleNode("@text:anchor-type", 
					Document.NamespaceManager);
				if (nd == null)
					nd = _node.Attributes.Append(Document.CreateAttribute("anchor-type", "text"));
				nd.InnerText = s;
			}
		}

		public XmlBoolean TableBackground
		{	
			get
			{
				XmlNode xn = _node.SelectSingleNode("@table:table-background", 
					Document.NamespaceManager);
				if (xn == null) return XmlBoolean.NotSet;

				string s = xn.InnerText;	
				XmlBoolean at;
				switch (s)
				{
					case "true": at = XmlBoolean.True; break;
					case "false": at = XmlBoolean.False; break;
					default: at = XmlBoolean.NotSet; break;
				}
				return at;
			}
			set
			{
				string s;
				switch (value)
				{
					case XmlBoolean.True: s = "true"; break;
					case XmlBoolean.False: s = "false"; break;
					default: return;
				}
				XmlNode nd = _node.SelectSingleNode("@table:table-background", 
					Document.NamespaceManager);
				if (nd == null)
					nd = _node.Attributes.Append(Document.CreateAttribute("table-background", "table"));
				nd.InnerText = s;
			}
		}



		public IDocument Document 
		{
			get
			{
				return _Document;	
			}
			set
			{
				_Document = value;
			}
		}
		
		
		
		/// <summary>
		/// Represents the XmlNode within the content.xml from the odt file.
		/// </summary>
		/// 
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
		/// The stylename which is referenced with the content object.
		/// If no style is available this is null.
		/// </summary>
		public string StyleName 
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@draw:style-name", 
					Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@draw:style-name", 
					Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("style-name", "draw"));
				nd.InnerText = value;
			}
		}

		/// <summary>
		/// The text style name.
		/// If no style is available this is null.
		/// </summary>
		public string TextStyleName 
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@draw:text-style-name", 
					Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@draw:text-style-name", 
					Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("text-style-name", "draw"));
				nd.InnerText = value;
			}
		}


		/// <summary>
		/// A Style class wich is referenced with the text content object.
		/// If no style is available this is null.
		/// </summary>
		public IStyle TextStyle 
		{
			get
			{
				return _Document.Styles.GetStyleByName(TextStyleName);
			}
			set
			{
				TextStyleName = value.StyleName;
			}
		}

		/// <summary>
		/// A Style class wich is referenced with the content object.
		/// If no style is available this is null.
		/// </summary>
		public IStyle Style 
		{
			get
			{
				return _Document.Styles.GetStyleByName(StyleName);
			}
			set
			{
				StyleName = value.StyleName;
			}
		}

		public ODFControlRef(IDocument doc, XmlNode node)
		{
			_Document = doc;
			Node = node;
		}

		public ODFControlRef(IDocument doc, string draw_control)
		{
			_Document = doc;
			Node = Document.CreateNode("control", "draw");
			DrawControl = draw_control;
		}

		public ODFControlRef(IDocument doc, string draw_control, string x, string y, string width, string height)
		{
			_Document = doc;
			Node = Document.CreateNode("control", "draw");
			DrawControl = draw_control;
			X = x;
			Y = y;
			Width = width;
			Height = height;
		}
			
	}
}
