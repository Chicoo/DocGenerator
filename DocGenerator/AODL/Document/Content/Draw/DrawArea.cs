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
using AODL.Document.Content.OfficeEvents;
using AODL.Document;

namespace AODL.Document.Content.Draw
{
	// @@@@ Add EventListeners to DrawArea
	// @@@@ draw:area-circle attributes: svg:cx, svg:cy, svg:r
	// @@@@ draw:area-polygon:  omit
	// 

	/// <summary>
	/// DrawAreaRectangle represent draw area rectangle which
	/// could be used within a ImageMap.
	/// </summary>
	public class DrawAreaRectangle : DrawArea
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="DrawAreaRectangle"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="node">The node.</param>
		public DrawAreaRectangle(IDocument document, XmlNode node) : base(document)
		{
			Node			= node;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="DrawAreaRectangle"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="x">The x.</param>
		/// <param name="y">The y.</param>
		/// <param name="width">The width.</param>
		/// <param name="height">The height.</param>
		public DrawAreaRectangle(IDocument document,
			string x, string y, string width, string height) : base(document) 
		{
			X = x;
			Y = y;
			Width = width;
			Height = height;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="DrawAreaRectangle"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="x">The x.</param>
		/// <param name="y">The y.</param>
		/// <param name="width">The width.</param>
		/// <param name="height">The height.</param>
		/// <param name="listeners">The listeners.</param>
		public DrawAreaRectangle(IDocument document,
			string x, string y, string width, string height,
			EventListeners listeners)
			: base(document)
		{
			X = x;
			Y = y;
			Width = width;
			Height = height;

			if (listeners != null)
			{
				Content.Add(listeners);
			}
		}

		/// <summary>
		/// Gets or sets the x-position.
		/// </summary>
		/// <value>The description of the area-position.</value>
		public string X
		{
			get
			{
				XmlNode xn = Node.SelectSingleNode("@svg:x",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@svg:x",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("x", value, "svg");
				Node.SelectSingleNode("@svg:x",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the y-position.
		/// </summary>
		/// <value>The y-position of the area-rectangle.</value>
		public string Y
		{
			get
			{
				XmlNode xn = Node.SelectSingleNode("@svg:y",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@svg:y",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("y", value, "svg");
				Node.SelectSingleNode("@svg:y",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the width.
		/// </summary>
		/// <value>The width of the area-rectangle.</value>
		public string Width
		{
			get
			{
				XmlNode xn = Node.SelectSingleNode("@svg:width",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@svg:width",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("width", value, "svg");
				Node.SelectSingleNode("@svg:width",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the height.
		/// </summary>
		/// <value>The height of the area-rectangle.</value>
		public string Height
		{
			get
			{
				XmlNode xn = Node.SelectSingleNode("@svg:height",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@svg:height",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("height", value, "svg");
				Node.SelectSingleNode("@svg:height",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Create a new XmlNode.
		/// </summary>
		override protected void NewXmlNode()
		{			
			Node		= Document.CreateNode("area-rectangle", "draw");
		}
	}

	/// <summary>
	/// Summary for DrawAreaRectangle.
	/// </summary>
	public class DrawAreaCircle : DrawArea
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="DrawAreaCircle"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="node">The node.</param>
		public DrawAreaCircle(IDocument document, XmlNode node) : base(document)
		{
			Node			= node;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="DrawAreaCircle"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="cx">The cx.</param>
		/// <param name="cy">The cy.</param>
		/// <param name="radius">The radius.</param>
		public DrawAreaCircle(IDocument document,
			string cx, string cy, string radius)
			: base(document)
		{
			CX = cx;
			CY = cy;
			Radius = radius;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="DrawAreaCircle"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="cx">The cx.</param>
		/// <param name="cy">The cy.</param>
		/// <param name="radius">The radius.</param>
		/// <param name="listeners">The listeners.</param>
		public DrawAreaCircle(IDocument document,
			string cx, string cy, string radius, EventListeners listeners)
			: base(document)
		{
			CX = cx;
			CY = cy;
			Radius = radius;

			if (listeners != null)
			{
				Content.Add(listeners);
			}
		}

		/// <summary>
		/// Gets or sets the cx-position.
		/// </summary>
		/// <value>The center of the area-circle.</value>
		public string CX
		{
			get
			{
				XmlNode xn = Node.SelectSingleNode("@svg:cx",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				if (value == "")
					return;
				XmlNode xn = Node.SelectSingleNode("@svg:cx",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("cx", value, "svg");
				Node.SelectSingleNode("@svg:cx",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the cy-position.
		/// </summary>
		/// <value>The center position of the area-cicle.</value>
		public string CY
		{
			get
			{
				XmlNode xn = Node.SelectSingleNode("@svg:cy",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				if (value == "")
					return;
				XmlNode xn = Node.SelectSingleNode("@svg:cy",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("cy", value, "svg");
				Node.SelectSingleNode("@svg:cy",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the radius.
		/// </summary>
		/// <value>The radius of the area-circle.</value>
		public string Radius
		{
			get
			{
				XmlNode xn = Node.SelectSingleNode("@svg:r",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				if (value == "")
					return;
				XmlNode xn = Node.SelectSingleNode("@svg:r",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("r", value, "svg");
				Node.SelectSingleNode("@svg:r",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Create a new XmlNode.
		/// </summary>
		override protected void NewXmlNode()
		{
			Node = Document.CreateNode("area-circle", "draw");
		}
	}

	/// <summary>
	/// Summary for DrawArea.
	/// </summary>
	abstract public class DrawArea: IContent, IContentContainer
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
					Document.NamespaceManager);
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
					Document.NamespaceManager);
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
		/// Gets or sets the description.
		/// </summary>
		/// <value>The description of the draw area.</value>
		public string Description
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@draw:desc",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:desc",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("desc", value, "draw");
				_node.SelectSingleNode("@draw:desc",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Create a XmlAttribute for propertie XmlNode.
		/// </summary>
		/// <param name="name">The attribute name.</param>
		/// <param name="text">The attribute value.</param>
		/// <param name="prefix">The namespace prefix.</param>
		protected void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = Document.CreateAttribute(name, prefix);
			xa.Value = text;
			Node.Attributes.Append(xa);
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="DrawArea"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		protected DrawArea(IDocument document)
		{
			Document = document;
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
		/// Initializes a new instance of the <see cref="DrawAreaRectangle"/> class.
		/// </summary>
		abstract protected void NewXmlNode();

		/// <summary>
		/// Content_s the inserted.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		protected void Content_Inserted(int index, object value)
		{
			if (Node != null)
				Node.AppendChild(((IContent)value).Node);
		}

		/// <summary>
		/// Content_s the removed.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		protected void Content_Removed(int index, object value)
		{
			Node.RemoveChild(((IContent)value).Node);
		}

		#region IContent Member		

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

		/// <summary>
		/// A draw:area-rectangle doesn't have a style-name.
		/// </summary>
		/// <value>The name</value>
		public string StyleName
		{
			get
			{
				return null;
			}
			set
			{
			}
		}

		/// <summary>
		/// A draw:area-rectangle doesn't have a style.
		/// </summary>
		public IStyle Style
		{
			get
			{
				return null;
			}
			set
			{
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

		#region IContentCollection Member

		private ContentCollection _content;
		/// <summary>
		/// Gets or sets the content collection.
		/// </summary>
		/// <value>The content collection.</value>
		public ContentCollection Content
		{
			get { return _content; }
			set { _content = value; }
		}

		#endregion

		#region IContent Member old

//		private IDocument _document;
//		/// <summary>
//		/// The IDocument to which this draw-area is bound.
//		/// </summary>
//		public IDocument Document
//		{
//			get
//			{
//				return this._document;
//			}
//			set
//			{
//				this._document = value;
//			}
//		}
//
//		private XmlNode _node;
//		/// <summary>
//		/// The XmlNode.
//		/// </summary>
//		public XmlNode Node
//		{
//			get
//			{
//				return this._node;
//			}
//			set
//			{
//				this._node = value;
//			}
//		}
//
//		/// <summary>
//		/// A draw:area-rectangle doesn't have a style-name.
//		/// </summary>
//		/// <value>The name</value>
//		public string Stylename
//		{
//			get
//			{
//				return null;
//			}
//			set
//			{
//			}
//		}
//
//		/// <summary>
//		/// A draw:area-rectangle doesn't have a style.
//		/// </summary>
//		public IStyle Style
//		{
//			get
//			{
//				return null;
//			}
//			set
//			{
//			}
//		}
//
//		/// <summary>
//		/// A draw:area-rectangle doesn't contain text.
//		/// </summary>
//		public ITextCollection TextContent
//		{
//			get
//			{
//				return null;
//			}
//			set
//			{
//			}
//		}

		#endregion
	
	}
}
