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
using AODL.Document.Content;
using AODL.Document.Content.Text;
using AODL.Document.Content.Tables;
using AODL.Document.Styles;
using AODL.Document.TextDocuments;
using AODL.Document.Helper;
using System.Collections;
using System.Xml;

namespace AODL.Document.Styles.MasterStyles
{
	/// <summary>
	/// Summary for TextPageHeaderFooterBase.
	/// </summary>
	public class TextPageHeaderFooterBase
	{
		private TextMasterPage _textMasterPage;
		/// <summary>
		/// Gets or sets the owner text master page.
		/// </summary>
		/// <value>The text master page.</value>
		public TextMasterPage TextMasterPage
		{
			get { return _textMasterPage; }
			set { _textMasterPage = value; }
		}
		
		private TextDocument _textDocument;
		/// <summary>
		/// Gets or sets the text document.
		/// </summary>
		/// <value>The text document.</value>
		public TextDocument TextDocument
		{
			get { return _textDocument; }
			set { _textDocument = value; }
		}
		
		private XmlNode _propertyNode;
		/// <summary>
		/// Gets or sets the property node.
		/// </summary>
		/// <value>The property node.</value>
		public XmlNode PropertyNode
		{
			get
			{
				return _propertyNode;
			}
			set
			{
				_propertyNode = value;
			}
		}
		
		private XmlNode _contentNode;
		/// <summary>
		/// Gets or sets the content node.
		/// </summary>
		/// <value>The content node.</value>
		public XmlNode ContentNode
		{
			get
			{
				return _contentNode;
			}
			set
			{
				_contentNode = value;
			}
		}

		private XmlNode _styleNode;
		/// <summary>
		/// The XmlNode which represent the page layout style element.
		/// </summary>
		/// <value>The node</value>
		public XmlNode StyleNode
		{
			get
			{
				return _styleNode;
			}
			set
			{
				_styleNode = value;
			}
		}

		/// <summary>
		/// Gets or sets the min height. e.g. 0cm 
		/// </summary>
		/// <value>The height of the min.</value>
		public string MinHeight
		{
			get 
			{ 
				XmlAttribute xmlAttr = _propertyNode.Attributes["min-height", "fo"];
				if (xmlAttr != null)
					return xmlAttr.InnerText;
				return null;
			}
			set
			{
				XmlAttribute xmlAttr = _propertyNode.Attributes["min-height", "fo"];
				if (xmlAttr == null)
					CreateAttribute("min-height", value, "fo");
			}
		}

		/// <summary>
		/// Gets or sets the margin right. e.g. 0cm
		/// </summary>
		/// <value>The margin right.</value>
		public string MarginRight
		{
			get 
			{ 
				XmlAttribute xmlAttr = _propertyNode.Attributes["margin-right", "fo"];
				if (xmlAttr != null)
					return xmlAttr.InnerText;
				return null;
			}
			set
			{
				XmlAttribute xmlAttr = _propertyNode.Attributes["margin-right", "fo"];
				if (xmlAttr == null)
					CreateAttribute("margin-right", value, "fo");
			}
		}

		/// <summary>
		/// Gets or sets the margin bottom. e.g. 0.499cm
		/// </summary>
		/// <value>The margin bottom.</value>
		public string MarginBottom
		{
			get 
			{ 
				XmlAttribute xmlAttr = _propertyNode.Attributes["margin-bottom", "fo"];
				if (xmlAttr != null)
					return xmlAttr.InnerText;
				return null;
			}
			set
			{
				XmlAttribute xmlAttr = _propertyNode.Attributes["margin-bottom", "fo"];
				if (xmlAttr == null)
					CreateAttribute("margin-bottom", value, "fo");
			}
		}

		/// <summary>
		/// Gets or sets the margin top. e.g. 0.499cm
		/// </summary>
		/// <value>The margin bottom.</value>
		public string MarginTop
		{
			get 
			{ 
				XmlAttribute xmlAttr = _propertyNode.Attributes["margin-top", "fo"];
				if (xmlAttr != null)
					return xmlAttr.InnerText;
				return null;
			}
			set
			{
				XmlAttribute xmlAttr = _propertyNode.Attributes["margin-top", "fo"];
				if (xmlAttr == null)
					CreateAttribute("margin-top", value, "fo");
			}
		}

		/// <summary>
		/// Gets or sets the margin left. e.g. 0cm
		/// </summary>
		/// <value>The margin left.</value>
		public string MarginLeft
		{
			get 
			{ 
				XmlAttribute xmlAttr = _propertyNode.Attributes["margin-top", "fo"];
				if (xmlAttr != null)
					return xmlAttr.InnerText;
				return null;
			}
			set
			{
				XmlAttribute xmlAttr = _propertyNode.Attributes["margin-left", "fo"];
				if (xmlAttr == null)
					CreateAttribute("margin-left", value, "fo");
			}
		}

		private ContentCollection _contentCollection;
		/// <summary>
		/// Gets or sets the content collection.
		/// </summary>
		/// <value>The content collection.</value>
		/// <remarks>This is the content which will be displayed within the footer or the header.</remarks>
		public ContentCollection ContentCollection
		{
			get { return _contentCollection; }
			set { _contentCollection = value; }
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="TextPageFooter"/> class.
		/// </summary>
		protected TextPageHeaderFooterBase()
		{
			_contentCollection = new ContentCollection();
			_contentCollection.Inserted += _contentCollection_Inserted;
			_contentCollection.Removed += _contentCollection_Removed;
		}

		/// <summary>
		/// Create a new TextPageFooter object.
		/// !!NOTICE: The TextPageLayout of the TextMasterPage object must exist!
		/// </summary>
		/// <param name="textMasterPage">The text master page.</param>
		/// <param name="typeName">Name of the type to create header or footer.</param>
		/// <remarks>The TextPageLayout of the TextMasterPage object must exist!</remarks>
		protected void New(TextMasterPage textMasterPage, string typeName)
		{
			try
			{
				_textMasterPage = textMasterPage;
				_textDocument = textMasterPage.TextDocument;
				_styleNode = TextDocument.CreateNode(
					"footer-style", "style");
				_textMasterPage.TextPageLayout.StyleNode.AppendChild(_styleNode);
			}
			catch(Exception)
			{
				throw;
			}
		}

		/// <summary>
		/// This method call will activate resp. create the header or footer 
		/// for a master page. There is no need of call it directly.
		/// Activation for footer and header will be done through
		/// the TextMasterPage object. 
		/// </summary>
		public void Activate()
		{
			string typeName = (this is TextPageHeader) ? "header" : "footer";

			// only if the content node doesn't exist
			if (_contentNode == null)
			{
				_contentNode = TextDocument.CreateNode(
					typeName, "style");
				_textMasterPage.Node.AppendChild(_contentNode);
			}

			// only if the property node doesn't exist
			if (_propertyNode == null)
			{
				_propertyNode = TextDocument.CreateNode(
					"header-footer-properties", "style");
				// Set defaults
				MarginLeft = "0cm";
				MarginRight = "0cm";
				MinHeight = "0cm";
				MarginBottom = (typeName.Equals("header")) ? "0.499cm" : "0cm";
				MarginTop = (typeName.Equals("footer")) ? "0.499cm" : "0cm";
				_styleNode.AppendChild(_propertyNode);
			}
		}

		/// <summary>
		/// Create a XmlAttribute for propertie XmlNode.
		/// </summary>
		/// <param name="name">The attribute name.</param>
		/// <param name="text">The attribute value.</param>
		/// <param name="prefix">The namespace prefix.</param>
		private void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = TextDocument.CreateAttribute(name, prefix);
			xa.Value		= text;
			PropertyNode.Attributes.Append(xa);
		}

		/// <summary>
		/// Called after content was added.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void _contentCollection_Inserted(int index, object value)
		{
			ContentNode.AppendChild(((IContent)value).Node);
		}

		/// <summary>
		/// Called after content was removed.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void _contentCollection_Removed(int index, object value)
		{
			ContentNode.RemoveChild(((IContent)value).Node);
		}
	}
}
