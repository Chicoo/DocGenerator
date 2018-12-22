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
using System.Xml ;
using AODL.Document .Styles ;
using AODL.Document .Collections ;
using AODL.Document .Content .Text ;

namespace AODL.Document.Content.Charts
{
	/// <summary>
	/// Summary description for ChartTitle.
	/// </summary>
	public class ChartTitle:IContent, IContentContainer
	{
        /// <summary>
        /// the chart which include the charttitle
        /// </summary>
		private Chart _chart;

		public Chart Chart
		{
			get { return _chart; }
			set { _chart = value; }
		}

		/// <summary>
		/// Gets or sets the title style.
		/// </summary>
		/// <value>The title style.</value>
		public TitleStyle TitleStyle
		{
			get
			{
				return (TitleStyle)Style;
			}
			set
			{
				StyleName	= ((TitleStyle)value).StyleName;
				Style = value;
			}
		}
		/// <summary>
		/// Gets or sets the horizontal position where
		/// the title should be
		/// anchored. 
		/// </summary>

		public string SvgX
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@svg:x",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@svg:x",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("x", value, "svg");
				_node.SelectSingleNode("@svg:x",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the vertical position where
		/// the title should be
		/// anchored. 
		/// </summary>

		public string SvgY
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@svg:y",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@svg:y",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("y", value, "svg");
				_node.SelectSingleNode("@svg:y",
					Document.NamespaceManager).InnerText = value;
			}
		}


		/// <summary>
		/// Initializes a new instance of the charttitle class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="node">The node.</param>
		
		public ChartTitle(IDocument document, XmlNode node)
		{
			Document			= document;
			Node				= node;
			InitStandards();
		}

		/// <summary>
		/// Initializes a new instance of the charttitle class.
		/// </summary>
		/// <param name="table">The table.</param>
		public ChartTitle(Chart chart)
		{
			Chart				= chart;
			Document			= chart.Document;
			NewXmlNode(null);
			InitStandards();
			TitleStyle = new TitleStyle (chart.Document ,null);
			Chart .Styles .Add (TitleStyle );
			//chart.Content .Add (this);
		}

		/// <summary>
		/// Initializes a new instance of the charttitle class.
		/// </summary>
		/// <param name="table">The table.</param>
		/// <param name="styleName">The style name.</param>
		public ChartTitle(Chart chart, string styleName)
		{
			Chart				= chart;
			Document			= chart.Document;
			NewXmlNode(styleName);
			InitStandards();

			if (styleName != null)
			{
				StyleName		= styleName;
				TitleStyle		= new TitleStyle(Document, styleName);
				Chart.Styles.Add(TitleStyle);
			}

			//chart.Content .Add (this);
		}

		/// <summary>
		/// Inits the standards.
		/// </summary>
		private void InitStandards()
		{
			Content			= new ContentCollection();
			Content.Inserted	+= Content_Inserted;
			Content.Removed	+= Content_Removed;
		}

		/// <summary>
		/// Init the title's default text
		/// </summary>

		public void InitTitle()
		{
			Paragraph para = new Paragraph (Chart.Document );
			para.TextContent .Add (new SimpleText (Document ,"Ö÷±êÌâ"));
			Content .Add (para);
		}

		/// <summary>
		/// set the title
		/// </summary>
		/// <param name="styleName"></param>
        
		public void SetTitle(string Title)
		{
			if (Content .Count!=0)			  
			{
				 for(int i=0 ;i<Content .Count;i++)
				    Content .RemoveAt(i);
			}
			Paragraph para = new Paragraph (Chart.Document );
			para.TextContent .Add (new SimpleText (Document ,Title));
			
			if (Chart .Document.IsLoadedFile&&!Chart .IsNewed)
			{
				XmlNode node = Chart .ChartDoc .ImportNode (para.Node,true );
				Node .AppendChild (node);
			}

			else
			{
				Content .Add (para);
			}
			
			
		}


		private void NewXmlNode(string styleName)
		{
			Node		= Document .CreateNode("title", "chart");

			XmlAttribute xa = Document.CreateAttribute("style-name", "chart");
			xa.Value		= styleName;
			Node.Attributes.Append(xa);
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
				XmlNode xn = _node.SelectSingleNode("@chart:style-name",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:style-name",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("style-name", value, "chart");
				_node.SelectSingleNode("@chart:style-name",
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
		/// Gets or sets the content.
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
		/// Content_s the inserted.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void Content_Inserted(int index, object value)
		{
			Node.AppendChild(((IContent)value).Node);
		}

		/// <summary>
		/// Content_s the removed.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void Content_Removed(int index, object value)
		{
			Node.RemoveChild(((IContent)value).Node);
		}

		private void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = Document.CreateAttribute(name, prefix);
			xa.Value		= text;
			Node.Attributes.Append(xa);
		}




	}
}

