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

namespace AODL.Document.Content.Charts
{
	/// <summary>
	/// Summary description for ChartLegend.
	/// </summary>
	public class ChartLegend: IContent
	{
		
		/// <summary>
		/// Gets or sets the horizontal position 
		/// where the legend should be anchored
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
		/// the legend should be
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
		/// the chart which contains the legend
		/// </summary>

		private Chart _chart;

		public Chart Chart
		{
			get    
			{ 
				return _chart;
			}
			set    
			{
				_chart=value;
			}
		}

		/// <summary>
		/// gets and sets the style of the legend
		/// </summary>

		public LegendStyle LegendStyle
		{
			get
			{
				return (LegendStyle)Style;
			}
			set
			{
				StyleName	= ((LegendStyle)value).StyleName;
				Style = value;
			}
		}

		/// <summary>
		/// gets and sets the  position of the legend
		/// </summary>


		public string LegendPosition
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:legend-position",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:legend-position",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("legend-position", value, "chart");
				_node.SelectSingleNode("@chart:legend-position",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the align of the legend
		/// </summary>

		public string LegendAlign
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:legend-align",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:legend-align",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("legend-align", value, "chart");
				_node.SelectSingleNode("@chart:legend-align",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// the consturctor of the chart legend
		/// </summary>
		/// <param name="document"></param>
		/// <param name="node"></param>

		public ChartLegend(IDocument document,XmlNode node)
		{
			Document =document;
			Node =node;
		}

		public ChartLegend(Chart chart)
		{
			Chart =chart;
			Document =chart.Document;			
			NewXmlNode (null);
			LegendStyle = new LegendStyle (chart.Document); 

			chart.Content .Add (this);
		}

		public ChartLegend(Chart chart,string styleName)
		{
			Chart =chart;
			Document =chart.Document;			
			NewXmlNode (styleName);
			LegendStyle = new LegendStyle (Document,styleName);
			Chart .Styles .Add (LegendStyle );

			chart.Content .Add (this);
		}



		private void NewXmlNode(string styleName)
		{
			Node =Document.CreateNode("legend","chart");
			XmlAttribute xa=Document.CreateAttribute ("style-name","chart");
			xa.Value  =styleName;
			Node.Attributes .Append (xa);

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

		private void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = Document.CreateAttribute(name, prefix);
			xa.Value		= text;
			Node.Attributes.Append(xa);
		}

		
	}
}

