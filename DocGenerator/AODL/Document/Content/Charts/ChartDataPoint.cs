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
using AODL.Document .Styles ;
using AODL.Document ;
using AODL.Document.Content ;

namespace AODL.Document.Content.Charts
{


	/// <summary>
	/// Summary description for ChartDataPoint.
	/// </summary>
	public class ChartDataPoint : IContent
	{
		
		/// <summary>
		/// the chart which include the chartdatapoint
		/// </summary>
		private Chart _chart;

		public Chart Chart
		{
			get { return _chart; }
			set { _chart = value; }
		}

		/// <summary>
		/// gets and sets the style of the datapoint
		/// </summary>

		public DataPointStyle DataPointStyle
		{
			get
			{
				return (DataPointStyle)Style;
			}
			set
			{
				StyleName	= ((DataPointStyle)value).StyleName;
				Style = value;
			}
		}

		/// <summary>
		/// gets and sets the cell repeat count
		/// </summary>

		public string Repeated
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:repeated",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:repeated",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("repeated", value, "chart");
				_node.SelectSingleNode("@chart:repeated",
					Document.NamespaceManager).InnerText = value;
			}
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

		public ChartDataPoint(IDocument document, XmlNode node)
		{
			Document			= document;
			Node				= node;
			//this.InitStandards();
		}

		/// <summary>
		/// Initializes a new instance of the ChartPlotArea class.
		/// This will create an empty cell that use the default cell style
		/// </summary>
		/// <param name="table">The table.</param>
		public ChartDataPoint(Chart chart)
		{
			Chart				= chart;
			Document			= chart.Document;
			NewXmlNode(null);
			DataPointStyle = new DataPointStyle (chart.Document);
			Chart .Styles .Add (DataPointStyle ); 
			//this.InitStandards();
		}

		public ChartDataPoint(Chart chart, string styleName)
		{
			Chart				= chart;
			Document			= chart.Document;
			NewXmlNode(styleName);

			//this.InitStandards();

			if (styleName != null)
			{
				StyleName		= styleName;
				DataPointStyle		= new DataPointStyle(Document, styleName);
				Chart.Styles.Add(DataPointStyle);
			}
		}

		public void NewXmlNode(string styleName)
		{
			Node		= Document.CreateNode("data-point", "chart");

			XmlAttribute xa = Document.CreateAttribute("style-name", "chart");
			xa.Value		= styleName;
			Node.Attributes.Append(xa);

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
