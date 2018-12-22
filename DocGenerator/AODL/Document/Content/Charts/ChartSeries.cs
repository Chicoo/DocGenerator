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

namespace AODL.Document.Content.Charts
{
	/// <summary>
	/// Summary description for ChartSeries.
	/// </summary>
	public class ChartSeries :IContent
	{
		/// <summary>
		/// the chart which contains the series
		/// </summary>
		private Chart _chart;

		public Chart Chart
		{
			get { return _chart; }
			set { _chart = value; }
		}

		/// <summary>
		/// the plotarea which contains the series
		/// </summary>

		private ChartPlotArea _plotarea;

		public ChartPlotArea PlotArea
		{
			get { return _plotarea; }
			set { _plotarea = value; }
		}

		/// <summary>
		/// the style of the series
		/// </summary>

		public SeriesStyle SeriesStyle
		{
			get
			{
				return (SeriesStyle)Style;
			}
			set
			{
				StyleName	= ((SeriesStyle)value).StyleName;
				Style = value;
			}
		}

		/// <summary>
		/// the collection of the data point which the series contains
		/// </summary>

		private DataPointCollection _datapointcollection;

		public DataPointCollection  DataPointCollection
		{
			get
			{
				return _datapointcollection;
			}

			set
			{
				_datapointcollection =value;
			}
		}

		/// <summary>
		/// gets and sets the value cell range address
		/// </summary>

		public string ValuesCellRangeAddress
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:values-cell-range-address",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:values-cell-range-address",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("values-cell-range-address", value, "chart");
				_node.SelectSingleNode("@chart:values-cell-range-address",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the label cell address
		/// </summary>

		public string LabelCellAddress
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:label-cell-address",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:label-cell-address",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("label-cell-address", value, "chart");
				_node.SelectSingleNode("@chart:label-cell-address",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the class 
		/// </summary>

		public string Class
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:class",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:class",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("class", value, "chart");
				_node.SelectSingleNode("@chart:class",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the attached axis
		/// </summary>

		public string AttachAxis
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:attached-axis",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:attached-axis",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("attached-axis", value, "chart");
				_node.SelectSingleNode("@chart:attached-axis",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// the constructor of the chartseries
		/// </summary>
		/// <param name="document"></param>
		/// <param name="node"></param>

		public ChartSeries(IDocument document, XmlNode node)
		{
			Document			= document;
			Node				= node;
			DataPointCollection =new DataPointCollection ();
			InitStandards();
		}

		/// <summary>
		/// Initializes a new instance of the ChartPlotArea class.
		/// This will create an empty cell that use the default cell style
		/// </summary>
		/// <param name="table">The table.</param>
		public ChartSeries(Chart chart)
		{
			Chart				= chart;
			Document			= chart.Document;
			DataPointCollection =new DataPointCollection ();	
			NewXmlNode(null);
			SeriesStyle =new SeriesStyle (Document);
			Chart .Styles .Add (SeriesStyle);
		
			InitStandards();
		}

		/// <summary>
		/// Initializes a new instance of the ChartPlotArea class.
		/// </summary>
		/// <param name="table">The table.</param>
		/// <param name="styleName">The style name.</param>
		public ChartSeries(Chart chart, string styleName)
		{
			Chart				= chart;
			Document			= chart.Document;
			NewXmlNode(styleName);
			InitStandards();
			DataPointCollection  = new DataPointCollection ();

			if (styleName != null)
			{
				StyleName		= styleName;
				SeriesStyle		= new SeriesStyle(Document, styleName);
				Chart.Styles.Add(SeriesStyle);
			}
		}

		private void NewXmlNode(string styleName)
		{
			Node		= Document.CreateNode("series", "chart");

			XmlAttribute xa = Document.CreateAttribute("style-name", "chart");
			xa.Value		= styleName;
			Node.Attributes.Append(xa);
		}

		private void InitStandards()
		{
			Content			= new ContentCollection();
			Content.Inserted	+= Content_Inserted;
			Content.Removed	+= Content_Removed;
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

		/// <summary>
		/// builds the content
		/// </summary>
		/// <returns></returns>

		public XmlNode BuildNode()
		{
			foreach(ChartDataPoint DataPoint in DataPointCollection)
			{
				//XmlNode node = this.Chart .ChartDoc .ImportNode (DataPoint.Node ,true);
				Node .AppendChild (DataPoint.Node );
			}
			return Node ;
		}
	
	}
}

