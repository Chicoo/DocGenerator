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
	/// Summary description for ChartPlotArea.
	/// </summary>
	public class ChartPlotArea : IContent , IContentContainer
	{
		/// <summary>
		/// the chart which contains the plotarea
		/// </summary>
		private Chart _chart;

		public Chart Chart
		{
			get { return _chart; }
			set { _chart = value; }
		}

		/// <summary>
		/// Gets or sets the plotarea style.
		/// </summary>
		/// <value>The plotarea style.</value>
		public PlotAreaStyle PlotAreaStyle
		{
			get
			{
				return (PlotAreaStyle)Style;
			}
			set
			{
				StyleName	= value.StyleName;
				Style = value;
			}
		}
		#region plotarea properties
		/// <summary>
		/// Gets or sets the horizontal position where
		/// the plotarea should be
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
		/// the plotarea should be
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
		/// the width of the plotarea
		/// </summary>

		public string Width
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@svg:width",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@svg:width",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("width", value, "svg");
				_node.SelectSingleNode("@svg:width",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// the height of the plotarea
		/// </summary>

		public string Height
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@svg:height",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@svg:height",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("height", value, "svg");
				_node.SelectSingleNode("@svg:height",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// sets and gets the data source has labels
		/// </summary>

		public string DataSourceHasLabels
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:data-source-has-labels",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:data-source-has-labels",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("data-source-has-labels", value, "chart");
				_node.SelectSingleNode("@chart:data-source-has-labels",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the table number list
		/// </summary>

		
		public string TableNumberList
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:table-number-list",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:table-number-list",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("table-number-list", value, "chart");
				_node.SelectSingleNode("@chart:table-number-list",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the dr3dvrp
		/// </summary>

		public string Dr3dVrp
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:vrp",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:vrp",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("vrp", value, "dr3d");
				_node.SelectSingleNode("@dr3d:vrp",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the dr3dvpn
		/// </summary>

		public string Dr3dVpn
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:vpn",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:vpn",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("vpn", value, "dr3d");
				_node.SelectSingleNode("@dr3d:vpn",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the dr3dvup
		/// </summary>

		public string Dr3dVup
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:vup",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:vup",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("vup", value, "dr3d");
				_node.SelectSingleNode("@dr3d:vup",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the projection
		/// </summary>

		public string Projection
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:projection",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:projection",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("projection", value, "dr3d");
				_node.SelectSingleNode("@dr3d:projection",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the distance
		/// </summary>

		public string  Distance
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:distance",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:distance",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("distance", value, "dr3d");
				_node.SelectSingleNode("@dr3d:distance",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the focallength
		/// </summary>

		public string  FocalLength
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:focal-length",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:focal-length",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("focal-length", value, "dr3d");
				_node.SelectSingleNode("@dr3d:focal-length",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the shadowslant
		/// </summary>

		public string  ShadowSlant
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:shadow-slant",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:shadow-slant",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("shadow-slant", value, "dr3d");
				_node.SelectSingleNode("@dr3d:shadow-slant",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the shademode
		/// </summary>

		
		public string  ShadeMode
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:shade-mode",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:shade-mode",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("shade-mode", value, "dr3d");
				_node.SelectSingleNode("@dr3d:shade-mode",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the ambientcolor
		/// </summary>

		public string  AmbientColor
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:ambient-color",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:ambient-color",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("ambient-color", value, "dr3d");
				_node.SelectSingleNode("@dr3d:ambient-color",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the lightingmode
		/// </summary>

		public string  LightingMode
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:lighting-mode",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:lighting-mode",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("lighting-mode", value, "dr3d");
				_node.SelectSingleNode("@dr3d:lighting-mode",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets table cell range
		/// </summary>

		public string  TableCellRange
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@table:cell-range-address",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@table:cell-range-address",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("cell-range-address", value, "table");
				_node.SelectSingleNode("@table:cell-range-address",
					Document.NamespaceManager).InnerText = value;
			}
		}

		#endregion 

		private AxisCollection _axiscollection;

		public AxisCollection AxisCollection
		{
			get {  return _axiscollection;}

			set { _axiscollection=value; }
		}

		private Dr3dLightCollection _dr3dlightcollection;

		public Dr3dLightCollection Dr3dLightCollection
		{ 
			get {  return _dr3dlightcollection;}

			set { _dr3dlightcollection=value; }
		}

		private SeriesCollection _seriescollection;

		public SeriesCollection SeriesCollection
		{ 
			get {  return _seriescollection;}

			set { _seriescollection=value; }
		}



		/// <summary>
		/// Initializes a new instance of the ChartPlotArea class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="node">The node.</param>

		public ChartPlotArea(IDocument document, XmlNode node)
		{
			Document			= document;
			Node				= node;
			InitStandards(); 
		}

		public ChartPlotArea(Chart chart)
		{
			Chart				= chart;
			Document			= chart.Document;
			NewXmlNode(null);
			InitStandards();
			PlotAreaStyle = new PlotAreaStyle (chart.Document);
			chart.Styles .Add (PlotAreaStyle );
			InitPlotArea ();
			chart.Content .Add (this);
		}

		

		/// <summary>
		/// Initializes a new instance of the ChartPlotArea class.
		/// </summary>
		/// <param name="table">The table.</param>
		/// <param name="styleName">The style name.</param>
		public ChartPlotArea(Chart chart, string styleName)
		{
			Chart				= chart;
			Document			= chart.Document;
			NewXmlNode(styleName);

			InitStandards();

			if (styleName != null)
			{
				StyleName		= styleName;			
				PlotAreaStyle = new PlotAreaStyle (chart.Document,styleName);
				chart.Styles .Add (PlotAreaStyle );
				
			}
			InitPlotArea ();

			chart.Content .Add (this);
		}




		private void InitStandards()
		{
			Content			     = new ContentCollection();
			AxisCollection          = new AxisCollection ();
			Dr3dLightCollection     = new Dr3dLightCollection ();
			SeriesCollection        = new SeriesCollection ();
            
			//	AxisCollection.Inserted      += new ZipStream.CollectionWithEvents.CollectionChange (AxisCollection_Inserted);
			//    AxisCollection.Removed       += new ZipStream.CollectionWithEvents.CollectionChange (AxisCollection_Removed);
			//	Dr3dLightCollection.Inserted += new ZipStream.CollectionWithEvents.CollectionChange (Dr3dLightCollection_Inserted);
			//    Dr3dLightCollection.Removed  += new ZipStream.CollectionWithEvents.CollectionChange (Dr3dLightCollection_Removed);
			//	SeriesCollection.Inserted    += new ZipStream.CollectionWithEvents.CollectionChange (SeriesCollection_Inserted);
			//	SeriesCollection.Removed     += new ZipStream.CollectionWithEvents.CollectionChange (SeriesCollection_Removed);	
			Content.Inserted	     += Content_Inserted;
			Content.Removed	     += Content_Removed;
		}


		private void NewXmlNode(string styleName)
		{
			Node		= Document.CreateNode("plot-area", "chart");

			XmlAttribute xa = Document.CreateAttribute("style-name", "chart");
			xa.Value		= styleName;
			Node.Attributes.Append(xa);
		}

	
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
		/// Builds the node.
		/// </summary>
		/// <returns>XmlNode</returns>

		public XmlNode  BuildNode()
		{
			foreach(ChartAxis axes in AxisCollection)
			{
				if (Document .IsLoadedFile &&axes.IsModified&&!Chart .IsNewed)
					axes.Node =Chart .ChartDoc .ImportNode (axes.Node ,true);
				Node.AppendChild (axes.Node);
			}
			
			foreach(Dr3dLight light in Dr3dLightCollection)
			{
				
				Node .AppendChild (light.Node );
			}

			foreach(ChartSeries series in SeriesCollection)
			{
				//node=this.Chart .ChartDoc.ImportNode (series.BuildNode (),true );
				Node .AppendChild (series.BuildNode ());
			}

			return Node ;
		}

		/// <summary>
		/// if  fistLine has labels or not
		/// </summary>
		/// <returns></returns>

		public bool FirstLineHasLabels()
		{
			if (DataSourceHasLabels=="row"||DataSourceHasLabels=="both")
				return true;
			else
				return false;

		}


		/// <summary>
		/// if  fistColumnl  has labels          
		/// </summary>
		/// <returns></returns>
		public bool FirstColumnHasLabels()
		{
			
			if (DataSourceHasLabels=="column"||DataSourceHasLabels=="both")
				return true;
			else
				return false;
		}

		/// <summary>
		/// Init the plotarea
		/// </summary>

		public void InitPlotArea()
		{
			string SeriesSource = PlotAreaStyle .PlotAreaProperties.SeriesSource ;		
			
			if (SeriesSource==null)			
				PlotAreaStyle .PlotAreaProperties .SeriesSource ="columns";
			//SeriesSource= this.PlotAreaStyle .PlotAreaProperties .SeriesSource;
		}

	
	}
}

