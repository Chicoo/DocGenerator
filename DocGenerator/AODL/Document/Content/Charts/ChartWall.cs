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
	/// Summary description for ChartWall.
	/// </summary>
	public class ChartWall : IContent
	{

		/// <summary>
		/// the chart which contains the wall
		/// </summary>
		private Chart _chart;

		public Chart Chart
		{
			get { return _chart; }
			set { _chart = value; }
		}
		
        
		/// <summary>
		/// the style of the chart wall
		/// </summary>
		public WallStyle WallStyle
		{
			get
			{
				return (WallStyle)Style;
			}
			set
			{
				StyleName	= ((WallStyle)value).StyleName;
				Style = value;
			}
		}
		/// <summary>
		/// Gets or sets the width of the wall
		/// </summary>

		public string SvgWidth
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
		/// the constructor of the chart wall
		/// </summary>
		/// <param name="document"></param>
		/// <param name="node"></param>

		public ChartWall(IDocument document,XmlNode node)
		{
			Document =document;
			Node =node;
		}


		public ChartWall(Chart chart)
		{
			Chart =chart;
			Document =chart.Document;
			NewXmlNode (null);
			WallStyle = new WallStyle (chart.Document);
			Chart .Styles .Add (WallStyle );
			Chart.Content.Add (this);
		}

		public ChartWall(Chart chart,string styleName)
		{
			Chart =chart;
			Document =chart.Document;
			NewXmlNode (styleName);

			if (styleName!=null)
			{
				StyleName =styleName;
				WallStyle = new WallStyle (chart.Document ,styleName);
				Chart .Styles .Add (WallStyle );
			}

			Chart .Content .Add (this);
		}



		private void NewXmlNode(string styleName)
		{
			Node		= Document.CreateNode("wall", "chart");

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

		private void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = Document.CreateAttribute(name, prefix);
			xa.Value		= text;
			Node.Attributes.Append(xa);
		}	
	}
}

