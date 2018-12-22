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

namespace AODL.Document .Styles .Properties
{
	/// <summary>
	/// Summary description for ChartProperties.
	/// </summary>
	public class ChartProperties : IProperty
	{

		/// <summary>
		/// gets and sets the direction
		/// </summary>
		public string Direction
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:direction",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@style:direction",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("direction", value, "style");
				_node.SelectSingleNode("@style:direction",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}
        
		/// <summary>
		/// gets and sets the three dimensional
		/// </summary>
		public string ThreeDimensional
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:three-dimensional",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:three-dimensional",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("three-dimensional", value, "chart");
				_node.SelectSingleNode("@chart:three-dimensional",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets candle stick
		/// </summary>

		public string CandleStick
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:japanese-candle-stick",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:japanese-candle-stick",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("japanese-candle-stick", value, "chart");
				_node.SelectSingleNode("@chart:japanese-candle-stick",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the data label number
		/// </summary>

		public string DataLabelNumber
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:data-label-number",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:data-label-number",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("data-label-number", value, "chart");
				_node.SelectSingleNode("@chart:data-label-number",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets data label text
		/// </summary>

		public string DataLabelText
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:data-label-text",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:data-label-text",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("data-label-text", value, "chart");
				_node.SelectSingleNode("@chart:data-label-text",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets data label symbol
		/// </summary>

		public string DataLabelSymbol
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:data-label-symbol",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:data-label-symbol",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("data-label-symbol", value, "chart");
				_node.SelectSingleNode("@chart:data-label-symbol",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets mean value
		/// </summary>

		public string MeanValue
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:mean-value",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:mean-value",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("mean-value", value, "chart");
				_node.SelectSingleNode("@chart:mean-value",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets error category
		/// </summary>

		public string ErrorCategory
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:error-category",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:error-category",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("error-category", value, "chart");
				_node.SelectSingleNode("@chart:error-category",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}
		

		/// <summary>
		/// gets and sets error percentage
		/// </summary>
		public string ErrorPercentage
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:error-percentage",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:error-percentage",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("error-percentage", value, "chart");
				_node.SelectSingleNode("@chart:error-percentage",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets error margin
		/// </summary>

		public string ErrorMargin
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:error-margin",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:error-margin",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("error-margin", value, "chart");
				_node.SelectSingleNode("@chart:error-margin",
					Style.Document.NamespaceManager).InnerText = value;
			 }
		}

		/// <summary>
		/// gets and sets error lower limit
		/// </summary>

		
		public string ErrorLowerLimit
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:error-lower-limit",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:error-lower-limit",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("error-lower-limit", value, "chart");
				_node.SelectSingleNode("@chart:error-lower-limit",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets error upper limit
		/// </summary>

		
		public string ErrorUpperLimit
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:error-upper-limit",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:error-upper-limit",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("error-upper-limit", value, "chart");
				_node.SelectSingleNode("@chart:error-upper-limit",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets error upper indicator
		/// </summary>

		
		public string ErrorUpperIndicator
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:error-upper-indicator",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:error-upper-indicator",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("error-upper-indicator", value, "chart");
				_node.SelectSingleNode("@chart:error-upper-indicator",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}


		/// <summary>
		/// gets and sets error lower indicator
		/// </summary>

		public string ErrorLowerIndicator
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:error-lower-indicator",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:error-lower-indicator",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("error-lower-indicator", value, "chart");
				_node.SelectSingleNode("@chart:error-lower-indicator",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets vertical
		/// </summary>

		public string Vertical
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:vertical",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:vertical",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("vertical", value, "chart");
				_node.SelectSingleNode("@chart:vertical",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets connect bars
		/// </summary>

		public string ConnectBars
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:connect-bars",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:connect-bars",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("connect-bars", value, "chart");
				_node.SelectSingleNode("@chart:connect-bars",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets gap width
		/// </summary>

		public string GapWidth
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:gap-width",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:gap-width",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("gap-width", value, "chart");
				_node.SelectSingleNode("@chart:gap-width",
					Style.Document.NamespaceManager).InnerText = value;
			}


		}

		/// <summary>
		/// gets and sets over lap
		/// </summary>

		public string OverLap
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:overlap",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:overlap",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("overlap", value, "chart");
				_node.SelectSingleNode("@chart:overlap",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		
		/// <summary>
		/// gets and sets deep
		/// </summary>
		public string Deep
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:deep",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:deep",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("deep", value, "chart");
				_node.SelectSingleNode("@chart:deep",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets symbol type
		/// </summary>

		public string SymbolType
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:symbol-type",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:symbol-type",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("symbol-type", value, "chart");
				_node.SelectSingleNode("@chart:symbol-type",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		
		/*	public string SymbolType
			{
				get 
				{ 
					XmlNode xn = this._node.SelectSingleNode("@chart:symbol-type",
						this.Style.Document.NamespaceManager);
					if (xn != null)
						return xn.InnerText;
					return null;
				}
				set
				{
					XmlNode xn = this._node.SelectSingleNode("@chart:symbol-type",
						this.Style.Document.NamespaceManager);
					if (xn == null)
						this.CreateAttribute("symbol-type", value, "chart");
					this._node.SelectSingleNode("@chart:symbol-type",
						this.Style.Document.NamespaceManager).InnerText = value;
				}
			}*/

		/// <summary>
		/// gets and sets lines
		/// </summary>
            
		public string Lines
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:lines",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:lines",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("lines", value, "chart");
				_node.SelectSingleNode("@chart:lines",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}


		/// <summary>
		/// gets and sets solid type
		/// </summary>

		public string SolidType
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:solid-type",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:solid-type",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("solid-type", value, "chart");
				_node.SelectSingleNode("@chart:solid-type",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets stacked
		/// </summary>

		public string Stacked
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:stacked",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:stacked",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("stacked", value, "chart");
				_node.SelectSingleNode("@chart:stacked",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}


		/// <summary>
		/// gets and sets percentage
		/// </summary>

		public string Percentage
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:percentage",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:percentaged",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("percentage", value, "chart");
				_node.SelectSingleNode("@chart:percentage",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets inter polation
		/// </summary>

		public string InterPolation
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@chart:interpolation",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@chart:interpolation",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("interpolation", value, "chart");
				_node.SelectSingleNode("@chart:interpolation",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}






		protected void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = Style.Document.CreateAttribute(name, prefix);
			xa.Value		= text;
			Node.Attributes.Append(xa);
		}


		
		/// <summary>
		/// The Constructor, create new instance of ChartProperties
		/// </summary>
		/// <param name="style">The ColumnStyle</param>
		public ChartProperties(IStyle style)
		{
			Style			= style;
			NewXmlNode();
		}
		
		/// <summary>
		/// Create the XmlNode which represent the propertie element.
		/// </summary>
		public void NewXmlNode()
		{
			Node		= Style.Document.CreateNode("chart-properties", "style");
		}


		#region IProperty Member
		private XmlNode _node;
		/// <summary>
		/// The XmlNode which represent the property element.
		/// </summary>
		/// <value>The node</value>
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

		private IStyle _style;
		/// <summary>
		/// The style object to which this property object belongs
		/// </summary>
		/// <value></value>
		public IStyle Style
		{
			get { return _style; }
			set { _style = value; }
		}
		#endregion
	}
}

