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

namespace AODL.Document.Styles.Properties
{
	/// <summary>
	/// Summary description for AxesProperties.
	/// </summary>
	public class AxesProperties : ChartProperties
	{
		/// <summary>
		/// gets and sets the link data format
		/// </summary>

		public string LinkDataFormat
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:link-data-style-to-source",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:link-data-style-to-source",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("link-data-style-to-source", value, "chart");
				Node .SelectSingleNode("@chart:link-data-style-to-source",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the visibility
		/// </summary>

		public string Visibility
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:axis-visible",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:axis-visible",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("axis-visible", value, "chart");
				Node.SelectSingleNode("@chart:axis-visible",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}
        
		/// <summary>
		/// gets and sets the logarithmic
		/// </summary>
		public string Logarithmic
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:axis-logarithmic",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:axis-logarithmic",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("axis-logarithmic", value, "chart");
				Node.SelectSingleNode("@chart:axis-logarithmic",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets tick mark major inner
		/// </summary>

		public string TickMarkMajorInner
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:tick-marks-major-inner",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:tick-marks-major-inner",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("tick-marks-major-inner", value, "chart");
				Node.SelectSingleNode("@chart:tick-marks-major-inner",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets tick mark major outer
		/// </summary>

		public string TickMarkMajorOuter
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:tick-marks-minor-outer",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:tick-marks-major-inner",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("tick-marks-minor-outer", value, "chart");
				Node.SelectSingleNode("@chart:tick-marks-minor-outer",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}
        
		/// <summary>
		/// gets and set the maximum
		/// </summary>
		public string Maximum
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:maximum",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:maximum",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("maximum", value, "chart");
				Node.SelectSingleNode("@chart:maximum",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets minimum
		/// </summary>

		public string Minimum
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:minimum",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:minimum",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("minimum", value, "chart");
				Node.SelectSingleNode("@chart:minimum",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and set origin
		/// </summary>

		public string Origin
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:origin",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:origin",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("origin", value, "chart");
				Node.SelectSingleNode("@chart:origin",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the interval major
		/// </summary>

		public string IntervalMajor
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:interval-major",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:interval-major",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("interval-major", value, "chart");
				Node.SelectSingleNode("@chart:interval-major",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets interval minor
		/// </summary>

		public string IntervalMinor
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:interval-minor",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:interval-minor",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("interval-minor", value, "chart");
				Node.SelectSingleNode("@chart:interval-minor",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets tick mark minor inner
		/// </summary>

		public string TickMarkMinorInner
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:tick-marks-minor-inner",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:tick-marks-minor-inner",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("tick-marks-minor-inner", value, "chart");
				Node.SelectSingleNode("@chart:tick-marks-minor-inner",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets tick mark minor outer
		/// </summary>

		public string TickMarkMinorOuter
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:tick-marks-minor-outer",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:tick-marks-minor-outer",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("tick-marks-minor-outer", value, "chart");
				Node.SelectSingleNode("@chart:tick-marks-minor-outer",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the display label
		/// </summary>

		public string DisplayLabel
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:display-label",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:display-label",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("display-label", value, "chart");
				Node.SelectSingleNode("@chart:display-label",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets text overlap
		/// </summary>

		public string TextOverLap
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:text-overlap",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:text-overlap",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("text-overlap", value, "chart");
				Node.SelectSingleNode("@chart:text-overlap",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets line break
		/// </summary>

		public string LineBreak
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@text:line-break",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@text:line-break",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("line-break", value, "text");
				Node.SelectSingleNode("@text:line-break",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the label arragement
		/// </summary>

		public string LabelArrangement
		{
			get 
			{ 
				XmlNode xn = Node.SelectSingleNode("@chart:label-arrangement",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = Node.SelectSingleNode("@chart:label-arrangement",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("label-arrangement", value, "chart");
				Node.SelectSingleNode("@chart:label-arrangement",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}



		private new void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = Style.Document.CreateAttribute(name, prefix);
			xa.Value		= text;
			Node.Attributes.Append(xa);
		}


		
		/// <summary>
		/// The Constructor, create new instance of ChartProperties
		/// </summary>
		/// <param name="style">The ColumnStyle</param>
		public AxesProperties(IStyle style): base(style)
		{
			//this.Style			= style;
			//this.NewXmlNode();
		}

		
		/// <summary>
		/// Create the XmlNode which represent the propertie element.
		/// </summary>
		private new void NewXmlNode()
		{
			Node		= Style.Document.CreateNode("chart-properties", "style");
		}

		#region IProperty Member
	
		#endregion

 




	}
}

