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

namespace AODL.Document .Styles .Properties
{
	/// <summary>
	/// Summary description for ChartGraphicProperties.
	/// </summary>
	public class ChartGraphicProperties : IProperty
	{
		/// <summary>
		/// gets and sets the draw stroke
		/// </summary>
		public string DrawStroke
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:stroke",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:stroke",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("stroke", value, "draw");
				_node.SelectSingleNode("@draw:stroke",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets stroke dash
		/// </summary>

		public string StrokeDash
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:stroke-dash",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:stroke-dash",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("stroke-dash", value, "draw");
				_node.SelectSingleNode("@draw:stroke-dash",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets stroke dash name
		/// </summary>

		public string StrokeDashNames
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:stroke-dash-names",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:stroke-dash-names",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("stroke-dash-names", value, "draw");
				_node.SelectSingleNode("@draw:stroke-dash-names",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the stroke width
		/// </summary>

		public string StrokeWidth
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@svg:stroke-width",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@svg:stroke-width",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("stroke-width", value, "draw");
				_node.SelectSingleNode("@svg:stroke-width",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets stroke color
		/// </summary>

		public string StrokeColor
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@svg:stroke-color",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@svg:stroke-color",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("stroke-color", value, "draw");
				_node.SelectSingleNode("@svg:stroke-color",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets marker start
		/// </summary>

		public string MarkerStart
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:marker-start",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:marker-start",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("marker-start", value, "draw");
				_node.SelectSingleNode("@draw:marker-start",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

        /// <summary>
        /// gets and sets marker end
        /// </summary>
		
		public string MarkerEnd
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:marker-end",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:marker-end",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("marker-end", value, "draw");
				_node.SelectSingleNode("@draw:marker-end",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets marker start width
		/// </summary>

		public string MarkerStartWidth
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:marker-start-width",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:marker-start-width",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("marker-start-width", value, "draw");
				_node.SelectSingleNode("@draw:marker-start-width",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}
        
		/// <summary>
		/// gets and set marker end width
		/// </summary>
		public string MarkerEndWidth
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:marker-end-width",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:marker-end-width",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("marker-end-width", value, "draw");
				_node.SelectSingleNode("@draw:marker-end-width",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets marker end center
		/// </summary>

		public string MarkerEndCenter
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:marker-end-center",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:marker-end-center",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("marker-end-center", value, "draw");
				_node.SelectSingleNode("@draw:marker-end-center",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets marker start center
		/// </summary>
 
		public string MarkerStartCenter
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:marker-start-center",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:marker-start-center",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("marker-start-center", value, "draw");
				_node.SelectSingleNode("@draw:marker-start-center",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets stroke opacity
		/// </summary>
		public string StrokeOpacity
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@svg:stroke-opacity",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@svg:stroke-opacity",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("stroke-opacity", value, "draw");
				_node.SelectSingleNode("@svg:stroke-opacity",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}


		/// <summary>
		/// gets and sets stroke line join
		/// </summary>
		public string StrokeLineJoin
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:stroke-linejoin",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:stroke-linejoin",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("stroke-linejoin", value, "draw");
				_node.SelectSingleNode("@draw:stroke-linejoin",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets fill color
		/// </summary>
		public string FillColor
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:fill-color",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:fill-color",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("fill-color", value, "draw");
				_node.SelectSingleNode("@draw:fill-color",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}
  
		/// <summary>
		/// gets and sets secondary fill color
		/// </summary>
		public string SecondaryFillColor
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:secondary-fill-color",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:secondary-fill-color",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("secondary-fill-color", value, "draw");
				_node.SelectSingleNode("@draw:secondary-fill-color",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}
        

		/// <summary>
		/// gets and sets fill gradient name
		/// </summary>
		public string FillGradientName
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:fill-gradient-name",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:fill-gradient-name",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("fill-gradient-name", value, "draw");
				_node.SelectSingleNode("@draw:fill-gradient-name",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}
        
		/// <summary>
		/// gets and sets gradient step count
		/// </summary>
		public string GradientStepCount
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:gradient-step-count",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:gradient-step-count",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("gradient-step-count", value, "draw");
				_node.SelectSingleNode("@draw:gradient-step-count",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets fill hatch name
		/// </summary>

		public string FillHatchName
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:fill-hatch-name",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:fill-hatch-name",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("fill-hatch-name", value, "draw");
				_node.SelectSingleNode("@draw:fill-hatch-name",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets fill hatch solid
		/// </summary>

		public string FillHatchSolid
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:fill-hatch-solid",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:fill-hatch-solid",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("fill-hatch-solid", value, "draw");
				_node.SelectSingleNode("@draw:fill-hatch-solid",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}
        
		/// <summary>
		/// gets and sets fill image name
		/// </summary>
		public string FillImageName
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:fill-image-name",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:fill-image-name",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("fill-image-name", value, "draw");
				_node.SelectSingleNode("@draw:fill-image-name",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets style repeat
		/// </summary>

		public string StyleRepeat
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:repeat",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@style:repeat",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("repeat", value, "style");
				_node.SelectSingleNode("@style:repeat",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets fill image width
		/// </summary>

		public string FillImageWidth
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:fill-image-width",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:fill-image-width",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("fill-image-width", value, "draw");
				_node.SelectSingleNode("@draw:fill-image-width",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets fill image height
		/// </summary>

		public string FillImageHeight
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:fill-image-height",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:fill-image-height",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("fill-image-height", value, "draw");
				_node.SelectSingleNode("@draw:fill-image-height",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets fill image ref point
		/// </summary>

		public string FillImageRefPoint
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:fill-image-ref-point",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:fill-image-ref-point",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("fill-image-ref-point", value, "draw");
				_node.SelectSingleNode("@draw:fill-image-ref-point",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets fill image ref pointx
		/// </summary>

		public string FillImageRefPointX
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:fill-image-ref-point-x",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:fill-image-ref-point-x",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("fill-image-ref-point-x", value, "draw");
				_node.SelectSingleNode("@draw:fill-image-ref-point-x",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets fill image ref pointy
		/// </summary>

		public string FillImageRefPointY
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:fill-image-ref-point-y",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:fill-image-ref-point-y",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("fill-image-ref-point-y", value, "draw");
				_node.SelectSingleNode("@draw:fill-image-ref-point-y",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets tile repeat offset
		/// </summary>

		public string TileRepeatOffset
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:tile-repeat-offset",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:tile-repeat-offset",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("tile-repeat-offset", value, "draw");
				_node.SelectSingleNode("@draw:tile-repeat-offset",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}


		/// <summary>
		/// gets and sets opacity
		/// </summary>

		public string Opacity
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:opacity",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:opacity",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("opacity", value, "draw");
				_node.SelectSingleNode("@draw:opacity",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets opacity name
		/// </summary>

		public string OpacityName
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:opacity-name",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:opacity-name",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("opacity-name", value, "draw");
				_node.SelectSingleNode("@draw:opacity-name",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets fill rule
		/// </summary>

		public string FillRule
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@svg:fill-rule",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@svg:fill-rule",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("fill-rule", value, "svg");
				_node.SelectSingleNode("@svg:fill-rule",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets symbol color
		/// </summary>

		public string SymbolColor
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@draw:symbol-color",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:symbol-color",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("symbol-color", value, "svg");
				_node.SelectSingleNode("@draw:symbol-color",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets horizontal segment
		/// </summary>

		public string HorizontalSegment
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:horizontal-segments",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:horizontal-segments",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("horizontal-segments", value, "dr3d");
				_node.SelectSingleNode("@dr3d:horizontal-segments",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets vertical segment
		/// </summary>

		public string VerticalSegment
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:vertical-segments",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:vertical-segments",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("vertical-segments", value, "dr3d");
				_node.SelectSingleNode("@dr3d:vertical-segments",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets edge rounding
		/// </summary>

		public string EdgeRounding
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:edge-rounding",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:edge-rounding",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("edge-rounding", value, "dr3d");
				_node.SelectSingleNode("@dr3d:edge-rounding",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets edge rounding mode
		/// </summary>

		public string EdgeRoundingMode
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:edge-rounding-mode",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:edge-rounding-mode",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("edge-rounding-mode", value, "dr3d");
				_node.SelectSingleNode("@dr3d:edge-rounding-mode",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets back scale
		/// </summary>

		public string BackScale
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:back-scale",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:back-scale",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("back-scale", value, "dr3d");
				_node.SelectSingleNode("@dr3d:back-scale",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets depth
		/// </summary>

		public string Depth
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:depth",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:depth",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("depth", value, "dr3d");
				_node.SelectSingleNode("@dr3d:depth",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets back face culling
		/// </summary>

		public string BackFaceCulling
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:backface-culling",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:backface-culling",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("backface-culling", value, "dr3d");
				_node.SelectSingleNode("@dr3d:backface-culling",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets end angle
		/// </summary>

		public string EndAngle
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:end-angle",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:end-angle",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("end-angle", value, "dr3d");
				_node.SelectSingleNode("@dr3d:end-angle",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets close front
		/// </summary>

		public string CloseFront
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:close-front",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:close-front",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("close-front", value, "dr3d");
				_node.SelectSingleNode("@dr3d:close-front",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets close back
		/// </summary>

		public string CloseBack
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:close-back",
					Style.Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:close-back",
					Style.Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("close-back", value, "dr3d");
				_node.SelectSingleNode("@dr3d:close-back",
					Style.Document.NamespaceManager).InnerText = value;
			}
		}

        /// <summary>
        /// the constructor of chart graphic property
        /// </summary>
        /// <param name="style"></param>
		public ChartGraphicProperties(IStyle style)
		{
			Style			= style;
			NewXmlNode();
		}

		/// <summary>
		/// Create the XmlNode which represent the propertie element.
		/// </summary>
		private void NewXmlNode()
		{
			Node		= Style.Document.CreateNode("style", "graphic-properties");
		}

		/// <summary>
		/// Create a XmlAttribute for propertie XmlNode.
		/// </summary>
		/// <param name="name">The attribute name.</param>
		/// <param name="text">The attribute value.</param>
		/// <param name="prefix">The namespace prefix.</param>
		private void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = Style.Document.CreateAttribute(name, prefix);
			xa.Value		= text;
			Node.Attributes.Append(xa);
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
