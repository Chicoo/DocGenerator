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
	/// Summary description for Dr3dLight.
	/// </summary>
	public class Dr3dLight : IContent
	{
		/// <summary>
		/// the chart which contains the dr3dlight
		/// </summary>
		private Chart _chart;

		public Chart Chart
		{
			get { return _chart; }
			set { _chart = value; }
		}

		/// <summary>
		/// the plotarea which contains the dr3dlight
		/// </summary>

		private ChartPlotArea _plotarea;

		public ChartPlotArea PlotArea
		{
			get { return _plotarea; }
			set { _plotarea = value; }
		}

		/// <summary>
		/// gets and sets the diffusecolor
		/// </summary>

		public string DiffuseColor
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:diffuse-color",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:diffuse-color",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("diffuse-color", value, "dr3d");
				_node.SelectSingleNode("@dr3d:diffuse-color",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the direction
		/// </summary>

		public string Direction
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:direction",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:direction",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("direction", value, "dr3d");
				_node.SelectSingleNode("@dr3d:direction",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the enabled
		/// </summary>

		public string Enabled
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:enabled",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:enabled",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("enabled", value, "dr3d");
				_node.SelectSingleNode("@dr3d:enabled",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the specular
		/// </summary>

		public string Specular
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@dr3d:specular",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@dr3d:specular",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("specular", value, "dr3d");
				_node.SelectSingleNode("@dr3d:specular",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// gets and sets the name of the style
		/// </summary>

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

		public XmlNode Node
		{
			get { return  _node;}
			set { _node=value; }
		}

		private IDocument _document;

		public IDocument  Document
		{
			get { return  _document;}
			set { _document=value; }
		}


		private void NewXmlNode()
		{
			Node =Document .CreateNode ("light","dr3d");

		}

		/// <summary>
		/// the constructor of the dr3dlight
		/// </summary>
		/// <param name="chart"></param>

		public Dr3dLight(Chart chart)
		{
			Chart =chart;
			NewXmlNode ();
		}

		public Dr3dLight(IDocument document,XmlNode node)
		{
			Document =document;
			Node =node;
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
