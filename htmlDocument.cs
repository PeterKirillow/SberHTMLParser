using System;
using System.Text;

namespace SberHTMLParser
{
	class htmlDocument
	{
		public string html;
		public htmlDocument(string fileName)
		{
			this.html = System.IO.File.ReadAllText(fileName);
		}
	}
}
