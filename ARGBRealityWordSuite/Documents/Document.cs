using ARGBRealityWordSuite.Documents.EventArgs;
using System;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace ARGBRealityWordSuite.Documents
{
	public partial class Document : IDisposable
	{
		public static Document Create(string fileName, Word._Document Doc)
		{
			Document document = new(
				fileName: fileName,
				doc: Doc);
			return document;
		}


		public Word._Document Doc { get; }

		public int Hwnd { get; } 

		public string FileName { get; }

		private Document(string fileName, Word._Document doc)
		{
			FileName = fileName;
			Doc = doc;
			Hwnd = Doc.Windows[1].Hwnd;
		}

		public void Dispose()
		{
			Marshal.ReleaseComObject(Doc);
		}
	}
}
