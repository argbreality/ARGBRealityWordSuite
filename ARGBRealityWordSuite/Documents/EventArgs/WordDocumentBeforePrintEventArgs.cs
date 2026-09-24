using Word = Microsoft.Office.Interop.Word;

namespace ARGBRealityWordSuite.Documents.EventArgs
{
	public class WordDocumentBeforePrintEventArgs : WordDocumentEventArgs
	{
		public WordDocumentBeforePrintEventArgs(Word.Document Doc, ref bool Cancel) : base(Doc)
		{
			this.Cancel = Cancel;
		}

		public bool Cancel { get; }
	}
}
