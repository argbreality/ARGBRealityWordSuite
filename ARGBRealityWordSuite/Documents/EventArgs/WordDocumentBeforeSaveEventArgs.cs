using Word = Microsoft.Office.Interop.Word;

namespace ARGBRealityWordSuite.Documents.EventArgs
{
	public class WordDocumentBeforeSaveEventArgs : WordDocumentEventArgs
	{
		public WordDocumentBeforeSaveEventArgs(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel) : base(Doc)
		{
			this.SaveAsUI = SaveAsUI;
			this.Cancel = Cancel;
		}

		public bool SaveAsUI { get; }
		
		public bool Cancel { get; }
	}
}
