using Word = Microsoft.Office.Interop.Word;

namespace ARGBRealityWordSuite.Documents.EventArgs
{
	public class WordDocumentEventArgs(Word.Document Doc) : System.EventArgs
	{
		public Word.Document Document { get; } = Doc;

	}
}
