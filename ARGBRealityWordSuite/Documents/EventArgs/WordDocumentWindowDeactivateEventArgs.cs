using Word = Microsoft.Office.Interop.Word;

namespace ARGBRealityWordSuite.Documents.EventArgs
{
	public class WordDocumentWindowDeactivateEventArgs(Word.Document Doc, Word.Window Wn) : WordDocumentWindowActivateEventArgs(Doc, Wn)
	{
	}
}
