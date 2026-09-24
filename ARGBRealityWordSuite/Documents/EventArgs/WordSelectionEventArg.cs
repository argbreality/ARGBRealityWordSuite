using Word = Microsoft.Office.Interop.Word;

namespace ARGBRealityWordSuite.Documents.EventArgs
{
	public class WordSelectionEventArgs(Word.Selection Sel) : System.EventArgs
	{
		public Word.Selection Selection { get; } = Sel;

	}
}
