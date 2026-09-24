using ARGBRealityWordSuite.Documents.EventArgs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using Word = Microsoft.Office.Interop.Word;

namespace ARGBRealityWordSuite.Documents
{
	public static class DocumentCollection
	{
		private static readonly Dictionary<int, Document> documents = [];

		private static Document GetDocument(Word.Document Doc)
		{
			int hwnd = Doc.Windows[1].Hwnd;

			if (documents.ContainsKey(hwnd))
			{
				return documents[hwnd];
			}
			else
			{
				return null;
			}
		}

		public static void Initialize(Word.Application wordApp)
		{
			((Word.ApplicationEvents4_Event)wordApp).NewDocument += new Word.ApplicationEvents4_NewDocumentEventHandler(DoNewDocument);
			wordApp.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(DoDocumentOpen);
			wordApp.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(DoDocumentBeforeClose);
			
			wordApp.WindowActivate += new Word.ApplicationEvents4_WindowActivateEventHandler(DoDocumentWindowActivate);
			wordApp.WindowDeactivate += new Word.ApplicationEvents4_WindowDeactivateEventHandler(DoDocumentWindowDeactivate);
			wordApp.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(DoDocumentSelectionChange);

			wordApp.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(DoDocumentBeforeSave);
			wordApp.DocumentBeforePrint += new Word.ApplicationEvents4_DocumentBeforePrintEventHandler(DoDocumentBeforePrint);
		}
		
		/// <summary>
		/// Активация окна документа.
		/// </summary>
		[Browsable(true), EditorBrowsable(EditorBrowsableState.Always)]
		public static event EventHandler<WordDocumentWindowActivateEventArgs> DocumentWindowActivate;

		private static void DoDocumentWindowActivate(Word.Document Doc, Word.Window Wn) => 
			DocumentWindowActivate?.Invoke(GetDocument(Doc), new WordDocumentWindowActivateEventArgs(Doc, Wn));
		
		/// <summary>
		/// Деактивация окна документа.
		/// </summary>
		[Browsable(true), EditorBrowsable(EditorBrowsableState.Always)]
		public static event EventHandler<WordDocumentWindowDeactivateEventArgs> DocumentWindowDeactivate;

		private static void DoDocumentWindowDeactivate(Word.Document Doc, Word.Window Wn) => 
			DocumentWindowDeactivate?.Invoke(GetDocument(Doc), new WordDocumentWindowDeactivateEventArgs(Doc, Wn));
			
		/// <summary>
		/// Событие, связанное с созданием нового документа.
		/// </summary>
		[Browsable(true), EditorBrowsable(EditorBrowsableState.Always)]
		public static event EventHandler<WordDocumentEventArgs> NewDocument;

		private static void DoNewDocument(Word.Document Doc) =>
			NewDocument?.Invoke(GetDocument(Doc), new WordDocumentEventArgs(Doc));

		/// <summary>
		/// Событие, связанное с открытием документа.
		/// </summary>
		[Browsable(true), EditorBrowsable(EditorBrowsableState.Always)]
		public static event EventHandler<WordDocumentEventArgs> DocumentOpen;

		private static void DoDocumentOpen(Word.Document Doc) =>
			DocumentOpen?.Invoke(GetDocument(Doc), new WordDocumentEventArgs(Doc));

		/// <summary>
		/// Событие, связанное с закрытием документа.
		/// </summary>
		[Browsable(true), EditorBrowsable(EditorBrowsableState.Always)]
		public static event EventHandler<WordDocumentBeforeCloseEventArgs> DocumentBeforeClose;

		private static void DoDocumentBeforeClose(Word.Document Doc, ref bool Cancel) =>
			DocumentOpen?.Invoke(GetDocument(Doc), new WordDocumentBeforeCloseEventArgs(Doc, ref Cancel));

		/// <summary>
		/// Событие, связанное с выделением части документа.
		/// </summary>
		[Browsable(true), EditorBrowsable(EditorBrowsableState.Always)]
		public static event EventHandler<WordSelectionEventArgs> DocumentSelectionChange;

		private static void DoDocumentSelectionChange(Word.Selection Sel) => 
			DocumentSelectionChange?.Invoke(GetDocument(Sel.Document), new WordSelectionEventArgs(Sel));
			
		/// <summary>
		/// Событие, связанное с печатью документа.
		/// </summary>
		[Browsable(true), EditorBrowsable(EditorBrowsableState.Always)]
		public static event EventHandler<WordDocumentBeforePrintEventArgs> DocumentBeforePrint;

		private static void DoDocumentBeforePrint(Word.Document Doc, ref bool Cancel) => 
			DocumentBeforePrint?.Invoke(GetDocument(Doc), new WordDocumentBeforePrintEventArgs(Doc, ref Cancel));
		
		/// <summary>
		/// Событие, связанное с сохранением документа.
		/// </summary>
		[Browsable(true), EditorBrowsable(EditorBrowsableState.Always)]
		public static event EventHandler<WordDocumentBeforeSaveEventArgs> DocumentBeforeSave;

		private static void DoDocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel) => 
			DocumentBeforeSave?.Invoke(GetDocument(Doc), new WordDocumentBeforeSaveEventArgs(Doc, ref SaveAsUI, ref Cancel));

		public static void Dispose()
		{
			
		}
	}
}
