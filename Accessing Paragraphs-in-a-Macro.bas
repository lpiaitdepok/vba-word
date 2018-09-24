'For instance, let's say you wanted to access each paragraph of a document, in turn, and do some processing on the text in that paragraph. Since each paragraph is a distinct object in the document, this is relatively easy. All of the paragraph objects are accessible as part of the Paragraphs collection. The following code will do the trick:

iParCount = ActiveDocument.Paragraphs.Count
For J = 1 To iParCount
    'set the sMyPar string equal to the text within the specified paragraph.
    sMyPar = ActiveDocument.Paragraphs(J).Range.Text
    'sets the document text equal to the modified text in the sMyPar string.
    ActiveDocument.Paragraphs(J).Range.Text = sMyPar
Next J
