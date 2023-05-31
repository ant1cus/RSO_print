from msoffice2pdf import convert


# import win32com
#
#
# def word2pdf(doc, pdf):
#     word = win32com.client.Dispatch('Word.Application')
#     word.Visible = False
#     wd_format_pdf_ = 17
#     doc_for_conv = word.Documents.Open(doc)
#     doc_for_conv.SaveAs(pdf, FileFormat=wd_format_pdf_)
#     doc_for_conv.Close()
#     word.Quit()


def word2pdf(doc, pdf):
    out = convert(source=doc, output_dir=pdf, soft=0)
    return out
