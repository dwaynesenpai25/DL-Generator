import pythoncom
from win32com import client as win32
import time

wdCollapseEnd = 0
wdSectionBreakNextPage = 2
wdHeaderFooterPrimary = 1
wdHeaderFooterFirstPage = 2
wdHeaderFooterEvenPages = 3

def copy_headers_footers(src_section, dest_section):
    for header_type in [wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages]:
        src_header = src_section.Headers(header_type)
        dest_header = dest_section.Headers(header_type)
        if src_header.Exists:
            dest_header.Range.FormattedText = src_header.Range.FormattedText
            dest_header.LinkToPrevious = False

    for footer_type in [wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages]:
        src_footer = src_section.Footers(footer_type)
        dest_footer = dest_section.Footers(footer_type)
        if src_footer.Exists:
            dest_footer.Range.FormattedText = src_footer.Range.FormattedText
            dest_footer.LinkToPrevious = False

def merge_docs_with_headers(files, output_path):
    pythoncom.CoInitialize()
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        master_doc = word.Documents.Add()

        for i, file_path in enumerate(files):
            src_doc = word.Documents.Open(file_path, ReadOnly=True)

            if i > 0:
                rng = master_doc.Content
                rng.Collapse(wdCollapseEnd)
                rng.InsertBreak(wdSectionBreakNextPage)

            rng = master_doc.Content
            rng.Collapse(wdCollapseEnd)
            rng.InsertFile(file_path)

            src_section = src_doc.Sections(1)
            dest_section = master_doc.Sections(master_doc.Sections.Count)

            copy_headers_footers(src_section, dest_section)

            src_doc.Close(False)
            time.sleep(0.3)

        master_doc.SaveAs(output_path)
        master_doc.Close()
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    # Replace these with your real files and output path
    files = [r"C:\Users\SPM\Desktop\ONLY SAVE FILE HERE\mail\output\doc_CAMANAVA_3_89867b.docx", r"C:\Users\SPM\Desktop\ONLY SAVE FILE HERE\mail\output\doc_CAMANAVA_4_331e12.docx"]
    output = r"merged_output.docx"
    merge_docs_with_headers(files, output)
