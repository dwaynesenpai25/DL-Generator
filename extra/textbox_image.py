import streamlit as st
from pathlib import Path
import win32com.client
import pythoncom
import os
import tempfile
import sys

def replace_in_text_boxes(doc, find_str, replace_with_image_path):
    # Process all shapes (which includes text boxes)
    for shape in doc.Shapes:
        if shape.TextFrame.HasText:
            text_range = shape.TextFrame.TextRange
            # Find all occurrences of the placeholder
            find = text_range.Find
            find.Text = find_str
            find.Forward = True
            find.Wrap = 1  # wdFindContinue
            
            while find.Execute():
                # Replace with image
                found_range = text_range.Duplicate
                found_range.Find.Execute(FindText=find_str)
                found_range.Text = ""
                shape = found_range.InlineShapes.AddPicture(
                    FileName=replace_with_image_path,
                    LinkToFile=False,
                    SaveWithDocument=True
                )

def replace_in_main_text(doc, find_str, replace_with_image_path):
    # Process main document text
    find = doc.Content.Find
    find.Text = find_str
    find.Forward = True
    find.Wrap = 1  # wdFindContinue
    
    while find.Execute():
        # Replace with image
        found_range = doc.Range(find.Parent.Start, find.Parent.End)
        found_range.Text = ""
        found_range.InlineShapes.AddPicture(
            FileName=replace_with_image_path,
            LinkToFile=False,
            SaveWithDocument=True
        )

def replace_text_with_image(doc_file, image_path, output_dir):
    pythoncom.CoInitialize()
    try:
        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False

        try:
            doc_path = str(doc_file.absolute())
            image_path_str = str(image_path.absolute())
            
            if not os.path.exists(doc_path):
                raise FileNotFoundError(f"Document not found: {doc_path}")
            if not os.path.exists(image_path_str):
                raise FileNotFoundError(f"Image not found: {image_path_str}")

            doc = word_app.Documents.Open(doc_path)
            
            # Replace in main document text
            replace_in_main_text(doc, "{{image}}", image_path_str)
            
            # Replace in text boxes
            replace_in_text_boxes(doc, "{{image}}", image_path_str)
            
            # Save the document
            output_path = output_dir / f"modified_{doc_file.name}"
            output_path_str = str(output_path.absolute())
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Save in the appropriate format
            file_format = 16 if doc_file.suffix.lower() == '.docx' else 0
            doc.SaveAs(output_path_str, FileFormat=file_format)
            doc.Close(SaveChanges=False)
            
            return output_path
            
        except Exception as e:
            st.error(f"Error processing document: {str(e)}")
            exc_type, exc_obj, exc_tb = sys.exc_info()
            st.error(f"Error occurred at line: {exc_tb.tb_lineno}")
            return None
        finally:
            word_app.Application.Quit()
    finally:
        pythoncom.CoUninitialize()

def main():
    st.title("DOCX Text Box Image Replacer")
    st.write("Replace {{image}} placeholders in both text and text boxes")

    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True, parents=True)

    uploaded_doc = st.file_uploader("Upload Word Document", type=["docx", "doc"])
    uploaded_image = st.file_uploader("Upload Replacement Image", type=["png", "jpg", "jpeg"])

    if uploaded_doc and uploaded_image:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir_path = Path(temp_dir)
            
            # Save with proper extensions
            doc_ext = "docx" if uploaded_doc.type.endswith("docx") else "doc"
            doc_path = temp_dir_path / f"document.{doc_ext}"
            with open(doc_path, "wb") as f:
                f.write(uploaded_doc.getbuffer())
            
            image_ext = uploaded_image.type.split('/')[-1]
            image_path = temp_dir_path / f"image.{image_ext}"
            with open(image_path, "wb") as f:
                f.write(uploaded_image.getbuffer())
            
            if st.button("Replace Placeholders"):
                with st.spinner("Processing document..."):
                    result_path = replace_text_with_image(doc_path, image_path, output_dir)
                
                if result_path and result_path.exists():
                    st.success("Successfully replaced all placeholders!")
                    with open(result_path, "rb") as f:
                        st.download_button(
                            "Download Modified Document",
                            data=f,
                            file_name=result_path.name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                else:
                    st.error("Failed to process the document")

if __name__ == "__main__":
    main()