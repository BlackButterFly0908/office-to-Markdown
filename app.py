import os
import tempfile
import re

import streamlit as st
from markitdown import MarkItDown

# Set page configuration
st.set_page_config(
    page_title="Office to Markdown",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="expanded",
)


# Helper functions
def get_file_extension(filename):
    """Extract the file extension from a filename."""
    return filename.rsplit(".", 1)[1].lower() if "." in filename else ""


def convert_file_to_markdown(file_data, filename):
    """
    Convert a file to Markdown using MarkItDown.

    Args:
        file_data: The binary content of the file
        filename: The name of the file

    Returns:
        Tuple of (markdown_content, error_message)
    """
    try:
        # Create a temporary file
        ext = get_file_extension(filename)
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp_file:
            tmp_file.write(file_data)
            tmp_file_path = tmp_file.name

        # Initialize MarkItDown and convert the file
        md = MarkItDown(enable_plugins=False)
        result = md.convert(tmp_file_path)

        # Clean up the temporary file
        os.unlink(tmp_file_path)

        return result.text_content, None
    except Exception as e:
        return "", str(e)


def get_supported_formats():
    """
    Get a dictionary of supported file formats categorized by type.

    Returns:
        Dictionary of supported formats
    """
    return {
        "📝 Documents": {
            "formats": ["Word (.docx, .doc)", "PDF", "EPub"],
            "extensions": ["docx", "doc", "pdf", "epub"],
        },
        "📊 Spreadsheets": {
            "formats": ["Excel (.xlsx, .xls)"],
            "extensions": ["xlsx", "xls"],
        },
        "📊 Presentations": {
            "formats": ["PowerPoint (.pptx, .ppt)"],
            "extensions": ["pptx", "ppt"],
        },
    }


def main():
    # Initialize session state for multiple files
    if "conversions" not in st.session_state:
        st.session_state.conversions = {}  # Dictionary to store multiple conversions

    # Header
    st.title("IRIS EMBEDDING MARKDOWN CONVERTER")


    # File upload - enable multiple files
    all_extensions = []
    formats = get_supported_formats()
    for category, info in formats.items():
        all_extensions.extend(info["extensions"])

    uploaded_files = st.file_uploader(
        "Upload files to convert",
        type=all_extensions,
        accept_multiple_files=True,
        help="Select one or more files to convert to Markdown",
    )

    # Convert button
    if st.button("Convert to Markdown", use_container_width=True, type="primary"):
        if not uploaded_files:
            st.error("Please upload at least one file")
        else:
            for uploaded_file in uploaded_files:
                with st.spinner(f"Converting {uploaded_file.name} to Markdown..."):
                    # Convert uploaded file
                    markdown_content, error = convert_file_to_markdown(
                        uploaded_file.getbuffer(), uploaded_file.name
                    )

                    if error:
                        st.error(f"Error converting {uploaded_file.name}: {error}")
                    else:
                        # Store in session state using filename as key
                        st.session_state.conversions[uploaded_file.name] = markdown_content
                        st.success(f"Converted {uploaded_file.name} successfully!")

    # Results section - display each converted file in its own container
    if st.session_state.conversions:
        st.divider()
        st.subheader("Converted Files")
        
        for file_name, markdown_content in st.session_state.conversions.items():
            with st.container():
                st.subheader(file_name)
                
                # Download button for this file
                md_file_name = file_name.rsplit(".", 1)[0] + ".md"
                st.download_button(
                    label=f"Download {md_file_name}",
                    data=markdown_content,
                    file_name=md_file_name,
                    mime="text/markdown",
                    key=f"download_{md_file_name}",
                    use_container_width=True,
                )
                
                # Preview for this file
                with st.expander("Preview"):
                    preview_content = markdown_content[:2000]
                    if len(markdown_content) > 2000:
                        preview_content += (
                            "...\n\n(Preview truncated. Download the full file to see all content.)"
                        )
                    st.text_area(label="", value=preview_content, height=300, disabled=True, key=f"preview_{md_file_name}")
                
                st.divider()


if __name__ == "__main__":
    main()
