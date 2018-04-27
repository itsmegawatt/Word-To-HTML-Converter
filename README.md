# Word-To-HTML-Converter
Python module that converts a word document into HTML

Just load it up and run this command:

    convert_to_html_lines_from_path(docx_path)

Making sure to replace the docx_path with the path of where your document is located. Afterwards run it, and it will return to you the HTML version of your word document.

If you want to save the returned text into a an HTML document, run this command afterwards, making sure to input the returned "HTML Lines" as the first argument:

    save_as_html(html_lines, filename)

You can make the filename anything, it will append ".html' automatically if you don't include it, but ideally you should name it like: "something.html".
