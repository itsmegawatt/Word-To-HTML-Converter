# Word-To-HTML-Converter
Python module that converts a word document into HTML

##Geting HTML from a Word Document

Just load it up and run this command:

    convert_to_html_lines_from_path(docx_path)

Making sure to replace the docx_path with the path of where your document is located. Afterwards run it, and it will return to you the HTML version of your word document.

_Example (example.docx):_

This is an example paragraph with lots of different styling added to it such as **bold** , _italics_, underline, and strikethrough. This is easily converted with the program to get it all out.

When I start a new paragraph, the word to html converter likewise starts off a new paragraph as well. But, I can also do something like to do lists for example:

List of things:
- Item one
- Item two
- Item three
- Item four

**Note:** It actually does _not_ convert **bulleted lists** or **numbered lists** yet but it can easily be modified to do that. The author of this program just hasn&#39;t had any use case for it yet, but if you would like to add that modification feel free to push a commit.

_Converts to:_

["<p>This is an example paragraph with lots of different styling added to it such as <b>bold</b>, <i>italics</i>, <span style='text-decoration: underline;'>underline</span>, and <del>strikethrough</del>. This is easily converted with the program to get it all out.</p>",

'<p>When I start a new paragraph, the word to html converter likewise starts off a new paragraph as well. But, I can also do something like to do lists for example&colon;</p>',

'<p>List of things&colon;<br />- Item one<br />- Item two<br />- Item three<br />- Item four</p>',

'<p><b>Note&colon;</b> It actually does <i>not</i> convert <b>bulleted lists </b>or <b>numbered lists </b>yet but it can easily be modified to do that. The author of this program just hasn&apos;t had any use case for it yet, but if you would like to add that modification feel free to push a commit.</p>']

##Saving HTML from a Word Document

If you want to save the returned text into a an HTML document, run this command afterwards, making sure to input the returned "HTML Lines" as the first argument:

    save_as_html(html_lines, filename)

You can make the filename anything, it will append ".html' automatically if you don't include it, but ideally you should name it like: "something.html".