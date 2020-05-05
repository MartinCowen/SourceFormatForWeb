# Source Format For Web
This utility was written to solve problems I was having when pasting code snippets into the WYSIWYG editor in Publii. Some of the problems in the current version of Publii (0.35.0) are

 - Each line ends up with separate code tags and line breaks. 
 - Indents get lost because tabs become one space
 - To use the prism.js plug in, it is easier to paste into the HTML view, but then you have to HTML encode the text to avoid problems with angle brackets
 - If you want code formatting as per the prism.js plug in, you need to add the code class = language tag.

# Usage

 - Paste your source code snippet into the left pane. It is automatically formatted as per the options in the right pane. 
 - Click Copy button and paste into the HTML view in Publii or your HTML editor of choice.

## Usage Notes
There is no Execute button because the formatter runs on every text change in the left hand pane, or if the Language drop down changes, or if the Auto detect or HTML checkbox states change.

You can type in the Language drop down box if Auto detect is off. Using this feature rather than  manually editing the result means that the Language is persistent during the session.

Auto detect for language is on by default and will run on every text change. It uses simple heuristics to detect which language the code is in, but all this affects is the code class tag in the right pane, so you can always change it manually. Only a few languages are supported at the moment because these are the ones I was using, but it would be simple to add them to the auto detect routine, the difficulty comes when the languages are similar in format because this is only intended for snippets, not complete programs. The more languages that are added, the harder this task becomes and it may be worth adding in proper parsers rather than the simplistic keyword & format heuristics.

The Clear button clears the input text box, and by automatic execution the output box, and sets focus in the input text box for the next paste.

HTML Encode check box controls whether angle brackets in the source are encoded or not.

Copy button is the same as Select All and Copy in the right text box pane. It will overwrite your clipboard.

The window is resizable and the pane splitter is movable.

## Source Code
Written in VB.NET as a WinForms application using .NET Framework 4.7.2 in Visual Studio 2019
No third party components are used.

