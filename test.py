#-*- coding:big5 -*-
 

The python-docx module
Using the python-docx module we can create and modify Word documents, which have the .docx file extension. We need to install the python-docx module first by running pip install python-docx through the command prompt as seen below:

Microsoft Windows [Version 6.1.7601]
Copyright (c) 2009 Microsoft Corporation.  All rights reserved.

C:\Users\Python>pip install python-docx
Collecting python-docx
Downloading https://files.pythonhosted.org/packages/00/ed/dc8d859eb32980ccf0e5
a9b1ab3311415baf55de208777d85826a7fb0b65/python-docx-0.8.7.tar.gz (5.4MB)
100% |||||||||||||||||||||||||||||||||| 5.4MB 624kB/s
Requirement already satisfied: lxml>=2.3.2 in c:\users\python\appdata\local\prog
rams\python\python36\lib\site-packages (from python-docx) (4.2.5)
Installing collected packages: python-docx
Running setup.py install for python-docx … done
Successfully installed python-docx-0.8.7

C:\Users\Python>

We will focus on .docx files which is represented by three different data types in Python-Docx. At the highest level, a Document object represents the entire document. The Document object contains a list of Paragraph objects for the paragraphs in the document. Each of these Paragraph objects contains a list of one or more Run objects.


The text in a Word document has font, size,color, and other styling information associated with it. A style in Word is a collection of these attributes. A Run object is a contiguous run of text with the same style. A new Run object is needed whenever the text style changes. Thus a word document’s text is more than a string. 


Let’s start using the python-docx module by making a program which will create a word document. See the code below:


import docx

document = docx.Document()

document.save(‘summary.docx’)

This creates a new document from the built-in default template and saves it unchanged to a file named ‘test.docx’. The so-called “default template” is actually just a Word file having no content, stored with the installed python-docx package. This file will be created in your program directory. The next step is to add some content to our summary.docx file which can be done in a form of a paragraph.

A paragraph has a variety of properties that specify its placement within its container (typically a page) and the way it divides its content into separate lines.


In general, it’s best to define a paragraph style collecting these attributes into a meaningful group and apply the appropriate style to each paragraph, rather than repeatedly apply those properties directly to each paragraph. The formatting properties of a paragraph are accessed using the ParagraphFormat object available using the paragraph’s paragraph_format property. 


Alignment 


The horizontal alignment of a paragraph (also known as justification) can be set to left, centered, right, or fully justified (aligned on both the left and right sides) using values from the enumeration WD_PARAGRAPH_ALIGNMENT:


import docx

from docx.enum.text import WD_ALIGN_PARAGRAPH


document = docx.Document()

paragraph = document.add_paragraph()

paragraph_format = paragraph.paragraph_format

paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

print(paragraph_format.alignment)


When we run this program we’ll get the alignment of the paragraph as shown below:


CENTER (1)


——————

(program exited with code: 0)


Press any key to continue . . .


So if we want to set the alignment of paragraph in our summary.docx document we can do as shown below:


import docx

from docx.enum.text import WD_ALIGN_PARAGRAPH


document = docx.Document(‘summary.docx’)

paragraph = document.add_paragraph()

paragraph_format = paragraph.paragraph_format

paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

document.save(‘summary.docx’)


After running this program, open the summary.docx file to see that the alignment is set to center. 


A paragraph can be indented separately on the left and right side. The first line can also have a different indentation than the rest of the paragraph. A first line indented further than the rest of the paragraph has first line indent. A first line indented less has a hanging indent.

Indentation 


Indentation is specified using a Length value, such as Inches, Pt, or Cm. Negative values are valid and cause the paragraph to overlap the margin by the specified amount. A value of None indicates the indentation value is inherited from the style hierarchy. Assigning None to an indentation property removes any directly-applied indentation setting and restores inheritance from the style hierarchy. Now lets create indentation in our summary.docx file as shown in the program:


import docx

from docx.shared import Inches

document = docx.Document(‘summary.docx’)

paragraph = document.add_paragraph()

paragraph_format = paragraph.paragraph_format

paragraph_format.left_indent = Inches(0.5)

paragraph_format.right_indent = Inches(0.5)

print(paragraph_format.left_indent)

print(paragraph_format.right_indent)

print(paragraph_format.left_indent.inches)

print(paragraph_format.right_indent.inches)


Run this program and check if the indentation in our summary.docx file is set to 0.5 icnhes at both left and right ends. First-line indent is specified using the first_line_indent property and is interpreted relative to the left indent. 


The tab stops


The tab stops determines the rendering of a tab character in the text of a paragraph. In particular, it specifies the position where the text following the tab character will start, how it will be aligned to that position, and an optional leader character that will fill the horizontal space spanned by the tab.


The tab stops for a paragraph or style are contained in a TabStops object accessed using the tab_stops property on ParagraphFormat as follows:


tab_stops = paragraph_format.tab_stops


A new tab stop is added using the add_tab_stop() method:


tab_stop = tab_stops.add_tab_stop(Inches(1.5)


The alignment defaults to left, but we specify by providing a member of the WD_TAB_ALIGNMENT enumeration. The leader character defaults to spaces, but can also be specified by providing a member of the WD_TAB_LEADER enumeration. Now let us set a tab stop as shown in the program:


import docx

from docx.shared import Inches

from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER

document = docx.Document(‘summary.docx’)

paragraph = document.add_paragraph()

paragraph_format = paragraph.paragraph_format

tab_stops = paragraph_format.tab_stops

tab_stop = tab_stops.add_tab_stop(Inches(1.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

print(tab_stop.alignment)

print(tab_stop.leader)


Run this program, the value of tab stop should be as shown:


RIGHT (2)

DOTS (1)


——————

(program exited with code: 0)


Press any key to continue . . .

Paragraph spacing
The space_before and space_after properties control the spacing between subsequent paragraphs, controlling the spacing before and after a paragraph, respectively. Inter-paragraph spacing is collapsed during page layout, meaning the spacing between two paragraphs is the maximum of the space_after for the first paragraph and the space_before of the second paragraph. Paragraph spacing is specified as a Length value, often using Pt as shown in the program :
import docx

from docx.shared import Inches,Pt


document = docx.Document(‘summary.docx’)

paragraph = document.add_paragraph()

paragraph_format = paragraph.paragraph_format

paragraph_format.space_before = Pt(18)

paragraph_format.space_after = Pt(12)


print(paragraph_format.space_before)

print(paragraph_format.space_after)
Run this program and you’ll receive the following output:
228600

152400


——————

(program exited with code: 0)


Press any key to continue . . .
Line spacing
Line spacing is the distance between subsequent baselines in the lines of a paragraph. Line spacing can be specified either as an absolute distance or relative to the line height (essentially the point size of the font used). A typical absolute measure would be 18 points. A typical relative measure would be double-spaced (2.0 line heights). The default line spacing is single-spaced (1.0 line heights).
Line spacing is controlled by the interaction of the line_spacing and line_spacing_rule properties. line_spacing is either a Length value, a (small-ish) float, or None. A Length value indicates an absolute distance. A float indicates a number of line heights. None indicates line spacing is inherited. line_spacing_rule is a member of the WD_LINE_SPACING enumeration or None. The following program shows the usage of line spacing:
import docx

from docx.shared import Inches,Pt,Length
document = docx.Document(‘summary.docx’)

paragraph = document.add_paragraph()

paragraph_format = paragraph.paragraph_format
paragraph_format.line_spacing = Pt(18)

paragraph_format.space_after = Pt(12)
print(paragraph_format.line_spacing.pt)

print(paragraph_format.line_spacing_rule)

print(‘\n’)
paragraph_format.line_spacing = 1.75
print(paragraph_format.line_spacing)

print(paragraph_format.line_spacing_rule)

Run this program and you’ll receive the following output:
18.0

EXACTLY (4)


1.75

MULTIPLE (5)


——————

(program exited with code: 0)


Press any key to continue . . .

Pagination properties
Four paragraph properties, keep_together, keep_with_next, page_break_before, and widow_control control aspects of how the paragraph behaves near page boundaries.
keep_together causes the entire paragraph to appear on the same page, issuing a page break before the paragraph if it would otherwise be broken across two pages.
keep_with_next keeps a paragraph on the same page as the subsequent paragraph. This can be used, for example, to keep a section heading on the same page as the first paragraph of the section.
page_break_before causes a paragraph to be placed at the top of a new page. This could be used on a chapter heading to ensure chapters start on a new page.
widow_control breaks a page to avoid placing the first or last line of the paragraph on a separate page from the rest of the paragraph.
All four of these properties are tri-state, meaning they can take the value True, False, or None. None indicates the property value is inherited from the style hierarchy. True means “on” and False means “off”. The following program demonstrates the usage of pagination properties:
import docx

from docx.shared import Inches,Pt,Length


document = docx.Document(‘summary.docx’)

paragraph = document.add_paragraph()

paragraph_format = paragraph.paragraph_format

print(paragraph_format.keep_together)

paragraph_format.keep_with_next = True

print(paragraph_format.keep_with_next)

paragraph_format.page_break_before = False

print(paragraph_format.page_break_before)
Run this program and you’ll receive the following output:
None

True

False


——————

(program exited with code: 0)


Press any key to continue . . .

Character formatting
Character formatting is applied at the Run level. Examples include font typeface and size, bold, italic, and underline.
A Run object has a read-only font property providing access to a Font object. A run’s Font object provides properties for getting and setting the character formatting for that run.
Several examples are provided here. For a complete set of the available properties, see the Font API documentation.
The font for a run can be accessed as shown in the following program:
from docx import Document


document = Document()

run = document.add_paragraph().add_run()

font = run.font
Typeface and size are set as shown in the following program:
from docx.shared import Pt


font.name = ‘Calibri’

font.size = Pt(12)
Many font properties are tri-state, meaning they can take the values True, False, and None. True means the property is “on”, False means it is “off”. Conceptually, the None value means “inherit”. A run exists in the style inheritance hierarchy and by default inherits its character formatting from that hierarchy. Any character formatting directly applied using the Font object overrides the inherited values.
Bold and italic are tri-state properties, as are all-caps, strikethrough, superscript, and many others. See the following program:
from docx import Document


document = Document()

run = document.add_paragraph().add_run()

font = run.font

print(font.italic)

print(font.bold)


font.italic = True

print(font.italic)


font.italic = False

print(font.italic)


font.italic = None

print(font.italic)
Run this program and you’ll receive the following output:
None

None

True

False

None


——————

(program exited with code: 0)


Press any key to continue . . .
Underline is a bit of a special case. It is a hybrid of a tri-state property and an enumerated value property. True means single underline, by far the most common. False means no underline, but more often None is the right choice if no underlining is wanted. The other forms of underlining, such as double or dashed, are specified with a member of the WD_UNDERLINE enumeration. See the program below:
from docx import Document

from docx.enum.text import WD_UNDERLINE
document = Document()

run = document.add_paragraph().add_run()

font = run.font
print(font.underline)

font.underline = True

print(font.underline)

font.underline = WD_UNDERLINE.DOT_DASH

print(font.underline)
Run this program and you’ll receive the following output:
None

True

DOT_DASH (9)


——————

(program exited with code: 0)


Press any key to continue . . .
Font color
Each Font object has a ColorFormat object that provides access to its color, accessed via its read-only color property.
The following program shows how to apply a specific RGB color to a font:
from docx import Document

from docx.shared import RGBColor


document = Document()

run = document.add_paragraph().add_run()

font = run.font

font.color.rgb = RGBColor(0x42, 0x24, 0xE9)
A font can also be set to a theme color by assigning a member of the MSO_THEME_COLOR_INDEX enumeration as shown in the program below:
from docx import Document

from docx.enum.dml import MSO_THEME_COLOR


document = Document()

run = document.add_paragraph().add_run()

font = run.font

font.color.theme_color = MSO_THEME_COLOR.ACCENT_1
A font’s color can be restored to its default (inherited) value by assigning None to either the rgb or theme_color attribute of ColorFormat. See the program below:
from docx import Document

from docx.shared import RGBColor


document = Document()

run = document.add_paragraph().add_run()

font = run.font

font.color.rgb = None
Determining the color of a font begins with determining its color type as shown in the program below:
from docx import Document

from docx.shared import RGBColor


document = Document()

run = document.add_paragraph().add_run()

font = run.font

font.color.rgb = RGBColor(0x42, 0x24, 0xE9)

print(font.color.type)
Run this program and you’ll receive the following output:
RGB (1)


——————

(program exited with code: 0)


Press any key to continue . . .

The value of the type property can be a member of the MSO_COLOR_TYPE enumeration or None. MSO_COLOR_TYPE.RGB indicates it is an RGB color. MSO_COLOR_TYPE.THEME indicates a theme color. MSO_COLOR_TYPE.AUTO indicates its value is determined automatically by the application, usually set to black. (This value is relatively rare.) None indicates no color is applied and the color is inherited from the style hierarchy; this is the most common case.
When the color type is MSO_COLOR_TYPE.RGB, the rgb property will be an RGBColor value indicating the RGB color. See the code below:
from docx import Document

from docx.shared import RGBColor


document = Document()

run = document.add_paragraph().add_run()

font = run.font

font.color.rgb = RGBColor(0x42, 0x24, 0xE9)


print(font.color.rgb)

print(font.color.theme_color)
Run this program and you’ll receive the following output:


4224E9

None


——————

(program exited with code: 0)


Press any key to continue . . .

Adding paragraph to the file

Now let’s add a paragraph to our summary.docx file. See the program below:
from docx import Document

document = Document(‘summary.docx’)

document.add_paragraph(‘Many font properties are tri-state, meaning they can take the values True, False, and None. ‘)
document.save(‘summary.docx’)
Run this program and check whether the summary.docx file now has the added paragraph. We can add paragraphs by calling the add_paragraph() method again with the new paragraph’s text. Or to add text to the end of an existing paragraph, we can call the paragraph’s add_run() method and pass it a string. See the program below:

from docx import Document

document = Document(‘summary.docx’)

po1 = document.add_paragraph(‘In general, it’s best to define a paragraph style collecting these attributes into a meaningful group and apply the appropriate style to each paragraph, rather than repeatedly apply those properties directly to each paragraph. ‘)


po2 = document.add_paragraph(‘The text in a Word document has font, size,color, and other styling information associated with it. ‘)


po1.add_run(‘Added to the second Paragraph’)
po2.add_run(‘Added to the third Paragraph’)

document.save(‘summary.docx’)
Run this program and check whether the summary.docx file now has the added paragraphs.The content of summary.docx file should look like:

Many font properties are tri-state, meaning they can take the values True, False, and None. 

In general, it’s best to define a paragraph style collecting these attributes into a meaningful group and apply the appropriate style to each paragraph, rather than repeatedly apply those properties directly to each paragraph. Added to the second Paragraph

The text in a Word document has font, size,color, and other styling information associated with it. Added to the third Paragraph

Adding paragraph with headings

 We can add paragraph with a heading style with the help of add_heading() method as shown below:

from docx import Document

import docx

document = docx.Document()

document.save(‘headingdemo.docx’)

document = Document(‘headingdemo.docx’)
document.add_heading(‘Title’, 0)
document.add_heading(‘Introduction’, 1)
document.add_heading(‘Summary’, 2)
document.save(‘headingdemo.docx’)

Run this program and check whether the headingdemo.docx file is created and has the specified headings. It should look like:

Title

Introduction
Summary
The arguments to add_heading() are a string of the heading text and an integer from 0 to 4. The integer 0 makes the heading the Title style, which is used for the top of the document. Integers 1 to 4 are for various heading levels, with 1 being the main heading and 4 the lowest subheading.


The add_heading() function returns a Paragraph object to save you the step of extracting it from the Document object as a separate step.

Adding Line and Page Breaks

We can add a line break (rather than starting a whole new paragraph), by calling the add_break() method on the Run object we want to have the break appear after. To add a page break instead, we  need to pass the value docx.text.WD_BREAK.PAGE as a lone argument to add_break(). See the following program:

import docx

document = docx.Document()
document.add_paragraph(‘This is on the first page!’)
document.paragraphs[0].runs[0].add_break(docx.enum.text.WD_BREAK.PAGE)
document.add_paragraph(‘This is on the second page!’)
document.save(‘headingdemo.docx’)  

Run this program and check whether the headingdemo.docx file now contains two pages with This is on the first page! on the first page and This is on the second page! on the second.


Adding pictures

Using the add_picture() method we can add an image to the end of the document. See the program below:


import docx

document = docx.Document()
document.add_picture(‘m2.jpg’, width=docx.shared.Inches(1),height=docx.shared.Cm(4))
document.save(‘headingdemo.docx’)


Run this program and check whether the headingdemo.docx file now contains the m2.jpg image. The first argument of add_picture is a string of the image’s filename. The optional width and height keyword arguments will set the width and height of the image in the document. If left out, the width and height will default to the normal size of the image. It is better to specify an image’s height and width in familiar units such as inches and centimeters, so we used the docx.shared.Inches() and docx.shared.Cm() functions while specifying the width and height keyword arguments.


Retrieving text from a file

1. Full text


Using the getText() function we can get the full text from a file. It accepts a filename of a .docx file and returns a single string value of its text. See the program below:


import docx

def getText(filename):
document = docx.Document(filename)
doc_text = []
for para in document.paragraphs:

doc_text.append(para.text)

return ‘\n’.join(doc_text)

print(getText(‘summary.docx’))



The getText() function opens the Word document, loops over all the Paragraph objects in the paragraphs list, and then appends their text to the list in fullText. After the loop, the strings in fullText are joined together with newline characters.

Run this program and your output should contain the text from summary.docx file as shown below:


Many font properties are tri-state, meaning they can take the values True, False
, and None.
In general, it’s best to define a paragraph style collecting these attributes in
to a meaningful group and apply the appropriate style to each paragraph, rather
than repeatedly apply those properties directly to each paragraph. Added to the
second Paragraph
The text in a Word document has font, size,color, and other styling information
associated with it. Added to the third Paragraph
——————
(program exited with code: 0)

Press any key to continue . . .


2. Reading text from a paragraph

Each of these Paragraph objects has a text attribute that contains a string of the text in that paragraph (without the style information). This can be used to read the text of a particular paragraph. See the program below:


import docx

document = docx.Document(‘summary.docx’)
print(“Text from para 1 is “+’\n’)
print(document.paragraphs[0].text)
print(“\nText from para 2 is “+’\n’)
print(document.paragraphs[1].text)
print(“\nText from para 3 is “+’\n’)
print(document.paragraphs[2].text)   


Run the program and the output should be as follows:

Text from para 1 is

Many font properties are tri-state, meaning they can take the values True, False
, and None.

Text from para 2 is

In general, it’s best to define a paragraph style collecting these attributes in
to a meaningful group and apply the appropriate style to each paragraph, rather
than repeatedly apply those properties directly to each paragraph. Added to the
second Paragraph

Text from para 3 is

The text in a Word document has font, size,color, and other styling information
associated with it. Added to the third Paragraph
——————
(program exited with code: 0)

Press any key to continue . . .

This brings an end to our discussion so till we meet next keep learning Python as Python is easy to learn! 

Advertisements

REPORT THIS AD
Advertisements

REPORT THIS AD
Share this:
Twitter
Facebook
Like this:
Like Loading...
Related
Published by

vswami11
I am learner who wish to help others who intend to learn View all posts by vswami11

Posted on December 24, 2018Author vswami11
Leave a Reply
Enter your comment here...

