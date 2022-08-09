INPUT.CSV
--------------

- This needs to contain 4 inputs:

font-size (number)
font (string)
bold (y or n)
italic (y or n)

- In the form:

font-size,font,bold,italic

- Example:

21,Calibri,y,n

- MAKE SURE THERE ARE NO STRAY WHITE SPACES.

FORMAT
------------

- All slides are changed to have a width of 33.867 cm and a height of 19.05 cm.
- All slides have their background set to black.
- All text is analyzed and then reformatted to be set on two lines. If a paragraph has too much text, it may overflow and wrap to a new line causing more than 2 lines.
- All text is centered on the SLIDE
- I did encounter a bug that caused some text to not be formatted correctly. This may occur sometimes it seems.
- Powerpoint makes a differentiation between a "textbox" and a "placholder". Due to this, it may occur that some slides won't have their text edited properly.
- All shapes (pictures, lines, etc...) are deleted from the slide.


BUGS
-----------------

I did no extensive bug testing. If you mess up the input.csv, there probably will be uncaught bugs.
So please try to make sure that your inputs are correct.
