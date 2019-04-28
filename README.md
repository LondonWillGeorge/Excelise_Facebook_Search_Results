# Excelise FB Search Results
By pasting the results list from a people search from FB into Excel, this will reformat to a neat table in new Excel workbook.

To use this you need:

To download and install Python 3.7 or later

Excel or OpenOffice etc freeware software which can save in .xlsx format, 

Know how to install Python packages and set up virtual environment, run a Python app in console (can google this fairly easily maybe)

The source workbook needs to be filled with search results from Fb.

NB this was written asap at work (not yet software dev when wrote this), so at first, didn't even encapsulate in functions, let alone 
as a class. Just wrote the whole thing out until it worked in one ever-expanding script.

Then later I attempted to break into a few functions, and encapsulated these in one class, so though the object model is a bit makeshift
and not planned from the start, but it's potentially extendible.

# How to do searches for specific groups of people in FB?
NB you can't get these kind of search results with the FB API, it won't let you, but these are all URLs you can just put into public
website, returning publicly listed information which everyone has consented to share publicly on Fb.

The example filter code in this Python class and example search URLs here are for people working in nursing, but could be adapted 
for any profession, or any kind of fb search which gives a list of people's fb profiles as results.

Eg Below URL would be a search for all people on fb who list themselves as currently living in "Newport" and who give one of their job 
titles, in the Facebook work title field, as including the words "staff" and "nurse"

i.e. most of them will be currently working as a Staff Nurse, maybe a few with previous but not current job as Staff Nurse, however 
that is rare in practice.

Note the way of constructing the FB URL search string.
https://www.facebook.com/search/str/staff+nurse/pages-named/employees/present/str/Newport/pages-named/residents/present/intersect

ie to change to people who list themselves as "sheep farmer" use same URL as above changing /staff+nurse/ to /sheep+farmer/ ,
(apologies to the Welsh for gratuitous stereotype :-)

Or get FB code for Newport eg Newport, Wales is: 112195725462212 (correct at 27th April, 2019)
so substituting the numeric code for the string above, for a very similar set of results, use: 
https://www.facebook.com/search/str/staff+nurse/pages-named/employees/present/112195725462212/residents/present/intersect
though note there will be some differences as second presumably returns all people whose geo location listed on FB within maximum
distance X of centre of Newport, Wales, or something along these lines, whereas the string version will return anyone in world who
typed "Newport" as part of location name, so would also include Newport in Australia for example.

What proportion of people who actually work as X in certain area might you get?
My very rough try with this indicates around 10% - this is an estimate!

To copy and paste results direct from FB, easiest way is:
1 Load page(s) formed using URL rules as above
2 Keep clicking arrow down or pulling side bar down in web browser until you reach end of results!
FB results like many sites are lazy loaded with Javascript, so long result list may need keep pressing down for a while!
3 Ctrl + A to copy whole page when fully loaded
IMPORTANT: in Excel sheet, place cursor in Column B row 3 or below and just paste in results
4 First several lines will be your own login details FB page generic text etc. - delete these first
5 Lots of profile photos will appear in pasted data, can ignore these, not processed here
6 Name this Excel worksheet as something which makes sense for you
7 Create a new sheet, and repeat above with the next search string, until finished
8 Save your file

The Excel files need to be in the same folder as the main

NB to run this app from the Windows command prompt, need to activate virtual environment first
Type facebookenv\Scripts\activate at command prompt with back slashes
otherwise the virtual environment only loaded packages won't be found and won't run!

# Next Steps
Using the Selenium package for Python, you could extract this data automatically without copying and pasting, but then you have to deal
with fb's extremely sophisticated 'suspicious behaviour' detection algorithms. Fb cannot detect selenium automatically, but the fb
account you use can get temporarily blocked quite quickly if you use selenium to: do anything too fast, 
access URLs in succession which 'normal users' wouldn't access directly, message too many people you don't know...very long and complex
list probably.

Have a draft selenium python file as part of this, use only with a test fb login (create new email and login might be best), you will
get the account blocked pretty quickly and easily.
