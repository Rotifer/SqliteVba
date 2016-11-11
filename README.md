# SqliteVba

## VBA classes written for use in Excel to interact with SQLite

I am working with Windows colleagues and need to provide a database for storage and management. 
Decided against MS Access amd opted for SQLite for reasons that I will explain in the future.
The plan is to use Excel as a front-end for end users with SQLite in backend.
This is a work in progress. The classes are not yet tested or documented.


## Aim to write this up in a lengthy blog entry

Requires the Sqlite ODBC driver and references to the ADO and scripting runtime libraries.

Links that I have used:

1. Microsoft ADO API Reference: https://msdn.microsoft.com/en-us/library/ms807498.aspx
3. Good tips on an object oriented approach - Class Module to wrap up classic ADO call to SQL-server: http://codereview.stackexchange.com/questions/116253/class-module-to-wrap-up-classic-ado-call-to-sql-server/116254
4. Creating ADODB Parameters on the fly: http://codereview.stackexchange.com/questions/46312/creating-adodb-parameters-on-the-fly
5. How to transfer data from an ADO Recordset to Excel with automation: https://support.microsoft.com/en-gb/kb/246335
6. Good review of dictionaries: http://excelmacromastery.com/vba-dictionary/

# General Notes, Whinges, Gripes, PITAs and a few Positives

* Excel VBA is old and cranky
* Microsoft ADO is old and cranky
* VBA isn't very pleasant to work with but it beats PHP and Java (personal opinion)
* Will Microsoft ever kill off VBA?
* VBA isn't dead but it is like the undead (https://en.wikipedia.org/wiki/Undead)
* Rubberduck takes VBA into the 21st century
* Rubberduck is not working properly for me and I don't know why (Windows 10 on Mac usng Parallels Desktop)
* SQLite is really great!
* SQL is great!
* Windows is still popular
* Excel will never die
* Excel + VBA + SQLite might make a good combination, we'll see
