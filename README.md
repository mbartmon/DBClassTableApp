DBClassTableApp
===============

Creates VB.NET or C#.NET classes from Access or SQL database tables.

This simplifies data access in projects in which Entity Framework was either not used or you simply do not want to use.'

This is very much a work in progress even though I have successfully used it in several commercial projects. As developers you know that the utility programs that you write for yourself are often incomplete, not necessarily user friendly, and contain bugs.

The intention was to write a program that would accept either MS Access or MS Sql tables as input and generate a class file for each table processed.  The resulting class file, along with the base class ### clsDb ###, provides all of the schema translation and basic data access methods for SELECT, UPDATE, INSERT, and DELETE along with <i>Properties</i> to set and get column values.

While the Access conversion allows for the selection of any Access database on your computer, the SQL conversion is hard coded for two of the databases that I am currently using in development projects. That is one area that can be generalized.

I will be working on a more complete description and some usage notes as well as inserting in-line method docs.  For now you can look at the code and puzzle over some of the aspects that may be peculiar to my project environment (although I think that may be a very small part of the whole).

So, what needs to be done?  Here is a incomplete list:

1.  As noted above the SQL user and backend code need to be generalized and offer the user a way to select a server and       database.

2.  There are two different methods used for writing VB classes: one for Access and the other for SQL.  These should be
    reduced to a single method used by both selections.
    
3.  There is no path for creating C# classes for Access. This should be remedied.

I hope you find this project interesting and usable.  I look forward to comments and requests for collaboration.
