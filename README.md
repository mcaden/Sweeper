# Sweeper - VSIX
_**Coming Soon**_  
Sweeper is a tool for C# code formatting. The legacy version was deployed as an add-in for VS 2008-2013. It will soon be revamped to be packaged as a VSIX.


-----
## Legacy Add-in
**A Visual Studio Add-in for C# Code Formatting in Visual Studio 2008, 2010, 2012, and 2013**
Includes:
- A UI for options, Enable or disable any specific task you want to.
- A decent list of formatting tasks.
- The ability to call other formatting-related addins after completion of a task - currently StyleCop and GhostDoc

It installs off a single MSI which does all the work. Simply install it and it will add the appropriate commands to the "Tools" menu (May need to restart Visual Studio). It will automatically install for both VS 2008 and VS 2010 - whichever you have installed.

Please keep in mind this is an BETA. I greatly appreciate any testing you can do - especially over large codebases, but please be smart about it. Also keep in mind that running stylecop saves your files.

The only time I've ever seen broken code with my testing has been with preprocessor directives on sorting elements within classes. #region and #endregion break this functionality. I currently have a warning before it runs that allows you to skip that task on that file. It will display the error if you're using any preprocessor directives whatsoever, however if the directives are inside a function, constructor, property, or outside of any class it will not break. This problem only affects directives at the class level.

This addin is for Visual Studio 2008, 2010, 2012, and 2013 on C# projects
