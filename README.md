# cs-vsto-large-doc-writer
Package provides a large report generator using on word docs using VSTO and C#

# Objective

While there have been quite a number of libraries which claims to be able to generate word document report. However, many of these libraries failed when it came to generate a very large report, one that may contains hundreds of pages or even more. This library was created to enable report generation in word document in these circumstances.

# Langauage

C# 

# Prequisites

The library was built using VS2015 Community Edition. Note that this library is based on VSTO and thus requires the availability of office 2007 for it to work. It also requires the following COM libraries to be available in the C# project's references

* Microsoft.Office.Core (Version: 2.4)
* Microsoft.Office.Interop.Excel (Version: 1.6)
* Microsoft.Office.Interop.Word (Version: 8.4)

This link below shows how to solve the COM error when uninstall vs 2007 and reinstall some other version of office and then reinstall vs 2007:

https://social.msdn.microsoft.com/Forums/vstudio/en-US/08f13e9d-895c-4102-b6d9-e327af8cf8c0/0x80029c4a-typeecantloadlibrary?forum=vsto


