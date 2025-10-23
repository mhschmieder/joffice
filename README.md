# JOffice

The JOffice library is a Java toolkit for producing presentation documents such as spreadsheets, slides, etc.

Initially this just contains some convenient methods for quickly working with spreadsheet formatting via Apache POI, and some helper methods for producing reports in PDF format using pdfJet, along with some XLS helper methods.

As I was previously using the paid commercial version of that library (I parked it at version 5.9.0), publishing OfficeToolkit as free open source required me to switch to the most recent free open source version of pdfJet, version 5.75 from 2016. This in turn forced a re-implementation of anything using the TextColumn class, so I have replaced those references with lists of Paragraphs, which then get gathered onto a Page using the TextFrame class.

PLEASE NOTE: I do not currently have a working "at home" application to use for testing the latest updates that fix the broken pdfJet dependencies, and am not yet using PDF on my current job, so it may be a while before I have time to  verify and validate this version of the code base in full. Additionally, the Maven source for the free version of pdfJet 5.75 is via a secondary source who is hosting it on their domain (com.hynnet).
  
As pdfJet appears to have been abandoned near the start of the COVID panedmic (no updates since early 2020), and as Apache PdfBox is growing quickly, as are several of its add-on libraries that provide higher-level page formatting and tables (uch as easytable, ph-pdf-layout, pdfbox-layout, or PdfLayoutManager), I am considering a switch away from pdfJet as time allows. PdfBox Version 3 is in alpha testing as of June 2022, and appears to be on par with iText (an expensive commercial product that is a non-starter for integration with free software projects and products).

I also have some code that works with PowerPoint slides but it has a free third-party dependency that is on SourceForge but not Maven Central, so I am not yet sure whether I will pull it into this library.

I intend to grow this library in the future. The problem is that existing report generators such as BIRT, Jasper, etc., all require the work to happen on the server, but many applications need to generate reports on the client side of things or don't have a client/server architecture to take advantage of. But any such work in this area is further down the road, and it is worth noting that others have abandoned such projects. Even a few utility methods may help out though.
