# officetoolkit
A Java toolkit for producing presentation documents such as spreadsheets, slides, etc.

Initially this just contains some convenient methods for quickly working with spreadsheet formatting via Apache POI, and some helper methods for producing reports in PDF format using pdfJet.

As I was previously using the commercial version of that library (I parked it at version 5.9.0), and switched tonight to the free version (stuck at 2016's 5.7.5), thereb is one class reference for TextColumn that doesn't exist, thus this code doesn't fully compile at first pass. I hope to have time very soon to find an alternate implementation of the PDF formatting to compensate for the missing functionalityn in the free version of pdfJet.

I also have some code that works with PowerPoint slides but which also has a free third-party dependency that is on SourceForge but not Maven Central, so I am not yet sure whether I will pull it into this library.

I intend to grow this library in the future. The problem is that existing report generators such as BIRT, Jasper, etc., all require the work to happen on the server, but many applications need to generate reports on the client side of things or don't have a client/server architecture to take advantage of. But any such work in this area is further down the road, and it is worth noting that others have abandoned such projects. Even a few utility methods may help out though.
