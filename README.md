# officetoolkit
A Java toolkit for producing presentation documents such as spreadsheets, slides, etc.

Initially this just contains some convenient methods for quickly working with spreadsheet formatting via Apache POI.

There is additional code that may not be appropriate for this library, related to producing PDF reports, but as it depends on a sometimes-free-sometimes-not third party library, I am slightly hesitant to pull it in.

I also have some code that works with slides but which also has a free third-party dependency that is on SourceForge but not Maven Central, so I am not yet sure whether I will pull it into this library.

I intend to grow this library in the future. The problem is that existing report generators such as BIRT, Jasper, etc., all require the work to happen on the server, but many applications need to generate reports on the client side of things or don't have a client/server architecture to take advantage of. But any such work in this area is further down the road, and it is worth noting that others have abandoned such projects. Even a few utility methods may help out though.
