This is a visual studio 2022 project. 
It requires .net framework 4.8 targeting pack
It uses the extension: visual studio installer project

It requires SQLite with FTS5 installed. Perform these from the package manager console:
 Install-Package Microsoft.Data.Sqlite
 Install-Package SQLitePCLRaw.bundle_e_sqlite3
 (This will add SQLite interop DLLs)

Add reference to System.Management, System.IO.Compression

Add nuget package: microsoft.web.webview2 