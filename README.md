# Fluent POI
An abstraction over the Excel POI library with a simpler fluent API.

Dependends on:
https://poi.apache.org/

You use it like this:
```
Book.create()
        .sheet("SimpleSheet")
        .row(0)
            .cell(0).bold().value("Name").end()
            .cell(1).bold().value("Job").end()
        .end()
        .value(1, 0, "Luke")
        .value(1, 1, "Coder")
        .value(2, 0, "Jane")
        .value(2, 1, "Coder")
        .done()
        .write("output/simplesheet.xlsx");
```
