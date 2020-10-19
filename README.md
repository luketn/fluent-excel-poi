# Fluent POI
This is a concept idea for abstracting the poor API of POI with a simpler Fluent API.

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

It only supports dates and strings, and only one font style - bold.

I found the POI syntax to be pretty obtuse. I guess they are dealing with a lot of legacy code to support and it is 
difficult to change a long-standing API but it could really use a refresh.

They may also be fighting with a poor XML model in XLSX although I haven't delved into that yet.
