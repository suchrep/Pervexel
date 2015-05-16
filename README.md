# Pervexel

XLS Reader that is small, fast and has no dependencies.

## Issues

There is almost no error-handling yet.

Formatting is not supported.  
Expect  
`string` from strings,  
`bool` from some formulas,  
`datetime` from cells with default date formats,  
`double` from everything else.

## Usage

```
using Pervexel;
...
using (var file = File.OpenRead(@"c:\stuff.xls") {

	var compound = new CompoundStream(file);
	if (!compound.Open("Workbook"))
		return;

	var workbook = new WorkbookReader(compound);
	if (!workbook.Open(worksheet))
		return;

	while (workbook.Read()) {
		Console.Write(workbook.Row + " " + workbook.Column + ": " + workbook.Value));
	}
}
...
```
