# Excel Batch Processor

## About
EBP allows you to easily use the [EPPlus](https://github.com/JanKallman/EPPlus) library to manipulate large numbers of Excel files.
You simply write a plugin that takes an EPPlus `ExcelWorkbook` object, compile it, and place the DLL file into the appropriate plugins directory.
Then, you tell EBP where to find the files you want to work on, where to output the modified files, and EBP will apply your plugin to all your files.

## Multithreaded operation
EBP uses all available processor cores to process files in parallel.

## Writing plugins
To write a plugin for EBP:

1. Include references to EPPlus and the `PluginContracts` assembly in your project
2. Create a class that implements the `IExcelProcess` interface
3. Build the library, and place the DLL in the appropriate folder

You can watch a detailed video tutorial on creating plugins [here](https://www.youtube.com/watch?v=A-mFIh22u7w&t=1s).
