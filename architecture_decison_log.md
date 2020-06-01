| date | decision |
|------|----------|
| 29.05.2020 | A separate class ValidationLog will gather validation results (e.g. violation messages per validated slide). The class will support internal logging as well as a summary report feature (not yet implemented).|
| 22.05.2020 | Implement Logger as a class instead as a module, so that logger can have an initial state (e.g. for log level or output format - not yet implemented).|
| 01.03.2020 | The testframework use the debug console for output. This way it's not limited to a specific office application.|
| 01.03.2020 | The testframeworks functionality is limited. For simplicity it won't parse examples from external textfile and will use text variables from feature classes.|
| 01.03.2020 | The application will include an example driven testframework which runs examples in Gherkin style.|
| 01.03.2020 | The application should run mac os as well as on windows -> it must not contain references to Windows only libraries (e.g. vbscript). |
