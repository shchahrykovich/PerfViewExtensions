# PerfView extensions

## Reporting

Reporting - extension that extracts stats between two events.

### Usage:
```
  PerfView.exe UserCommand Reporting.CalculateStatistics {trace-file-name} {provider-name} {start-event} {end-event} {output-file-name}
    Saves stats to XLSX file.
```

```
  PerfView.exe UserCommand Reporting.CalculateStatistics {trace-file-name} {provider-name} {start-event} {end-event}
    Shows stats in the log file.
```

### Examples:
    [https://github.com/shchahrykovich/PerfViewExtensions/blob/master/Stat.bat](https://github.com/shchahrykovich/PerfViewExtensions/blob/master/Stat.bat)    