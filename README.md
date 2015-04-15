# PerfView extensions

## Reporting

Reporting - extension that extracts stats between two events.

### Usage
1. Saves stats to XLSX file.
```
  PerfView.exe UserCommand Reporting.CalculateStatistics {trace-file-name} {provider-name} {start-event} {end-event} {output-file-name}
```

2. Shows stats in the log file.
```
  PerfView.exe UserCommand Reporting.CalculateStatistics {trace-file-name} {provider-name} {start-event} {end-event}
```

### Examples
    [https://github.com/shchahrykovich/PerfViewExtensions/blob/master/Stat.bat](https://github.com/shchahrykovich/PerfViewExtensions/blob/master/Stat.bat)    