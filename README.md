# PLOGGER - Logger for VBA.

> [!IMPORTANT]
> WARN: project is unfinished

## Examples

1. Get the default (root) logger:

```vb
Dim Root As PLogger
Set Root = PLogging.Root
' Root's level is llError by default

Root.Trace "Hi"
Root.Debug_ "Hi"
Root.Info "Hi"
Root.Warn "Hi"
Root.Error "Hi"
Root.Fatal "Hi"
```

output:

```
ERROR:Root:Hi
FATAL:Root:Hi
```

2. Add one simple stream logger:

```vb
Dim MyLogger As PLogger
Set MyLogger = PLogging.GetLogger("MyLogger")

MyLogger.Trace "Hi"
MyLogger.Debug_ "Hi"
MyLogger.Info "Hi"
MyLogger.Warn "Hi"
MyLogger.Error "Hi"
MyLogger.Fatal "Hi"
```

output:

```console
TRACE:MyLogger:Hi
DEBUG:MyLogger:Hi
INFO:MyLogger:Hi
WARN:MyLogger:Hi
ERROR:MyLogger:Hi
FATAL:MyLogger:Hi
```

3. Add two simple stream loggers:

```vb
Dim MyLogger As PLogger
Set MyLogger = PLogging.GetLogger("MyLogger")

MyLogger.SetLevel LogLevels.llError
MyLogger.Trace "Hi"
MyLogger.Debug_ "Hi"
MyLogger.Info "Hi"
MyLogger.Warn "Hi"
MyLogger.Error "Hi"
MyLogger.Fatal "Hi"

Dim Log As PLogger
Set Log = PLogging.GetLogger("Log")
Log.Trace "Hi"
Log.Debug_ "Hi"
Log.Info "Hi"
Log.Warn "Hi"
Log.Error "Hi"
Log.Fatal "Hi"
```

output:

```console
ERROR:MyLogger:Hi
FATAL:MyLogger:Hi
TRACE:Log:Hi
DEBUG:Log:Hi
INFO:Log:Hi
WARN:Log:Hi
ERROR:Log:Hi
FATAL:Log:Hi
```
